#requires -Version 5.1

# A handful of common Microsoft 365 SKUs mapped to a friendly name, purely
# for display in the license-check step. Anything not in this table just
# falls back to showing its raw SkuPartNumber, which is still meaningful to
# an admin even if less pretty.
$Script:FriendlySkuNames = @{
    'MCOMEETADV'                 = 'Microsoft Teams Rooms Basic'
    'Microsoft_Teams_Rooms_Pro'  = 'Microsoft Teams Rooms Pro'
    'MEETING_ROOM'               = 'Microsoft Teams Rooms Standard'
    'ENTERPRISEPACK'             = 'Office 365 E3'
    'ENTERPRISEPREMIUM'          = 'Office 365 E5'
    'SPE_E3'                     = 'Microsoft 365 E3'
    'SPE_E5'                     = 'Microsoft 365 E5'
    'O365_BUSINESS_PREMIUM'      = 'Microsoft 365 Business Standard'
    'SPB'                        = 'Microsoft 365 Business Premium'
    'STANDARDPACK'               = 'Office 365 E1'
}

# Marks a room account so a dynamic group's membership rule can exclude it
# (e.g. from Self-Service Password Reset scope). extensionAttribute1 is a
# legacy on-prem-AD-schema name but is fully writable/readable via Graph for
# pure cloud-only objects too - verified live against a cloud-only test user
# (see the tool's README, SSPR exclusion section) before this was built on.
$Script:SsprExclusionMarkerAttribute = 'extensionAttribute1'
$Script:SsprExclusionMarkerValue     = 'MeetingRoomProvisioningTool'

function Get-TenantDomains {
    <#
        Returns every verified domain in the tenant so the GUI can offer a
        picker instead of a hardcoded domain name.
    #>
    [CmdletBinding()]
    param()

    Get-MgDomain -All | Where-Object { $_.IsVerified } | Select-Object -ExpandProperty Id | Sort-Object
}

function Get-DefaultTenantDomain {
    <#
        Returns the tenant's default domain (the one marked IsDefault by
        Graph - normally the <tenant>.onmicrosoft.com domain unless a
        custom domain was made default), so the GUI can build a Room
        List's email address automatically instead of asking for one.
    #>
    [CmdletBinding()]
    param()

    Get-MgDomain -All | Where-Object { $_.IsDefault } | Select-Object -First 1 -ExpandProperty Id
}

function Get-ConditionalAccessExcludedGroups {
    <#
        Scans every Conditional Access policy and returns the distinct set
        of groups excluded from at least one policy, so the GUI can offer
        them as "reuse this group" candidates instead of always creating a
        new "Meeting Rooms" group.
    #>
    [CmdletBinding()]
    param()

    $policies = Get-MgIdentityConditionalAccessPolicy -All
    $groupIds = $policies |
        ForEach-Object { $_.Conditions.Users.ExcludeGroups } |
        Where-Object { $_ } |
        Select-Object -Unique

    foreach ($groupId in $groupIds) {
        try {
            Get-MgGroup -GroupId $groupId -ErrorAction Stop | Select-Object Id, DisplayName
        } catch {
            # Group referenced by a policy but no longer exists / not resolvable - skip it.
        }
    }
}

function Sync-GroupExclusionAcrossConditionalAccessPolicies {
    <#
        Makes sure $GroupId is excluded from EVERY Conditional Access
        policy that exists right now - including ones created after the
        group itself was set up. Call this on every provisioning run (not
        just once, at group-creation time) so the exclusion stays
        consistent as new CA policies get added over time instead of
        silently drifting out of sync.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string]$GroupDisplayName,
        [scriptblock]$LogCallback
    )

    # See README "Why every callback uses GetNewClosure" for why this
    # wrapper needs GetNewClosure() - without it, $LogCallback can resolve
    # incorrectly once invoked from a different function's scope.
    $emit = {
        param($msg)
        if ($LogCallback) { & $LogCallback $msg }
    }.GetNewClosure()

    $policies = Get-MgIdentityConditionalAccessPolicy -All
    $alreadyExcludedCount = 0
    $newlyExcludedCount = 0

    foreach ($policy in $policies) {
        $excludeGroups = @($policy.Conditions.Users.ExcludeGroups)
        if ($excludeGroups -contains $GroupId) {
            $alreadyExcludedCount++
            continue
        }

        $updatedUsers = $policy.Conditions.Users
        $updatedUsers.ExcludeGroups = @($excludeGroups + $GroupId)

        # A policy that was itself only just created (e.g. by another admin
        # moments ago) can briefly 404 on write even though it just showed up
        # in the -All listing above - same directory-replication lag the
        # room-provisioning retries elsewhere guard against. A few quick
        # local retries are enough; this is a short backend sync step with no
        # per-item UI progress, so it doesn't need the full cancellable
        # Invoke-WithRetryProgress machinery.
        $updateSucceeded = $false
        $lastError = $null
        for ($attempt = 1; $attempt -le 5; $attempt++) {
            try {
                Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id `
                    -Conditions @{ Users = $updatedUsers; Applications = $policy.Conditions.Applications } `
                    -ErrorAction Stop
                $updateSucceeded = $true
                break
            } catch {
                $lastError = $_
                if ($attempt -lt 5) { Start-Sleep -Seconds 3 }
            }
        }

        if ($updateSucceeded) {
            & $emit "Excluded '$GroupDisplayName' from CA policy '$($policy.DisplayName)' (new since the group was last synced)."
            $newlyExcludedCount++
        } else {
            & $emit "FAILED to exclude '$GroupDisplayName' from CA policy '$($policy.DisplayName)': $($lastError.Exception.Message)"
        }
    }

    if ($newlyExcludedCount -eq 0) {
        & $emit "'$GroupDisplayName' was already excluded from all $alreadyExcludedCount existing CA polic$(if ($alreadyExcludedCount -eq 1) { 'y' } else { 'ies' })."
    }
}

function New-ConditionalAccessExclusionGroup {
    <#
        Creates a new security group and excludes it from every existing
        Conditional Access policy in the tenant. Returns the created group.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [scriptblock]$LogCallback
    )

    $group = New-MgGroup -DisplayName $DisplayName `
        -MailEnabled:$false `
        -SecurityEnabled:$true `
        -MailNickname ($DisplayName -replace '\s', '')
    if ($LogCallback) { & $LogCallback "Created group '$DisplayName' ($($group.Id))." }

    Sync-GroupExclusionAcrossConditionalAccessPolicies -GroupId $group.Id -GroupDisplayName $DisplayName -LogCallback $LogCallback

    return $group
}

function Get-RoomLicenseInfo {
    <#
        Reports whether a room mailbox currently has any license assigned.
        Informational only - this tool never assigns a license itself.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName
    )

    try {
        $licenseDetails = Get-MgUserLicenseDetail -UserId $UserPrincipalName -ErrorAction Stop
    } catch {
        return [pscustomobject]@{ UserPrincipalName = $UserPrincipalName; HasLicense = $false; Licenses = @() }
    }

    $names = $licenseDetails | ForEach-Object {
        if ($Script:FriendlySkuNames.ContainsKey($_.SkuPartNumber)) {
            $Script:FriendlySkuNames[$_.SkuPartNumber]
        } else {
            $_.SkuPartNumber
        }
    }

    [pscustomobject]@{
        UserPrincipalName = $UserPrincipalName
        HasLicense        = [bool]$names
        Licenses          = @($names)
    }
}

function Set-RoomPassword {
    <#
        Single attempt at setting a room account's password profile via
        Graph - wrap this in Invoke-WithRetryProgress since a just-created
        room account can take a while to become visible to Graph writes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [Parameter(Mandatory)][string]$Password
    )

    Update-MgUser -UserId $UserPrincipalName -PasswordPolicies DisablePasswordExpiration -ErrorAction Stop
    Update-MgUser -UserId $UserPrincipalName -PasswordProfile @{
        ForceChangePasswordNextSignIn = $false
        Password                      = $Password
    } -ErrorAction Stop
}

function Test-SelfServicePasswordResetEnabled {
    <#
        Reports whether SSPR is enabled tenant-wide. This is the only SSPR
        setting Graph exposes at all - there is no API surface to read or
        set SSPR's group scoping (All users vs. Selected groups, or which
        group), confirmed by testing the stable and beta Graph SDKs plus raw
        REST calls. That's a genuine Microsoft platform gap, not something
        this tool works around - see README "SSPR exclusion" for what that
        means for this feature.
    #>
    [CmdletBinding()]
    param()

    (Get-MgPolicyAuthorizationPolicy -ErrorAction Stop).AllowedToUseSspr
}

function Set-RoomSsprExclusionMarker {
    <#
        Tags a room account with a fixed marker so any dynamic group whose
        membership rule excludes on it (see New-SsprDynamicExclusionGroup /
        Sync-RoomExclusionOnSsprGroup) leaves this room out. Safe to call
        even if no such group exists yet in this tenant - it's an inert tag
        until something actually checks it, which means rooms stay correctly
        excluded even if the SSPR-targeted group is created in a later run.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName
    )

    Update-MgUser -UserId $UserPrincipalName -OnPremisesExtensionAttributes @{
        $Script:SsprExclusionMarkerAttribute = $Script:SsprExclusionMarkerValue
    } -ErrorAction Stop
}

function New-SsprDynamicExclusionGroup {
    <#
        Creates a dynamic security group of "real" user accounts - enabled,
        Member-type (not guests), with at least one license - while
        excluding anything tagged by Set-RoomSsprExclusionMarker. This does
        NOT and cannot retarget SSPR itself: Graph has no API for SSPR's
        group scope, so an admin still has to point SSPR at this group by
        hand in the Entra admin center (Password reset > Properties).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [scriptblock]$LogCallback
    )

    $rule = '(user.accountEnabled -eq true) and (user.userType -eq "Member") and (user.assignedPlans -any (assignedPlan.capabilityStatus -eq "Enabled")) and not (user.{0} -eq "{1}")' -f $Script:SsprExclusionMarkerAttribute, $Script:SsprExclusionMarkerValue

    $group = New-MgGroup -DisplayName $DisplayName `
        -MailEnabled:$false `
        -SecurityEnabled:$true `
        -MailNickname ($DisplayName -replace '\s', '') `
        -GroupTypes @('DynamicMembership') `
        -MembershipRule $rule `
        -MembershipRuleProcessingState 'On'

    if ($LogCallback) { & $LogCallback "Created dynamic group '$DisplayName' ($($group.Id)): licensed, active, Member-type accounts, excluding meeting rooms tagged by this tool." }
    return $group
}

function Sync-RoomExclusionOnSsprGroup {
    <#
        Makes an EXISTING group (one the admin says is already targeted by
        SSPR) exclude meeting rooms tagged by Set-RoomSsprExclusionMarker:
        - Dynamic group: appends the exclusion clause to its membership rule
          if not already present, leaving the rest of the rule untouched.
        - Assigned (static) group: no changes - rooms are never added to a
          static group automatically, so there's nothing to exclude.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$GroupName,
        [scriptblock]$LogCallback
    )

    $emit = {
        param($msg)
        if ($LogCallback) { & $LogCallback $msg }
    }.GetNewClosure()

    # A couple of quick retries here for the same reason as the CA policy
    # sync (see Sync-GroupExclusionAcrossConditionalAccessPolicies): Graph
    # load-balances directory queries across replicas that converge
    # independently, so a -Filter lookup can occasionally miss a group that
    # genuinely exists, especially one just touched moments earlier.
    $escapedName = $GroupName -replace "'", "''"
    $foundGroups = @()
    $lookupError = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            # -ConsistencyLevel eventual routes this filter to Graph's
            # advanced-query backend, which is far less prone to the
            # replica-lag misses plain -Filter lookups showed under testing
            # (see README "SSPR exclusion" for the empirical writeup).
            $foundGroups = @(Get-MgGroup -Filter "displayName eq '$escapedName'" -ConsistencyLevel eventual -CountVariable groupLookupCount -ErrorAction Stop)
            if ($foundGroups.Count -gt 0) { break }
        } catch {
            $lookupError = $_
        }
        if ($attempt -lt 3) { Start-Sleep -Seconds 3 }
    }
    if ($foundGroups.Count -eq 0) {
        if ($lookupError) { throw $lookupError }
        throw "No group named '$GroupName' was found."
    }
    if ($foundGroups.Count -gt 1) { throw "More than one group is named '$GroupName' - rename one of them or use a unique name." }
    $group = $foundGroups[0]

    if ($group.GroupTypes -notcontains 'DynamicMembership') {
        & $emit "'$($group.DisplayName)' is an assigned (not dynamic) group - no changes made, since rooms are never added to it automatically."
        return $group
    }

    $exclusionClause = 'not (user.{0} -eq "{1}")' -f $Script:SsprExclusionMarkerAttribute, $Script:SsprExclusionMarkerValue
    if ($group.MembershipRule -like "*$exclusionClause*") {
        & $emit "'$($group.DisplayName)' already excludes meeting rooms tagged by this tool - no changes needed."
        return $group
    }

    $newRule = "($($group.MembershipRule)) and $exclusionClause"

    # Same replication-lag retry as the CA policy sync above - a group that
    # was itself only just created/modified can briefly 404 on write even
    # though it was just read successfully.
    $updateSucceeded = $false
    $lastError = $null
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            Update-MgGroup -GroupId $group.Id -MembershipRule $newRule -ErrorAction Stop
            $updateSucceeded = $true
            break
        } catch {
            $lastError = $_
            if ($attempt -lt 5) { Start-Sleep -Seconds 3 }
        }
    }
    if (-not $updateSucceeded) { throw $lastError }

    & $emit "Updated '$($group.DisplayName)' dynamic membership rule to also exclude meeting rooms tagged by this tool (existing rule logic left unchanged)."
    return $group
}

function Add-RoomToGroup {
    <#
        Single attempt at adding a room account to a group - wrap in
        Invoke-WithRetryProgress for the same replication-delay reason as
        Set-RoomPassword.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string]$UserPrincipalName
    )

    $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop

    $existingMember = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop |
        Where-Object { $_.Id -eq $user.Id }
    if ($existingMember) { return }

    New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $user.Id -ErrorAction Stop
}

Export-ModuleMember -Function Get-TenantDomains, Get-DefaultTenantDomain, Get-ConditionalAccessExcludedGroups, New-ConditionalAccessExclusionGroup, Sync-GroupExclusionAcrossConditionalAccessPolicies, Get-RoomLicenseInfo, Set-RoomPassword, Add-RoomToGroup, Test-SelfServicePasswordResetEnabled, Set-RoomSsprExclusionMarker, New-SsprDynamicExclusionGroup, Sync-RoomExclusionOnSsprGroup
