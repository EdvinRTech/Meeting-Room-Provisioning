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

# Standard name for the SSPR exclusion group across every tenant this tool
# is used against - fixed rather than admin-typed, since the tool needs to
# find it by name on every run (Create or Edit) without asking again.
$Script:SsprGroupName = 'SSPR Users'

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

function Set-RoomPasswordPolicy {
    <#
        Single attempt at disabling password expiration on a room account
        via Graph - wrap this in Invoke-WithRetryProgress for the same
        replication-delay reason as Set-RoomPassword below. Kept as its own
        step (rather than folded silently into Set-RoomPassword, which is
        how this used to work) so a failure here is reported on its own
        line in the run log instead of being indistinguishable from a
        failure to set the password value itself - a room with its actual
        password set but still subject to normal expiration is a
        meaningfully different, and worse, outcome than the reverse, and
        the previous combined call couldn't tell them apart.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName
    )

    Update-MgUser -UserId $UserPrincipalName -PasswordPolicies DisablePasswordExpiration -ErrorAction Stop
}

function Set-RoomPassword {
    <#
        Single attempt at setting a room account's password profile via
        Graph - wrap this in Invoke-WithRetryProgress since a just-created
        room account can take a while to become visible to Graph writes.
        Does not touch PasswordPolicies - see Set-RoomPasswordPolicy above,
        called as its own separate step.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [Parameter(Mandatory)][string]$Password
    )

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

function Get-SsprGroupDisplayName {
    <#
        Returns the standard SSPR exclusion group name so the GUI can
        reference it in log/status text without duplicating the literal.
    #>
    [CmdletBinding()]
    param()

    $Script:SsprGroupName
}

function Get-SsprExclusionGroup {
    <#
        Looks up the standard "SSPR Users" group by its fixed name. Returns
        $null if it doesn't exist - callers should treat that as "nothing to
        keep in sync", not an error, since creating this group is optional
        and this lookup runs on every Create/Edit run regardless of whether
        the admin has ever set one up.
    #>
    [CmdletBinding()]
    param()

    # A couple of quick retries for the same reason as the CA policy sync
    # (see Sync-GroupExclusionAcrossConditionalAccessPolicies): Graph
    # load-balances directory queries across replicas that converge
    # independently, so a -Filter lookup can occasionally miss a group that
    # genuinely exists, especially one just created moments earlier in the
    # same run. -ConsistencyLevel eventual routes the query to Graph's
    # advanced-query backend, which is far less prone to that.
    $escapedName = $Script:SsprGroupName -replace "'", "''"
    $lookupError = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $found = @(Get-MgGroup -Filter "displayName eq '$escapedName'" -ConsistencyLevel eventual -CountVariable ssprGroupLookupCount -ErrorAction Stop)
            if ($found.Count -gt 0) { return $found[0] }
            return $null
        } catch {
            $lookupError = $_
            if ($attempt -lt 3) { Start-Sleep -Seconds 3 }
        }
    }
    throw $lookupError
}

function New-SsprDynamicExclusionGroup {
    <#
        Creates the standard "SSPR Users" dynamic group: enabled,
        Member-type (not guests), with at least one license. No room
        exclusions yet at creation time - those are appended one at a time
        by the caller via Add-RoomToSsprExclusionRule as each room is
        processed. This does NOT and cannot retarget SSPR itself: Graph has
        no API for SSPR's group scope, so an admin still has to point SSPR
        at this group by hand in the Entra admin center (Password reset >
        Properties).
    #>
    [CmdletBinding()]
    param(
        [scriptblock]$LogCallback
    )

    $rule = '(user.assignedPlans -any (assignedPlan.servicePlanId -ne "" -and assignedPlan.capabilityStatus -eq "Enabled")) and (user.userType -eq "Member") and (user.accountEnabled -eq true)'

    $group = New-MgGroup -DisplayName $Script:SsprGroupName `
        -MailEnabled:$false `
        -SecurityEnabled:$true `
        -MailNickname ($Script:SsprGroupName -replace '\s', '') `
        -GroupTypes @('DynamicMembership') `
        -MembershipRule $rule `
        -MembershipRuleProcessingState 'On'

    if ($LogCallback) { & $LogCallback "Created '$($Script:SsprGroupName)' dynamic group ($($group.Id)): licensed, active, Member-type accounts." }
    return $group
}

function Add-RoomToSsprExclusionRule {
    <#
        Single attempt at updating the SSPR group's membership rule to
        $NewRule - wrap in Invoke-WithRetryProgress for the same
        replication-delay reason as Set-RoomPassword/Add-RoomToGroup: a
        group that was itself only just created or modified can briefly
        404 on write even though it was just read successfully.

        The caller (Start-MeetingRoomProvisioning.ps1) composes $NewRule by
        appending this room's "and (user.userPrincipalName -ne '...')"
        exclusion clause to the group's locally-known current rule, and
        keeps tracking that locally across rooms in the same run rather
        than re-reading the rule from Graph between updates - a fresh read
        could land on a replica that hasn't caught up with the previous
        room's write yet, which would silently clobber it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string]$NewRule
    )

    Update-MgGroup -GroupId $GroupId -MembershipRule $NewRule -ErrorAction Stop
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

Export-ModuleMember -Function Get-TenantDomains, Get-DefaultTenantDomain, Get-ConditionalAccessExcludedGroups, New-ConditionalAccessExclusionGroup, Sync-GroupExclusionAcrossConditionalAccessPolicies, Get-RoomLicenseInfo, Set-RoomPasswordPolicy, Set-RoomPassword, Add-RoomToGroup, Test-SelfServicePasswordResetEnabled, Get-SsprGroupDisplayName, Get-SsprExclusionGroup, New-SsprDynamicExclusionGroup, Add-RoomToSsprExclusionRule
