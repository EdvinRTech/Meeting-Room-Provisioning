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

function Get-TenantDomains {
    <#
        Returns every verified domain in the tenant so the GUI can offer a
        picker instead of a hardcoded domain name.
    #>
    [CmdletBinding()]
    param()

    Get-MgDomain -All | Where-Object { $_.IsVerified } | Select-Object -ExpandProperty Id | Sort-Object
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

    $log = { param($msg) if ($LogCallback) { & $LogCallback $msg } }

    $group = New-MgGroup -DisplayName $DisplayName `
        -MailEnabled:$false `
        -SecurityEnabled:$true `
        -MailNickname ($DisplayName -replace '\s', '')
    & $log "Created group '$DisplayName' ($($group.Id))."

    $policies = Get-MgIdentityConditionalAccessPolicy -All
    foreach ($policy in $policies) {
        $excludeGroups = @($policy.Conditions.Users.ExcludeGroups)
        if ($excludeGroups -contains $group.Id) { continue }

        $updatedUsers = $policy.Conditions.Users
        $updatedUsers.ExcludeGroups = @($excludeGroups + $group.Id)

        try {
            Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id `
                -Conditions @{ Users = $updatedUsers; Applications = $policy.Conditions.Applications } `
                -ErrorAction Stop
            & $log "Excluded '$DisplayName' from CA policy '$($policy.DisplayName)'."
        } catch {
            & $log "FAILED to exclude '$DisplayName' from CA policy '$($policy.DisplayName)': $($_.Exception.Message)"
        }
    }

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

Export-ModuleMember -Function Get-TenantDomains, Get-ConditionalAccessExcludedGroups, New-ConditionalAccessExclusionGroup, Get-RoomLicenseInfo, Set-RoomPassword, Add-RoomToGroup
