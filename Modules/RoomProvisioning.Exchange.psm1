#requires -Version 5.1

function Get-ExistingRoomLists {
    <#
        Returns every Room List distribution group in the tenant, so the
        GUI can offer "use an existing one" instead of a hardcoded address.
    #>
    [CmdletBinding()]
    param()

    Get-DistributionGroup -ResultSize Unlimited -RecipientTypeDetails RoomList -ErrorAction Stop |
        Select-Object Name, PrimarySmtpAddress, Identity
}

function New-RoomList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$PrimarySmtpAddress
    )

    New-DistributionGroup -Name $Name -RoomList -PrimarySmtpAddress $PrimarySmtpAddress -ErrorAction Stop
}

function New-RoomMailboxIfMissing {
    <#
        Creates a room mailbox if it doesn't already exist. Returns an
        object describing whether it was created or already present, so
        the GUI can log it either way without duplicating that check.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EmailAddress,
        [Parameter(Mandatory)][string]$Password
    )

    $existing = Get-Mailbox -Identity $EmailAddress -ErrorAction SilentlyContinue
    if ($existing) {
        return [pscustomobject]@{ EmailAddress = $EmailAddress; Created = $false; Error = $null }
    }

    $alias = $EmailAddress.Split('@')[0]
    try {
        New-Mailbox -MicrosoftOnlineServicesID $EmailAddress `
            -Name $alias `
            -Alias $alias `
            -Room `
            -EnableRoomMailboxAccount $true `
            -RoomMailboxPassword (ConvertTo-SecureString -String $Password -AsPlainText -Force) `
            -ErrorAction Stop | Out-Null
        return [pscustomobject]@{ EmailAddress = $EmailAddress; Created = $true; Error = $null }
    } catch {
        return [pscustomobject]@{ EmailAddress = $EmailAddress; Created = $false; Error = $_.Exception.Message }
    }
}

function Add-RoomToRoomList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RoomListIdentity,
        [Parameter(Mandatory)][string]$RoomEmailAddress
    )

    Add-DistributionGroupMember -Identity $RoomListIdentity -Member $RoomEmailAddress -ErrorAction Stop
}

function Set-RoomPlaceInfo {
    <#
        Wraps Set-Place. Any property in $PlaceInfo that is $null or an
        empty string is left out of the call entirely, rather than being
        passed through as blank - so the room only gets the fields the
        user actually filled in. Capacity is numeric, so 0 counts as "not
        provided" here too (a real room has at least 1 seat).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][hashtable]$PlaceInfo
    )

    $params = @{ Identity = $Identity; ErrorAction = 'Stop' }
    foreach ($key in $PlaceInfo.Keys) {
        $value = $PlaceInfo[$key]
        $isBlank = ($null -eq $value) -or ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) -or ($value -is [int] -and $value -eq 0)
        if (-not $isBlank) {
            $params[$key] = $value
        }
    }

    Set-Place @params
}

Export-ModuleMember -Function Get-ExistingRoomLists, New-RoomList, New-RoomMailboxIfMissing, Add-RoomToRoomList, Set-RoomPlaceInfo
