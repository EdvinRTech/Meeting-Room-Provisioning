#requires -Version 5.1

# Modules the tool depends on. We stay on Microsoft Graph for every
# directory/group/policy operation (no AzureAD module) so there is only
# one sign-in flow to Graph plus one to Exchange Online.
$Script:RequiredModules = @(
    'ExchangeOnlineManagement',
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Identity.SignIns',
    'Microsoft.Graph.Identity.DirectoryManagement'
)

$Script:GraphScopes = @(
    'User.ReadWrite.All',
    'Group.ReadWrite.All',
    'Policy.Read.All',
    'Policy.ReadWrite.ConditionalAccess',
    'Directory.Read.All',
    'Organization.Read.All'
)

function Install-RoomProvisioningModules {
    <#
        Makes sure every module the tool needs is installed and imported.
        Returns a log array of strings so the caller (GUI or console) can
        display progress without this function knowing about the UI.
    #>
    [CmdletBinding()]
    param(
        [scriptblock]$ProgressCallback
    )

    $log = [System.Collections.Generic.List[string]]::new()
    $write = {
        param($msg)
        $log.Add($msg)
        if ($ProgressCallback) { & $ProgressCallback $msg }
    }

    foreach ($module in $Script:RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            & $write "Installing module $module ..."
            try {
                Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                & $write "Installed $module."
            } catch {
                & $write "FAILED to install $module`: $($_.Exception.Message)"
                throw
            }
        } else {
            & $write "$module already installed."
        }

        try {
            Import-Module -Name $module -ErrorAction Stop
        } catch {
            & $write "FAILED to import $module`: $($_.Exception.Message)"
            throw
        }
    }

    return $log
}

function Connect-RoomProvisioningServices {
    <#
        Signs in once to Exchange Online and once to Microsoft Graph, using
        the same Graph scope set the original script's password-policy step
        relied on, extended with the scopes the group/policy features need.
        Both connections are kept open for the lifetime of the GUI so we are
        not repeatedly disconnecting/reconnecting between features.
    #>
    [CmdletBinding()]
    param()

    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Connect-MgGraph -Scopes $Script:GraphScopes -NoWelcome -ErrorAction Stop

    $context = Get-MgContext
    if (-not $context) {
        throw "Connected to Exchange Online, but Microsoft Graph connection could not be verified."
    }
}

function Disconnect-RoomProvisioningServices {
    [CmdletBinding()]
    param()

    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
}

Export-ModuleMember -Function Install-RoomProvisioningModules, Connect-RoomProvisioningServices, Disconnect-RoomProvisioningServices
