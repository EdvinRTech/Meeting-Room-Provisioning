#requires -Version 5.1

# Modules the tool depends on. We stay on Microsoft Graph for every
# directory/group/policy operation (no AzureAD module) so there is only
# one sign-in flow to Graph plus one to Exchange Online.
#
# The Microsoft.Graph.* submodules below are versioned as a wave - each
# release ships every submodule at (usually) the same version number, and
# they are only guaranteed to work together when they match. Mixing
# versions, or having more than one version of a submodule physically
# installed, is the single most common cause of assembly-loading errors
# like "Could not load file or assembly 'Azure.Core, Version=x.x.x.x'"
# because .NET cannot load two different versions of the same
# strong-named assembly into one process. So instead of a plain
# Install-Module, Install-RoomProvisioningModules below force-removes any
# existing installs of these modules and reinstalls them all pinned to
# one version that's confirmed to exist for every one of them.
$Script:GraphSubModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Identity.SignIns',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Applications'   # Get-MgServicePrincipal, used by Test-RequiredGraphScopesGranted
)

# Fixed rather than looked up live: this used to query Find-Module for
# each of the 5 submodules above and pick the newest version published
# for all of them (Get-MatchedGraphModuleVersion, since removed) - that
# added several PSGallery round-trips' worth of delay to every Connect.
# 2.40.0 is confirmed available for every submodule as of this writing.
# Bump this by hand if a future submodule release ever becomes required
# (e.g. for a security fix) - there's no automatic re-check anymore.
$Script:GraphModuleVersion = '2.40.0'

# ExchangeOnlineManagement is a separate product with its own release
# cadence - it is not part of the Graph version-matching above. Pinned to
# 3.6.0 specifically: newer versions default their interactive sign-in to
# the Windows account broker (WAM), which can silently pick the Windows
# account already signed in on the PC instead of prompting for the admin
# account being typed in, causing sign-in to fail against the wrong
# account entirely (surfaces as AADSTS500014 if that PC account isn't
# licensed/enabled for Exchange Online). 3.6.0 predates that default.
$Script:ExchangeOnlineManagementVersion = '3.6.0'
$Script:RequiredModules = @('ExchangeOnlineManagement') + $Script:GraphSubModules

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
        Guarantees a clean, matched set of required modules regardless of
        whatever was already on the machine: disconnects any existing
        Exchange Online / Graph sessions first (an active connection can
        hold the module's files open, which is the most common reason
        Uninstall-Module fails with "module is in use"), then force-
        removes every existing install of each required module (both
        CurrentUser and AllUsers scope - Start-MeetingRoomProvisioning.ps1
        always relaunches itself elevated before calling this,
        specifically so AllUsers-scope removal can succeed), then installs
        and imports the fixed, known-good version set ($Script:
        ExchangeOnlineManagementVersion / $Script:GraphModuleVersion).
        This is intentionally invasive (it will remove other versions of
        these modules that other scripts on the machine might be using) -
        it trades that for the tool reliably working the same way on any
        machine, instead of failing with hard-to-diagnose assembly-version
        errors depending on whatever happened to be installed already.

        Returns a log array of strings so the caller (GUI or console) can
        display progress without this function knowing about the UI.
    #>
    [CmdletBinding()]
    param(
        [scriptblock]$ProgressCallback
    )

    # $write always records the full detail to $log (returned to the
    # caller for troubleshooting), but only pushes a message to the GUI's
    # single-line status text when $progress is also given - this used to
    # push every per-module line live, including a failed Uninstall-Module
    # call's full, often multi-sentence .NET exception text (e.g. "Could
    # not remove Microsoft.Graph.Users 2.40.0: ..."), which is far more
    # detail than a "here's what's happening" status line needs. The full
    # detail is still in $log, and a genuine failure still throws the real
    # exception (see the catch blocks below), so nothing is actually lost -
    # it's just not flashed past the user one module at a time.
    $log = [System.Collections.Generic.List[string]]::new()
    $write = {
        param($detail, [string]$progress)
        $log.Add($detail)
        if ($ProgressCallback -and $progress) { & $ProgressCallback $progress }
    }.GetNewClosure()

    & $write 'Disconnecting any existing Exchange Online / Microsoft Graph sessions...' 'Disconnecting existing sessions...'
    Disconnect-RoomProvisioningServices

    # Start-MeetingRoomProvisioning.ps1 always relaunches itself elevated
    # before this ever runs, specifically so this step can succeed:
    # removing an AllUsers-scope install needs admin rights. Without that,
    # an old/wrong version left on disk could still get loaded instead of
    # the matched one this function installs below - see README "Module
    # installation: matched versions, elevated, clean every run".
    & $write 'Removing any existing installs of required modules for a clean, matched set (this can take a few minutes)...' 'Removing old module versions (this can take a few minutes)...'
    foreach ($module in $Script:RequiredModules) {
        Get-Module -Name $module -ErrorAction SilentlyContinue | Remove-Module -Force -ErrorAction SilentlyContinue

        $installed = @(Get-InstalledModule -Name $module -AllVersions -ErrorAction SilentlyContinue)
        foreach ($installedVersion in $installed) {
            try {
                Uninstall-Module -Name $module -RequiredVersion $installedVersion.Version -Force -ErrorAction Stop
                & $write "Removed existing $module $($installedVersion.Version)."
            } catch {
                & $write "Could not remove $module $($installedVersion.Version): $($_.Exception.Message)"
            }
        }
    }

    $graphVersion = $Script:GraphModuleVersion
    & $write "Using Microsoft.Graph module version $graphVersion (fixed - not queried live from PSGallery, to avoid the extra delay that added)."

    # Installed-to-CurrentUser modules are imported by their authoritative
    # InstalledLocation (queried right back from Get-InstalledModule) and
    # not by name+version search. A name+version Import-Module still walks
    # $env:PSModulePath, and a Graph submodule's manifest can separately
    # trigger an internal, unpinned load of another submodule by name -
    # if an AllUsers copy this tool couldn't remove is found along the
    # way, .NET can end up with two incompatible builds of the same type
    # loaded at once. Importing by exact file path removes that ambiguity
    # for our own top-level imports entirely.
    try {
        & $write "Installing ExchangeOnlineManagement $Script:ExchangeOnlineManagementVersion..." 'Installing modules (this can take a while the first time)...'
        Install-Module -Name ExchangeOnlineManagement -RequiredVersion $Script:ExchangeOnlineManagementVersion -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        $exoInfo = Get-InstalledModule -Name ExchangeOnlineManagement -RequiredVersion $Script:ExchangeOnlineManagementVersion -ErrorAction Stop
        & $write "Installed ExchangeOnlineManagement $($exoInfo.Version)."

        $graphModuleInfo = @{}
        foreach ($module in $Script:GraphSubModules) {
            & $write "Installing $module $graphVersion..."
            Install-Module -Name $module -RequiredVersion $graphVersion -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            $graphModuleInfo[$module] = Get-InstalledModule -Name $module -RequiredVersion $graphVersion -ErrorAction Stop
            & $write "Installed $module $graphVersion."
        }
    } catch {
        & $write "FAILED to install required modules: $($_.Exception.Message)" 'Module install failed.'
        throw
    }

    try {
        & $write 'Importing modules...' 'Importing modules...'
        Import-Module -Name (Join-Path $exoInfo.InstalledLocation 'ExchangeOnlineManagement.psd1') -Force -ErrorAction Stop
        foreach ($module in $Script:GraphSubModules) {
            $manifestPath = Join-Path $graphModuleInfo[$module].InstalledLocation "$module.psd1"
            Import-Module -Name $manifestPath -Force -ErrorAction Stop
        }
        & $write 'All modules installed and imported at a matched, known-good version set.'
    } catch {
        & $write "FAILED to import required modules: $($_.Exception.Message)" 'Module import failed.'
        throw
    }

    return $log
}

if (-not ('RoomProvisioning.ConsoleWindow' -as [type])) {
    Add-Type -Namespace RoomProvisioning -Name ConsoleWindow -MemberDefinition '
        [DllImport("kernel32.dll")] public static extern System.IntPtr GetConsoleWindow();
        [DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
    '
}

# Only true when Start-MeetingRoomProvisioning.ps1 relaunched itself to
# get this console (see -RelaunchedForGui there) - never touches a
# console the user opened themselves. Both Connect-ExchangeOnline and
# Connect-MgGraph sign in via their own native popup window (WAM), not
# anything printed to the console, so once hidden this console has
# nothing to show and just stays hidden for the rest of the run.
$Script:OwnsConsoleWindow = $false

function Initialize-ConsoleVisibilityControl {
    [CmdletBinding()]
    param([switch]$Enabled)
    $Script:OwnsConsoleWindow = [bool]$Enabled
    if ($Script:OwnsConsoleWindow) { Hide-RoomProvisioningConsole }
}

function Hide-RoomProvisioningConsole {
    if (-not $Script:OwnsConsoleWindow) { return }
    $hwnd = [RoomProvisioning.ConsoleWindow]::GetConsoleWindow()
    [RoomProvisioning.ConsoleWindow]::ShowWindow($hwnd, 0) | Out-Null # SW_HIDE
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

    # Default interactive sign-in (Windows account broker / WAM), same as
    # Connect-ExchangeOnline above - shows its own native sign-in window,
    # not a console prompt. This previously used -UseDeviceCode to work
    # around a Windows-PowerShell-5.1-specific assembly bug (see the
    # relaunch-to-pwsh logic in Start-MeetingRoomProvisioning.ps1); now
    # that this tool always runs under PowerShell 7, that bug doesn't
    # apply and the normal WAM sign-in window works fine.
    #
    # -ContextScope CurrentUser matters a lot here specifically: Connect-
    # MgGraph defaults to -ContextScope Process, meaning its signed-in
    # token cache is only valid for the current process. This script
    # relaunches itself into a brand-new process on every single launch
    # (for elevation/STA/PowerShell 7 - see the top of Start-
    # MeetingRoomProvisioning.ps1), so with the default scope, every
    # launch throws away any previous sign-in and forces a full
    # interactive sign-in (including MFA) from scratch, even seconds
    # after the last one. CurrentUser persists the token cache to disk,
    # so a still-valid sign-in from a previous launch is reused silently
    # - no popup, no MFA prompt - and only a genuinely expired session
    # requires signing in again.
    Connect-MgGraph -Scopes $Script:GraphScopes -NoWelcome -ContextScope CurrentUser -ErrorAction Stop

    $context = Get-MgContext
    if (-not $context) {
        throw "Connected to Exchange Online, but Microsoft Graph connection could not be verified."
    }

    # Checks the REAL server-side consent grant, not just what
    # (Get-MgContext).Scopes reports - that list can show a scope as
    # "requested" even when the actual admin-consent record behind it was
    # never fully established, which is exactly what caused a days-long
    # debugging saga: Update-MgUser -PasswordProfile kept failing with
    # Authorization_RequestDenied despite Global Administrator being
    # active and User.ReadWrite.All appearing "granted" - the interactive
    # consent prompt for this tool's full scope set had apparently never
    # actually completed properly in that tenant before. This turns that
    # into an immediate, clear failure right here instead of a mysterious
    # one during whatever step happens to need the missing permission.
    $consentCheck = Test-RequiredGraphScopesGranted -RequiredScopes $Script:GraphScopes
    if (-not $consentCheck.AllGranted) {
        throw @"
Signed in, but the following Microsoft Graph permissions are missing from
the actual consent grant for '$($consentCheck.AppDisplayName)' in this
tenant (not just requested - genuinely not consented server-side):
$($consentCheck.MissingScopes -join ', ')

This can happen the first time this exact combination of permissions is
requested in a tenant, if the interactive consent prompt didn't fully
complete. Fix: in the Entra admin center, go to Identity > Applications >
Enterprise applications > '$($consentCheck.AppDisplayName)' > Permissions,
click "Grant admin consent for <tenant>", then Connect again.
"@
    }
}

function Test-RequiredGraphScopesGranted {
    <#
        Queries the tenant's actual OAuth2 permission grants for this
        app+resource combination via Graph itself, rather than trusting
        the signed-in session's own reported scope list.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$RequiredScopes
    )

    $context = Get-MgContext
    $clientSp = Get-MgServicePrincipal -Filter "appId eq '$($context.ClientId)'" -ErrorAction Stop
    $graphResourceSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction Stop

    $grants = @(Get-MgOauth2PermissionGrant -All -ErrorAction Stop |
        Where-Object { $_.ClientId -eq $clientSp.Id -and $_.ResourceId -eq $graphResourceSp.Id })

    $grantedScopes = [System.Collections.Generic.List[string]]::new()
    foreach ($grant in $grants) {
        if ($grant.Scope) { $grant.Scope -split '\s+' | Where-Object { $_ } | ForEach-Object { $grantedScopes.Add($_) } }
    }
    $grantedScopes = @($grantedScopes | Select-Object -Unique)

    $missing = @($RequiredScopes | Where-Object { $grantedScopes -notcontains $_ })

    [pscustomobject]@{
        AllGranted     = ($missing.Count -eq 0)
        MissingScopes  = $missing
        GrantedScopes  = $grantedScopes
        AppDisplayName = $clientSp.DisplayName
    }
}

function Disconnect-RoomProvisioningServices {
    [CmdletBinding()]
    param()

    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
}

Export-ModuleMember -Function Install-RoomProvisioningModules, Connect-RoomProvisioningServices, Disconnect-RoomProvisioningServices, Initialize-ConsoleVisibilityControl
