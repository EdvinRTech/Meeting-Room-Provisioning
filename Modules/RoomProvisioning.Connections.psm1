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
    'Microsoft.Graph.Identity.DirectoryManagement'
)

# ExchangeOnlineManagement is a separate product with its own release
# cadence - it is not part of the Graph version-matching below, just
# always installed/imported at its own latest version.
$Script:RequiredModules = @('ExchangeOnlineManagement') + $Script:GraphSubModules

$Script:GraphScopes = @(
    'User.ReadWrite.All',
    'Group.ReadWrite.All',
    'Policy.Read.All',
    'Policy.ReadWrite.ConditionalAccess',
    'Directory.Read.All',
    'Organization.Read.All'
)

function Get-MatchedGraphModuleVersion {
    <#
        Finds the newest version of each module in $ModuleNames that is
        published on PSGallery, then returns the *lowest* of those
        "latest" versions - i.e. the newest version that every module in
        the set has actually released - and confirms that exact version
        really exists for every module (a submodule occasionally skips a
        release). This is what lets every Graph submodule be installed at
        one identical, known-to-exist version instead of each one
        independently grabbing its own latest (and possibly mismatched)
        release.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$ModuleNames,
        [scriptblock]$LogCallback
    )

    # NOTE: PowerShell scriptblocks resolve unqualified variables/functions
    # by walking the DYNAMIC call stack at invocation time, not by where
    # they were lexically written - so a wrapper like this, invoked later
    # from a caller's own scope, can silently pick up a same-named local
    # variable from whatever function happens to invoke it instead of the
    # one it was meant to close over. .GetNewClosure() snapshots the
    # variables it references at creation time, making it safe to pass
    # around and invoke from anywhere. Every callback in this tool is
    # built this way - see README "Why every callback uses GetNewClosure".
    $emit = {
        param($msg)
        if ($LogCallback) { & $LogCallback $msg }
    }.GetNewClosure()

    $latestPerModule = @{}
    foreach ($name in $ModuleNames) {
        & $emit "Checking the latest published version of $name..."
        $found = Find-Module -Name $name -ErrorAction Stop
        $latestPerModule[$name] = [version]$found.Version
    }

    $targetVersion = ($latestPerModule.Values | Sort-Object)[0]
    & $emit "Matched version for all Microsoft.Graph modules: $targetVersion"

    foreach ($name in $ModuleNames) {
        if ($latestPerModule[$name] -ne $targetVersion) {
            if (-not (Find-Module -Name $name -RequiredVersion $targetVersion -ErrorAction SilentlyContinue)) {
                throw "No published release of '$name' matches version $targetVersion - cannot pin a single common Graph module version automatically. Try again later or report this so the version set can be adjusted."
            }
        }
    }

    return $targetVersion
}

function Install-RoomProvisioningModules {
    <#
        Guarantees a clean, matched set of required modules regardless of
        whatever was already on the machine: force-removes every existing
        install of each required module, works out one Graph module
        version common to all of them, then installs and imports that
        exact set. This is intentionally invasive (it will remove other
        versions of these modules that other scripts on the machine might
        be using) - it trades that for the tool reliably working the same
        way on any machine, instead of failing with hard-to-diagnose
        assembly-version errors depending on whatever happened to be
        installed already.

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
    }.GetNewClosure()

    # Uninstalling existing copies was dropped: on a typical machine it
    # can't succeed anyway (AllUsers-scope installs need admin rights to
    # remove, which this tool intentionally doesn't require), so it only
    # produced a failed-removal message every run without changing
    # anything. Reliability instead comes from installing a matched
    # version to CurrentUser scope and importing it by its exact
    # installed path below - see README "CurrentUser modules are
    # preferred, and imported unambiguously".
    foreach ($module in $Script:RequiredModules) {
        Get-Module -Name $module -ErrorAction SilentlyContinue | Remove-Module -Force -ErrorAction SilentlyContinue
    }

    & $write 'Working out a single matched version for all Microsoft.Graph modules...'
    $graphVersion = Get-MatchedGraphModuleVersion -ModuleNames $Script:GraphSubModules -LogCallback $write

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
        & $write 'Installing ExchangeOnlineManagement (latest)...'
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        $exoInfo = Get-InstalledModule -Name ExchangeOnlineManagement -ErrorAction Stop
        & $write "Installed ExchangeOnlineManagement $($exoInfo.Version)."

        $graphModuleInfo = @{}
        foreach ($module in $Script:GraphSubModules) {
            & $write "Installing $module $graphVersion..."
            Install-Module -Name $module -RequiredVersion $graphVersion -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            $graphModuleInfo[$module] = Get-InstalledModule -Name $module -RequiredVersion $graphVersion -ErrorAction Stop
            & $write "Installed $module $graphVersion."
        }
    } catch {
        & $write "FAILED to install required modules: $($_.Exception.Message)"
        throw
    }

    try {
        & $write 'Importing modules...'
        Import-Module -Name (Join-Path $exoInfo.InstalledLocation 'ExchangeOnlineManagement.psd1') -Force -ErrorAction Stop
        foreach ($module in $Script:GraphSubModules) {
            $manifestPath = Join-Path $graphModuleInfo[$module].InstalledLocation "$module.psd1"
            Import-Module -Name $manifestPath -Force -ErrorAction Stop
        }
        & $write 'All modules installed and imported at a matched, known-good version set.'
    } catch {
        & $write "FAILED to import required modules: $($_.Exception.Message)"
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
# console the user opened themselves.
$Script:OwnsConsoleWindow = $false

function Initialize-ConsoleVisibilityControl {
    [CmdletBinding()]
    param([switch]$Enabled)
    $Script:OwnsConsoleWindow = [bool]$Enabled
    if ($Script:OwnsConsoleWindow) { Hide-RoomProvisioningConsole }
}

function Show-RoomProvisioningConsole {
    if (-not $Script:OwnsConsoleWindow) { return }
    $hwnd = [RoomProvisioning.ConsoleWindow]::GetConsoleWindow()
    [RoomProvisioning.ConsoleWindow]::ShowWindow($hwnd, 5) | Out-Null # SW_SHOW
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

    # -UseDeviceCode deliberately avoids the Windows broker (WAM)/embedded
    # sign-in window entirely: Connect-MgGraph instead prints a one-time
    # code and https://microsoft.com/devicelogin, and you finish signing
    # in in your normal web browser. The default interactive flow relies
    # on a WAM broker component that's inconsistently present/working
    # across machines - device code sidesteps that dependency completely,
    # at the cost of one extra manual step (typing the code).
    #
    # That code is printed to the console, not the GUI window, so the
    # (normally hidden - see Initialize-ConsoleVisibilityControl) console
    # is un-hidden just for this call and hidden again immediately after,
    # whether it succeeded or not.
    try {
        Show-RoomProvisioningConsole
        Connect-MgGraph -Scopes $Script:GraphScopes -NoWelcome -UseDeviceCode -ErrorAction Stop
    } finally {
        Hide-RoomProvisioningConsole
    }

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

Export-ModuleMember -Function Install-RoomProvisioningModules, Connect-RoomProvisioningServices, Disconnect-RoomProvisioningServices, Initialize-ConsoleVisibilityControl
