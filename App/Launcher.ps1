<#
.SYNOPSIS
    Source for "M365 Meeting Room Tool.exe" - lets the tool be
    double-clicked directly instead of needing "right-click > Run with
    PowerShell" on the .ps1. Compiled via Build-Exe.ps1 (PS2EXE); rebuild
    and re-commit the .exe after changing this file.

.DESCRIPTION
    Deliberately does almost nothing itself. PS2EXE-compiled executables
    always host Windows PowerShell 5.1 Desktop internally, regardless of
    which PowerShell version builds them - confirmed empirically while
    building this (there's no way to get PS2EXE to produce a
    PowerShell-7-hosted .exe). That's exactly the engine
    Start-MeetingRoomProvisioning.ps1's own relaunch logic exists to get
    away from, for the Microsoft Graph SDK's "GetTokenAsync ... lacks an
    implementation" bug. So instead of compiling the real ~1300-line
    application (and inheriting that Desktop-only limitation for good),
    this stub's only job is to find the real
    Start-MeetingRoomProvisioning.ps1 - sitting in the .\App folder next to
    this .exe, kept out of sight (hidden attribute) so the distributed
    folder shows just the .exe and nothing else - and hand off to a
    genuine pwsh.exe (or powershell.exe, if PowerShell 7 isn't installed)
    process running it.

    This hands off directly to the SAME end state the real script's own
    relaunch logic targets - elevated, STA, -RelaunchedForGui already set
    - in one Start-Process call, rather than launching a plain hop that
    then has to elevate *itself* a second time. That second, nested
    "-Verb RunAs combined with -WindowStyle Hidden, on a process that was
    itself already relaunched" combination is the one that was confirmed
    to leave a visible elevated console behind (full of module-install
    warning text) once the .exe launcher started going through this file
    instead of a direct "Run with PowerShell" - a single RunAs+Hidden hop,
    straight from the interactive process the user actually double-clicked
    (exactly what right-clicking the .ps1 always did), is the combination
    that's actually been proven to hide correctly. Every one of the real
    script's own mechanisms - STA, PowerShell 7 preference/auto-install,
    startup error logging - still applies completely unchanged once handed
    off; only the extra elevation hop is skipped.
#>

$ErrorActionPreference = 'Stop'
try {
    # $PSScriptRoot/$PSCommandPath are both empty in a compiled exe - the
    # running process's own module path is the only reliable way to find
    # "the folder this .exe is sitting in".
    $here = Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    $appFolder = Join-Path $here 'App'
    $realScript = Join-Path $appFolder 'Start-MeetingRoomProvisioning.ps1'

    if (-not (Test-Path -LiteralPath $realScript)) {
        throw "Expected to find an App\Start-MeetingRoomProvisioning.ps1 next to this .exe ($here), but it isn't there. Copy the whole distributed folder, including the (hidden) App folder - this .exe alone isn't enough."
    }

    # Re-applied on every launch, not just once at build time: a hidden
    # attribute is filesystem-level metadata, not file content, so it does
    # not survive being re-zipped/re-extracted or re-cloned the way the
    # folder's actual contents do. Cheap and idempotent either way.
    try { (Get-Item -LiteralPath $appFolder -Force).Attributes = 'Directory, Hidden' } catch {}

    # PATH first, then the default per-machine install location directly -
    # same two-step lookup Start-MeetingRoomProvisioning.ps1 itself uses
    # (Get-Pwsh7Path), needed because an installer updates the registry's
    # Environment key, not this already-running process's in-memory PATH.
    $pwsh = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue
    $hostExe = if ($pwsh) {
        $pwsh.Source
    } else {
        $defaultPwshPath = Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
        if (Test-Path -LiteralPath $defaultPwshPath) { $defaultPwshPath } else { 'powershell.exe' }
    }

    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [System.Security.Principal.WindowsPrincipal]::new($identity)
    $isElevated = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

    # -RelaunchedForGui here means "already elevated and STA, nothing left
    # for your own relaunch logic to do" - the real script checks this
    # itself and, in the ordinary case, takes no further relaunch hop at
    # all.
    $startArgs = @{
        FilePath     = $hostExe
        WindowStyle  = 'Hidden'
        ArgumentList = @('-NoProfile', '-STA', '-File', "`"$realScript`"", '-RelaunchedForGui')
        ErrorAction  = 'Stop'
    }
    # Only request elevation if we don't already have it - if the .exe
    # itself was already run as Administrator, adding -Verb RunAs here
    # would trigger a second, redundant UAC prompt.
    if (-not $isElevated) { $startArgs.Verb = 'RunAs' }

    # No -Wait: this stub's job ends the moment the real process starts.
    Start-Process @startArgs
} catch {
    # Same principle as the real script's own top-level error handling:
    # a launcher that fails silently is indistinguishable from "nothing
    # happens", which is exactly the class of bug this tool has already
    # been bitten by once.
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show(
            "Meeting Room Provisioning couldn't start:`n`n$($_.Exception.Message)",
            'Meeting Room Provisioning - Launcher Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } catch {}
}
