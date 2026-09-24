<#
.SYNOPSIS
    Source for Start-MeetingRoomProvisioning.exe - lets the tool be
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
    Start-MeetingRoomProvisioning.ps1 sitting next to it and hand off to a
    genuine pwsh.exe (or powershell.exe, if PowerShell 7 isn't installed)
    process running it. $PSScriptRoot is empty inside a compiled
    executable (also confirmed empirically) but resolves completely
    normally in that handed-off process, so every one of the real
    script's own mechanisms - elevation, STA, PowerShell 7
    preference/auto-install, startup error logging - applies completely
    unchanged, exactly as if you'd right-clicked the .ps1 and chosen "Run
    with PowerShell" yourself.
#>

$ErrorActionPreference = 'Stop'
try {
    # $PSScriptRoot/$PSCommandPath are both empty in a compiled exe - the
    # running process's own module path is the only reliable way to find
    # "the folder this .exe is sitting in".
    $here = Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    $realScript = Join-Path $here 'Start-MeetingRoomProvisioning.ps1'

    if (-not (Test-Path -LiteralPath $realScript)) {
        throw "Expected to find Start-MeetingRoomProvisioning.ps1 in the same folder as this .exe ($here), but it isn't there. Copy the whole MeetingRoomProvisioning folder - this .exe alone isn't enough."
    }

    $pwsh = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue
    $hostExe = if ($pwsh) { $pwsh.Source } else { 'powershell.exe' }

    # No -Wait: this stub's job ends the moment the real process starts.
    # No -Verb RunAs either - the real script already does its own
    # elevation check and self-relaunch, so duplicating that here would
    # just be a second, redundant UAC prompt in the case it's needed.
    Start-Process -FilePath $hostExe -ArgumentList @('-NoProfile', '-STA', '-File', "`"$realScript`"") -ErrorAction Stop
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
