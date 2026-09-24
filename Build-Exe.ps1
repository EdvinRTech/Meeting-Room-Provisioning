<#
.SYNOPSIS
    Builds Start-MeetingRoomProvisioning.exe from Launcher.ps1 via PS2EXE.

.DESCRIPTION
    Run this after changing Launcher.ps1 and commit the resulting .exe -
    it is not generated automatically, so it can silently go stale if a
    change to Launcher.ps1 is committed without also re-running this.
    (Changes to Start-MeetingRoomProvisioning.ps1 itself do NOT need a
    rebuild - the .exe is only a launcher stub that hands off to that
    script by file path, it doesn't embed it. See Launcher.ps1 for why.)

    Installs the ps2exe module to CurrentUser scope if it isn't already
    present.

    The very first time a freshly-built (unsigned) .exe runs on a given
    machine, Windows Defender/SmartScreen commonly scans it before letting
    it start - this can look like a 5-15 second hang on that first launch
    specifically. Confirmed while building this: every run after that
    first one starts and hands off promptly and consistently. Not a bug,
    just worth knowing before assuming something's wrong.
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'Start-MeetingRoomProvisioning.exe')
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host 'Installing ps2exe module (CurrentUser scope)...'
    Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber
}
Import-Module ps2exe -Force

# Deliberately NOT -noConsole: confirmed by actually running the built exe
# (not just reading PS2EXE's docs) that -noConsole hangs indefinitely here
# instead of exiting once Launcher.ps1's work is done - a real reliability
# regression, and exactly the kind of "looks like nothing happens" failure
# this tool has already been bitten by once. A plain console-subsystem
# build exits promptly and correctly every time it was tested; the cost is
# a brief, harmless console flash while the launcher runs (well under a
# second - Launcher.ps1 does almost nothing before handing off).
Invoke-ps2exe `
    -inputFile (Join-Path $PSScriptRoot 'Launcher.ps1') `
    -outputFile $OutputPath `
    -STA `
    -title 'Meeting Room Provisioning' `
    -company 'Asurgent AB' `
    -product 'Meeting Room Provisioning' `
    -description 'Launcher for the Meeting Room Provisioning wizard - hands off to Start-MeetingRoomProvisioning.ps1' `
    -version '1.0.0.0'

Write-Host "Built $OutputPath" -ForegroundColor Green
