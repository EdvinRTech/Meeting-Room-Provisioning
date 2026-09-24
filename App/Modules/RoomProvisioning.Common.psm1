#requires -Version 5.1

function Invoke-WithRetryProgress {
    <#
        Runs $Action up to $MaxRetries times, waiting $DelaySeconds between
        attempts. Before each attempt (including the first) it calls
        $ProgressCallback with the current attempt number and $MaxRetries,
        so a GUI can drive a progress bar (Value = attempt, Maximum =
        MaxRetries) instead of showing a generic spinner.

        Used for Graph operations that can lag behind object creation
        (directory replication), e.g. adding a brand-new room account to a
        security group, or setting its password profile right after it was
        created.

        This tool has no background thread - everything, including this
        wait, runs on the single WPF UI thread. A plain Start-Sleep would
        freeze the whole window for the entire delay, so nothing (not even
        a Cancel button) could respond during it. Instead, the delay is
        cut into small chunks (SleepStepMilliseconds each) with -SleepStep
        invoked between chunks - the caller uses that hook to pump its own
        UI dispatcher, which is what lets a button click actually get
        processed and its handler run while this function is "sleeping".
        -CancelCheck is polled before each attempt and between every sleep
        chunk, so a user-requested cancellation takes effect within about
        one chunk's delay rather than only between whole retry attempts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [int]$MaxRetries = 10,
        [int]$DelaySeconds = 20,
        [scriptblock]$ProgressCallback,
        [scriptblock]$LogCallback,
        [scriptblock]$CancelCheck,
        [scriptblock]$SleepStep,
        [int]$SleepStepMilliseconds = 200
    )

    function Test-Cancelled {
        if ($CancelCheck) { return [bool](& $CancelCheck) }
        return $false
    }

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        if (Test-Cancelled) {
            return [pscustomobject]@{ Success = $false; Attempts = $attempt; Error = 'Cancelled by user.'; Cancelled = $true }
        }
        if ($ProgressCallback) { & $ProgressCallback $attempt $MaxRetries }
        try {
            & $Action
            return [pscustomobject]@{ Success = $true; Attempts = $attempt; Error = $null; Cancelled = $false }
        } catch {
            $errorMessage = $_.Exception.Message
            if ($LogCallback) { & $LogCallback "Attempt $attempt of $MaxRetries failed: $errorMessage" }
            if ($attempt -lt $MaxRetries) {
                if ($SleepStep) {
                    $elapsedMs = 0
                    $totalMs = $DelaySeconds * 1000
                    while ($elapsedMs -lt $totalMs) {
                        if (Test-Cancelled) {
                            return [pscustomobject]@{ Success = $false; Attempts = $attempt; Error = 'Cancelled by user.'; Cancelled = $true }
                        }
                        $thisChunk = [Math]::Min($SleepStepMilliseconds, $totalMs - $elapsedMs)
                        Start-Sleep -Milliseconds $thisChunk
                        & $SleepStep
                        $elapsedMs += $thisChunk
                    }
                } else {
                    Start-Sleep -Seconds $DelaySeconds
                }
            } else {
                return [pscustomobject]@{ Success = $false; Attempts = $attempt; Error = $errorMessage; Cancelled = $false }
            }
        }
    }
}

Export-ModuleMember -Function Invoke-WithRetryProgress
