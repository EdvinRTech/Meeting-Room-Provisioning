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
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [int]$MaxRetries = 10,
        [int]$DelaySeconds = 20,
        [scriptblock]$ProgressCallback,
        [scriptblock]$LogCallback
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        if ($ProgressCallback) { & $ProgressCallback $attempt $MaxRetries }
        try {
            & $Action
            return [pscustomobject]@{ Success = $true; Attempts = $attempt; Error = $null }
        } catch {
            $errorMessage = $_.Exception.Message
            if ($LogCallback) { & $LogCallback "Attempt $attempt of $MaxRetries failed: $errorMessage" }
            if ($attempt -lt $MaxRetries) {
                Start-Sleep -Seconds $DelaySeconds
            } else {
                return [pscustomobject]@{ Success = $false; Attempts = $attempt; Error = $errorMessage }
            }
        }
    }
}

Export-ModuleMember -Function Invoke-WithRetryProgress
