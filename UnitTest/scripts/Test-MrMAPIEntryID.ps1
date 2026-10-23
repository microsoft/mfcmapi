param(
    [Parameter(Mandatory = $true)]
    [string]$MrMAPI
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$executable = (Resolve-Path -LiteralPath $MrMAPI).Path
$script:checks = 0

function Check-Command {
    param([string[]]$Arguments, [int]$ExitCode, [string]$ExpectedText)

    $start = [System.Diagnostics.ProcessStartInfo]::new()
    $start.FileName = $executable
    $start.UseShellExecute = $false
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.StandardOutputEncoding = [System.Text.Encoding]::Unicode
    $start.ArgumentList.Add("-NoAddins")
    foreach ($argument in $Arguments) { $start.ArgumentList.Add($argument) }
    $process = [System.Diagnostics.Process]::Start($start)
    try {
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(30000)) {
            $process.Kill()
            throw "MrMAPI timed out: $Arguments"
        }
        $output = $stdout.GetAwaiter().GetResult()
        $errors = $stderr.GetAwaiter().GetResult()
        if ($process.ExitCode -ne $ExitCode -or ($ExpectedText -and -not $output.Contains($ExpectedText))) {
            throw "Unexpected result ($($process.ExitCode)): $Arguments`n$output`n$errors"
        }
        $script:checks++
    }
    finally {
        $process.Dispose()
    }
}

# These invalid commands must exit before MAPI initialization, with no profile access.
$request = @("-EntryID", "00AB", "-Profile", "Fixture only - never log on", "-Store", "0", "PR_SUBJECT_W")
Check-Command -Arguments @("-EntryID") -ExitCode 2 -ExpectedText "Invalid -EntryID"
Check-Command -Arguments @("-EntryID", "00AB", "PR_SUBJECT_W") -ExitCode 2 -ExpectedText "Invalid -EntryID"
foreach ($hex in @("", "0", "0x00AB", "00 AB", "zz")) {
    $invalid = $request.Clone()
    $invalid[1] = $hex
    Check-Command -Arguments $invalid -ExitCode 2 -ExpectedText "Invalid -EntryID"
}
foreach ($property in @("subject", "0037", "0037001Fjunk", "00370000")) {
    $invalid = $request.Clone()
    $invalid[6] = $property
    Check-Command -Arguments $invalid -ExitCode 2 -ExpectedText "Invalid -EntryID"
}
Check-Command -Arguments ($request + @("-Contents")) -ExitCode 2 -ExpectedText "Invalid -EntryID"
Check-Command -Arguments ($request + @("-Output", "must-not-be-created.xml")) -ExitCode 2 -ExpectedText "Invalid -EntryID"

# Version 0 exits after successful parsing and lists implementations WITHOUT logging on.
foreach ($property in @("PR_SUBJECT_W", "0037001F", "0x0037001f")) {
    $valid = $request.Clone()
    $valid[6] = $property
    Check-Command -Arguments ($valid + @("-Version", "0")) -ExitCode 0
}
Check-Command -Arguments @("-Help") -ExitCode 0 -ExpectedText "Message by EntryID:"
Check-Command -Arguments @("PR_SUBJECT_W") -ExitCode 0 -ExpectedText "PR_SUBJECT"
Write-Output "Passed $script:checks mailbox-free MrMAPI process checks."
