param(
    [switch]$Check,
    [string[]]$Files
)

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$clang = Join-Path $projectRoot "node_modules\clang-format\bin\win32\clang-format.exe"

if (-not (Test-Path $clang)) {
    Write-Error "clang-format not found at $clang. Run 'npm install' first."
    exit 1
}

Write-Host "clang-format found at $clang"
& $clang --version

Push-Location $projectRoot

if ($Files) {
    # Resolve provided paths to absolute
    $files = $Files | ForEach-Object { Resolve-Path $_ | Select-Object -ExpandProperty Path }
} else {
    $files = Get-ChildItem -Recurse -Include *.cpp, *.h, *.c |
        Where-Object { $_.FullName -notlike "*\mapistub\*" -and $_.FullName -notlike "*\.git\*" } |
        Select-Object -ExpandProperty FullName
}

if (-not $files) {
    Write-Host "No C/C++ files found."
    Pop-Location
    exit 0
}

Write-Host "Found $($files.Count) files."

# Batch into chunks of 30 to avoid command-line length limits
$chunkSize = 30
$exitCode = 0

$baseArgs = @('--style=file', '--fallback-style=Microsoft', '--verbose')
if ($Check) {
    Write-Host "Checking formatting..."
    $baseArgs += @('--dry-run', '-Werror')
} else {
    Write-Host "Formatting files..."
    $baseArgs += '-i'
}

for ($i = 0; $i -lt $files.Count; $i += $chunkSize) {
    $chunk = $files[$i..([Math]::Min($i + $chunkSize - 1, $files.Count - 1))]
    & $clang @baseArgs @chunk
    if ($LASTEXITCODE -ne 0) { $exitCode = $LASTEXITCODE }
}

Pop-Location
exit $exitCode
