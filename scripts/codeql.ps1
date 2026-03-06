# Define CodeQL workspace under repo root
$scriptroot = $PSScriptRoot
# src root is parent of scripts root
$srcRoot = Join-Path -Path $scriptroot -ChildPath ".."
$codeqlRoot = Join-Path -Path $srcRoot -ChildPath "codeql"
$codeqlBin = Join-Path $codeqlRoot "bin"
$codeqlDbRoot = Join-Path $codeqlRoot "db"
$codeqlSarifRoot = Join-Path $codeqlRoot ".sarif"
$codeql = Join-Path $codeqlBin "codeql.exe"
$db = Join-Path $codeqlDbRoot "database"
$sarif = Join-Path $codeqlSarifRoot "codeql.sarif"

Write-Host "CodeQL root: $codeqlRoot"
Write-Host "CodeQL bin: $codeqlBin"
Write-Host "CodeQL db: $db"
Write-Host "CodeQL sarif: $sarif"
Write-Host "Source root: $srcRoot"
Write-Host "CodeQL: $codeql"

if (-not (Test-Path -Path $codeqlRoot)) {
    New-Item -ItemType Directory -Force -Path $codeqlRoot | Out-Null
}
if (-not (Test-Path -Path $codeqlBin)) {
    New-Item -ItemType Directory -Force -Path $codeqlBin | Out-Null
}
if (-not (Test-Path -Path $codeqlDbRoot)) {
    New-Item -ItemType Directory -Force -Path $codeqlDbRoot | Out-Null
}
if (-not (Test-Path -Path $codeqlSarifRoot)) {
    New-Item -ItemType Directory -Force -Path $codeqlSarifRoot | Out-Null
}

# if codeql.exe does not exist, install it by running install-codeql.ps1
if (-not (Test-Path -Path $codeql)) {
    Write-Host "CodeQL is not installed. Installing..."
    & $scriptroot\install-codeql.ps1 -InstallDir $codeqlBin
}

# Verify installation by displaying the CodeQL version
& $codeql --version

# & $codeql resolve languages --format=betterjson --extractor-options-verbosity=4 --extractor-include-aliases
# & $codeql resolve queries cpp-code-scanning.qls --format=bylanguage

Write-Host "Creating database..."
if (Test-Path -Path $db) {
    Remove-Item -Path $db -Recurse -Force
}

$buildCommand = "npm run clean:release:x64 && npm run build:release:x64"
Write-Host "CodeQL build command: $buildCommand"
& $codeql database create $db --overwrite --language=cpp --command=$buildCommand

Write-Host "Running analysis..."
& $codeql database analyze $db --threads=4 -v --format sarif-latest -o $sarif