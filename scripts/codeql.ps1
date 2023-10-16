# Define the installation directory as a subdir of scripts
$scriptroot = $PSScriptRoot
$codeqlRoot = Join-Path -Path $scriptroot -ChildPath "codeql"
# src root is parent of scripts root
$srcRoot = Join-Path -Path $scriptroot -ChildPath ".."
$codeql = Join-Path $codeqlRoot "codeql.exe"
$db = Join-Path $env:TEMP "codeql_databases"

Write-Host "CodeQL root: $codeqlRoot"
Write-Host "Source root: $srcRoot"
Write-Host "CodeQL: $codeql"

# if codeql.exe does not exist, install it by running install-codeql.ps1
if (-not (Test-Path -Path "$codeqlRoot\codeql.exe")) {
    Write-Host "CodeQL is not installed. Installing..."
    & $scriptroot\install-codeql.ps1
}

# Verify installation by displaying the CodeQL version
& $codeql --version

# & $codeql resolve languages --format=betterjson --extractor-options-verbosity=4 --extractor-include-aliases
# & $codeql resolve queries cpp-code-scanning.qls --format=bylanguage

Write-Host "Creating database..."
if (Test-Path -Path $db) {
    Remove-Item -Path $db -Recurse -Force
}

& $codeql database create $db --overwrite --language=cpp --command="powershell $scriptroot\build.ps1"

Write-Host "Running analysis..."
& $codeql database analyze $db --threads=4 -v --format sarif-latest -o $srcRoot\.sarif\codeql.sarif