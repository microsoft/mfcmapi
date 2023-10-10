# Define the installation directory as a subdir of scripts
$codeqlRoot = Join-Path -Path $PSScriptRoot -ChildPath "codeql"
$codeql = Join-Path $codeqlRoot "codeql.exe"

# if codeql.exe does not exist, install it by running install-codeql.ps1
if (-not (Test-Path -Path "$codeqlRoot\codeql.exe")) {
    Write-Host "CodeQL is not installed. Installing..."
    & $PSScriptRoot\install-codeql.ps1
}

# Verify installation by displaying the CodeQL version
& $codeql --version

& $codeql resolve languages --format=betterjson --extractor-options-verbosity=4 --extractor-include-aliases
& $codeql resolve queries cpp-code-scanning.qls --format=bylanguage