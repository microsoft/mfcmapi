param(
    [string]$InstallDir = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath "..") -ChildPath "codeql\\bin")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$workflowPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath "..") -ChildPath ".github\\workflows\\codeql.yml"

if (-not (Test-Path -Path $workflowPath)) {
    throw "CodeQL workflow not found at '$workflowPath'."
}

$workflowContent = Get-Content -Path $workflowPath -Raw
$shaMatch = [regex]::Match($workflowContent, 'uses:\s*github/codeql-action/init@([0-9a-f]{40})')
if (-not $shaMatch.Success) {
    throw "Unable to find a pinned github/codeql-action/init@<sha> entry in '$workflowPath'."
}

$actionSha = $shaMatch.Groups[1].Value
$defaultsUrl = "https://raw.githubusercontent.com/github/codeql-action/$actionSha/src/defaults.json"
Write-Host "Resolving CodeQL bundle from action SHA: $actionSha"

$defaults = Invoke-RestMethod -Uri $defaultsUrl
if (-not $defaults.bundleVersion) {
    throw "bundleVersion was not found in $defaultsUrl."
}

$bundleVersion = [string]$defaults.bundleVersion
if ($bundleVersion -notmatch '^codeql-bundle-v[0-9]+\.[0-9]+\.[0-9]+$') {
    throw "Unexpected bundleVersion '$bundleVersion' in $defaultsUrl."
}

$bundleName = "codeql-bundle-win64.tar.gz"

$codeql = Join-Path $installDir "codeql.exe"

# Create the installation directory if it doesn't exist
if (-not (Test-Path -Path $installDir)) {
    New-Item -ItemType Directory -Force -Path $installDir
}

$versionFile = Join-Path -Path $installDir -ChildPath ".version"
#Check the version in the version file - if it matches the bundle version and codeql.exe exists, exit
if ((Test-Path -Path $versionFile) -and ((Get-Content -Path $versionFile) -eq $bundleVersion) -and (Test-Path -Path "$installDir\codeql.exe")) {
    Write-Host "CodeQL is already installed."
    & $codeql --version
    exit 0
}

$tarFile = Join-Path -Path $installDir -ChildPath $bundleName

# Download the latest CodeQL CLI release
Push-Location $installDir
$downloadUrl = "https://github.com/github/codeql-action/releases/download/" + $bundleVersion + '/'+ $bundleName
# Invoke-WebRequest -Uri $downloadUrl -OutFile $tarFile
Write-Host "Downloading CodeQL CLI from $downloadUrl"
Write-Host "Saving to $installDir"
curl.exe -L -o $tarFile $downloadUrl

Write-Host "Preparing install directory $installDir"
Get-ChildItem -Path $installDir -Force | Where-Object { $_.Name -notin @($bundleName, ".version") } | Remove-Item -Recurse -Force

# Extract the downloaded archive
Write-Host "Extracting $tarFile"
tar -xf $tarFile --strip-components=1 -C $installDir

Write-Host "Extracted to $installDir"
Pop-Location

if (-not (Test-Path -Path $codeql)) {
    throw "Installation failed: '$codeql' was not found after extraction."
}

# Write the installed bundle version after successful extraction
$bundleVersion | Out-File -FilePath $versionFile -Encoding ascii

# Verify installation by displaying the CodeQL version
& $codeql --version
