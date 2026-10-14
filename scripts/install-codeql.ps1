param(
    [string]$InstallDir = (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath "..") -ChildPath "codeql\\bin")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# When updating the CodeQL Action pin, update the expected hash below using the
# matching bundle's release asset metadata:
# https://api.github.com/repos/github/codeql-action/releases/tags/<bundle-version>
$expectedHash = "C4834B337E2DB85BDD4E7E97455319FF70D28F06ABB2987EFF9EC0F62D51DCEA"

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
$defaults = Invoke-RestMethod -Uri $defaultsUrl -ErrorAction Stop
if (-not $defaults.bundleVersion) {
    throw "bundleVersion was not found in $defaultsUrl."
}

$bundleVersion = [string]$defaults.bundleVersion
if ($bundleVersion -notmatch '^codeql-bundle-v[0-9]+\.[0-9]+\.[0-9]+$') {
    throw "Unexpected bundleVersion '$bundleVersion' in $defaultsUrl."
}

$bundleName = "codeql-bundle-win64.tar.gz"

try {
    $latestDefaults = Invoke-RestMethod `
        -Uri "https://raw.githubusercontent.com/github/codeql-action/main/src/defaults.json" `
        -ErrorAction Stop
    $latestVersion = ([string]$latestDefaults.bundleVersion) -replace '^codeql-bundle-v', ''
    $pinnedVersion = $bundleVersion -replace '^codeql-bundle-v', ''
    if ([version]$pinnedVersion -lt [version]$latestVersion) {
        Write-Warning "Pinned CodeQL bundle $bundleVersion is older than the current CodeQL Action default $($latestDefaults.bundleVersion). Review the update instructions at the top of this script."
    }
}
catch {
    Write-Warning "Unable to check for a newer CodeQL bundle; continuing with pinned version $bundleVersion. $($_.Exception.Message)"
}

#Create the installation directory if it doesn't exist
if (-not (Test-Path -Path $InstallDir)) {
    New-Item -ItemType Directory -Force -Path $InstallDir
}

$codeql = Join-Path $InstallDir "codeql.exe"
$versionFile = Join-Path -Path $InstallDir -ChildPath ".version"
#Check the version in the version file - if it matches the bundle version and codeql.exe exists, exit
if ((Test-Path -Path $versionFile) -and ((Get-Content -Path $versionFile) -eq $bundleVersion) -and (Test-Path -Path "$InstallDir\codeql.exe")) {
    Write-Host "CodeQL is already installed."
    & $codeql --version
    if ($LASTEXITCODE -ne 0) {
        throw "Installed CodeQL CLI exited with code $LASTEXITCODE."
    }
    exit 0
}

$tarFile = Join-Path -Path $InstallDir -ChildPath $bundleName
$downloadUrl = "https://github.com/github/codeql-action/releases/download/" + $bundleVersion + '/' + $bundleName

Write-Host "Downloading CodeQL CLI from $downloadUrl"
& curl.exe --fail --location --show-error --silent --output $tarFile $downloadUrl
if ($LASTEXITCODE -ne 0) {
    throw "CodeQL bundle download failed with curl exit code $LASTEXITCODE."
}

$actualHash = (Get-FileHash -Path $tarFile -Algorithm SHA256).Hash
if ($actualHash -ine $expectedHash) {
    throw "CodeQL bundle SHA-256 mismatch. Expected $expectedHash; received $actualHash. If the CodeQL Action pin changed, update the expected hash using its release asset metadata."
}
Write-Host "CodeQL bundle SHA-256 verified."

Write-Host "Preparing install directory $InstallDir"
Get-ChildItem -Path $InstallDir -Force | Where-Object { $_.Name -notin @($bundleName, ".version") } | Remove-Item -Recurse -Force

Write-Host "Extracting $tarFile"
& tar.exe -xf $tarFile --strip-components=1 -C $InstallDir
if ($LASTEXITCODE -ne 0) {
    throw "CodeQL bundle extraction failed with tar exit code $LASTEXITCODE."
}

Write-Host "Extracted to $InstallDir"
if (-not (Test-Path -Path $codeql)) {
    throw "Installation failed: '$codeql' was not found after extraction."
}

# Verify installation by displaying the CodeQL version
& $codeql --version
if ($LASTEXITCODE -ne 0) {
    throw "Installed CodeQL CLI exited with code $LASTEXITCODE."
}

# Write the installed bundle version after successful verification
$bundleVersion | Out-File -FilePath $versionFile -Encoding ascii
