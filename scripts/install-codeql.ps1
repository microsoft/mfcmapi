# Define the GitHub repository and the desired CodeQL version
$repoUrl = "https://github.com/github/codeql-action/releases/latest"
$bundleName = "codeql-bundle-win64.tar.gz"

# Define the installation directory as a subdir of scripts
$installDir = Join-Path -Path $PSScriptRoot -ChildPath "codeql"
$codeql = Join-Path $installDir "codeql.exe"

# Create the installation directory if it doesn't exist
if (-not (Test-Path -Path $installDir)) {
    New-Item -ItemType Directory -Force -Path $installDir
}

# Get the latest release URL
$latestReleasePage = Invoke-RestMethod -Uri $repoUrl
$bundleVersion = ($latestReleasePage | Select-String -Pattern "https://github.com/github/codeql-action/releases/expanded_assets/(codeql-bundle-v\d+\.\d+\.\d+)").Matches.Groups[1].Value
$versionFile = Join-Path -Path $installDir -ChildPath ".version"
#Check the version in the version file - if it matches the bundle version and codeql.exe exists, exit
if ((Test-Path -Path $versionFile) -and ((Get-Content -Path $versionFile) -eq $bundleVersion) -and (Test-Path -Path "$installDir\codeql.exe")) {
    Write-Host "CodeQL is already installed."
    & $codeql --version
    exit 0
}

$tarFile = Join-Path -Path $installDir -ChildPath $bundleName

# create a .version file with the #bundleVersion written to it
$bundleVersion | Out-File -FilePath $versionFile -Encoding ascii

# Download the latest CodeQL CLI release
Push-Location $installDir
$downloadUrl = "https://github.com/github/codeql-action/releases/download/" + $bundleVersion + '/'+ $bundleName
# Invoke-WebRequest -Uri $downloadUrl -OutFile $tarFile
Write-Host "Downloading CodeQL CLI from $downloadUrl"
Write-Host "Saving to $installDir"
curl.exe -LO  $downloadUrl

# Extract the downloaded archive
Write-Host "Extracting $tarFile"
tar -xf $tarFile -C $PSScriptRoot
Write-Host "Extracted to $PSScriptRoot"
Pop-Location

# Verify installation by displaying the CodeQL version
& $codeql --version
