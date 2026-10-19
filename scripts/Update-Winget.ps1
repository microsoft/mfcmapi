<#
.SYNOPSIS
Generates and validates WinGet manifests for a published MFCMAPI release.
.DESCRIPTION
Requires winget on PATH and installs or updates WingetCreate through WinGet.
No applicable WingetCreate update is a normal result.
Runs WingetCreate's GitHub OAuth sign-in flow and caches credentials before
retrieving manifests. Follow the sign-in instructions displayed by WingetCreate.
Uses GitHub's latest published stable MFCMAPI release unless ReleaseTag is specified,
not a local build. Downloads the x86 and x64 executable ZIPs through WingetCreate,
generates user-scope manifests, and validates them with winget validate.

Without Submit, no public GitHub changes are made. Network requests, downloads,
and normal tool telemetry can still occur. Generated manifests are retained locally;
the script prints their location and a command to submit them after review.

Submission is opt-in and uses the credentials cached during sign-in,
not a PAT supplied by this script. Submission can create
a public fork, branch, commit, and pull request in microsoft/winget-pkgs.
It does not merge the PR or immediately publish the package.
.PARAMETER ReleaseTag
Exact GitHub release tag to use instead of the latest stable release.
Drafts and prereleases are rejected.
.PARAMETER OutputDirectory
New directory for generated manifests. Defaults to a unique temporary directory.
Existing directories are rejected to prevent overwrites.
.PARAMETER Submit
Submit immediately after validation. Omit to review the generated YAML first.
.EXAMPLE
.\scripts\Update-Winget.ps1
Generates and validates the latest stable release without submitting a PR.
.EXAMPLE
.\scripts\Update-Winget.ps1 -ReleaseTag 26.0.26291.01 -Submit
Generates, validates, and submits the specified stable release.
.NOTES
Check for an existing release PR before submitting to avoid duplicates.

Authentication is part of every run; no separate authentication command is needed.
Sign-in alone does not create a public fork or PR. Authentication remains subject
to GitHub and organization policies and does not guarantee freedom from throttling.
.LINK
https://github.com/microsoft/winget-create/blob/main/doc/token.md
#>
[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$ReleaseTag,

    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory = (Join-Path ([System.IO.Path]::GetTempPath()) ("mfcmapi-winget-" + [guid]::NewGuid())),

    [switch]$Submit
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$winget = Get-Command winget.exe -ErrorAction SilentlyContinue
if (-not $winget) {
    throw "Install the WinGet client (App Installer) before running this script."
}

Write-Host "Installing or updating WingetCreate through WinGet..."
& $winget install --id Microsoft.WingetCreate --exact --source winget
$noApplicableUpdate = [int]0x8A15002B
if ($LASTEXITCODE -eq $noApplicableUpdate) {
    Write-Host "WingetCreate has no applicable update."
}
elseif ($LASTEXITCODE -ne 0) {
    throw "WingetCreate installation or update failed with exit code $LASTEXITCODE."
}
$wingetCreate = Get-Command wingetcreate.exe -ErrorAction SilentlyContinue
if (-not $wingetCreate) {
    throw "WingetCreate is not on PATH after installation or update. Open a new terminal and run this script again."
}

Write-Host "Signing in to GitHub for WingetCreate. Follow the OAuth sign-in instructions..."
& $wingetCreate token --store
if ($LASTEXITCODE -ne 0) {
    throw "WingetCreate GitHub authentication failed with exit code $LASTEXITCODE."
}

$releaseUrl = "https://api.github.com/repos/microsoft/mfcmapi/releases/latest"
if ($ReleaseTag) {
    $releaseUrl = "https://api.github.com/repos/microsoft/mfcmapi/releases/tags/$([uri]::EscapeDataString($ReleaseTag))"
}
$release = Invoke-RestMethod -Uri $releaseUrl -Headers @{
    Accept = "application/vnd.github+json"
    "User-Agent" = "MFCMAPI-Update-Winget"
}
if ($release.draft -or $release.prerelease) {
    throw "Only published stable releases are supported. Release '$($release.tag_name)' is a draft or prerelease."
}

$version = $release.tag_name -ireplace '^v', ''
if ($version -notmatch '^\d+\.\d+\.\d+\.\d+$') {
    throw "Unexpected MFCMAPI release version '$version'. Expected four numeric components."
}

$x86Assets = @($release.assets | Where-Object { $_.name -match '^MFCMAPI\.exe\..+\.zip$' })
$x64Assets = @($release.assets | Where-Object { $_.name -match '^MFCMAPI\.x64\.exe\..+\.zip$' })
if ($x86Assets.Count -ne 1 -or $x64Assets.Count -ne 1) {
    throw "Expected exactly one x86 and one x64 MFCMAPI executable ZIP in release '$($release.tag_name)'; found $($x86Assets.Count) x86 and $($x64Assets.Count) x64."
}
foreach ($asset in @($x86Assets[0], $x64Assets[0])) {
    $url = [uri]$asset.browser_download_url
    if (-not $url.IsAbsoluteUri -or $url.Scheme -ne "https" -or $url.Host -ne "github.com") {
        throw "Unexpected release asset URL '$($asset.browser_download_url)'."
    }
}

$OutputDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
if (Test-Path -LiteralPath $OutputDirectory) {
    throw "Output directory '$OutputDirectory' already exists. Choose a new directory to avoid overwriting manifests."
}
$null = New-Item -ItemType Directory -Path $OutputDirectory

Write-Host "Generating Microsoft.MFCMAPI $version from $($release.html_url)"
Write-Host "Manifests will be retained in $OutputDirectory"
$updateArguments = @(
    "update", "Microsoft.MFCMAPI",
    "--version", $version,
    "--out", $OutputDirectory,
    "--urls", "$($x86Assets[0].browser_download_url)|x86|user", "$($x64Assets[0].browser_download_url)|x64|user"
)
& $wingetCreate @updateArguments
if ($LASTEXITCODE -ne 0) {
    throw "WingetCreate update failed with exit code $LASTEXITCODE. Output directory: $OutputDirectory"
}

$manifestDirectory = Join-Path $OutputDirectory "manifests\m\Microsoft\MFCMAPI\$version"
if (-not (Test-Path -LiteralPath $manifestDirectory -PathType Container)) {
    throw "WingetCreate did not produce the expected manifest directory '$manifestDirectory'."
}
& $winget validate --manifest $manifestDirectory
if ($LASTEXITCODE -ne 0) {
    throw "WinGet validation failed with exit code $LASTEXITCODE. Review '$manifestDirectory'."
}

if ($Submit) {
    Write-Host "Submitting validated manifests. WingetCreate may prompt for GitHub OAuth sign-in."
    & $wingetCreate submit $manifestDirectory
    if ($LASTEXITCODE -ne 0) {
        throw "WingetCreate submit failed with exit code $LASTEXITCODE. Manifests remain at '$manifestDirectory'."
    }
}
else {
    Write-Host "Validated manifests: $manifestDirectory"
    Write-Host "No PR was created. Review the YAML, then submit it with:"
    Write-Host "wingetcreate submit `"$manifestDirectory`""
}
