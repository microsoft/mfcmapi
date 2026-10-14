param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$MSBuildArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$pf86 = ${env:ProgramFiles(x86)}
$vswhere = Join-Path $pf86 "Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path -Path $vswhere)) {
    throw "vswhere.exe not found at '$vswhere'. Install Visual Studio Build Tools or Visual Studio."
}

$vsWhereArgs = @(
    "-latest"
    "-prerelease"
    "-version"
    "[18.0,19.0)"
    "-requires"
    "Microsoft.Component.MSBuild"
    "Microsoft.VisualStudio.Component.VC.ATLMFC.Spectre"
    "-property"
    "installationPath"
)
$vsPath = & $vswhere @vsWhereArgs
if (-not $vsPath) {
    throw "Visual Studio 2026 with MSBuild and Spectre-mitigated MFC was not found via vswhere."
}

$msbuild = Join-Path $vsPath "MSBuild\Current\Bin\amd64\msbuild.exe"
if (-not (Test-Path -Path $msbuild)) {
    throw "MSBuild executable not found at '$msbuild'."
}

if (-not $MSBuildArgs -or $MSBuildArgs.Count -eq 0) {
    throw "No MSBuild arguments were provided."
}

$srcRoot = Join-Path -Path $PSScriptRoot -ChildPath ".."
Push-Location $srcRoot
try {
    & $msbuild @MSBuildArgs
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
finally {
    Pop-Location
}