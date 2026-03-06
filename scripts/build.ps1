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

$vsPath = & $vswhere -latest -requires Microsoft.Component.MSBuild -property installationPath
if (-not $vsPath) {
    throw "MSBuild installation path not found via vswhere."
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