param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$VSTestArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$pf86 = ${env:ProgramFiles(x86)}
$vswhere = Join-Path $pf86 "Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path -Path $vswhere)) {
    throw "vswhere.exe not found at '$vswhere'. Install Visual Studio Build Tools or Visual Studio."
}

$vsPath = & $vswhere -latest -requires Microsoft.VisualStudio.PackageGroup.TestTools.Core -property installationPath
if (-not $vsPath) {
    throw "Visual Studio installation path with Test Tools not found via vswhere."
}

$vstest = Join-Path $vsPath "Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe"
if (-not (Test-Path -Path $vstest)) {
    throw "vstest.console.exe not found at '$vstest'."
}

if (-not $VSTestArgs -or $VSTestArgs.Count -eq 0) {
    throw "No VSTest arguments were provided."
}

$srcRoot = Join-Path -Path $PSScriptRoot -ChildPath ".."
Push-Location $srcRoot
try {
    & $vstest @VSTestArgs
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
finally {
    Pop-Location
}