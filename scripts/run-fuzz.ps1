param(
    [ValidateSet("x64", "x86")]
    [string]$Arch = "x64",
    [int]$MaxTotalTime = 60,
    [string]$CorpusPath = "fuzz\corpus",
    [string]$ArtifactPrefix = "fuzz\artifacts\",
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$AdditionalArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$srcRoot = Join-Path -Path $PSScriptRoot -ChildPath ".."
$fullCorpusPath = [System.IO.Path]::GetFullPath((Join-Path $srcRoot $CorpusPath))
$fullArtifactPrefix = [System.IO.Path]::GetFullPath((Join-Path $srcRoot $ArtifactPrefix))

if (-not $fullArtifactPrefix.EndsWith("\") -and -not $fullArtifactPrefix.EndsWith("/")) {
    $fullArtifactPrefix = "$fullArtifactPrefix\"
}

if (-not (Test-Path -Path $fullCorpusPath)) {
    throw "Corpus path not found: '$fullCorpusPath'. Run 'npm run fuzz:corpus' first."
}

$pf86 = ${env:ProgramFiles(x86)}
$vswhere = Join-Path $pf86 "Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path -Path $vswhere)) {
    throw "vswhere.exe not found at '$vswhere'. Install Visual Studio Build Tools or Visual Studio."
}

$vsPath = & $vswhere -latest -property installationPath
if (-not $vsPath) {
    throw "Visual Studio installation path not found via vswhere."
}

$msvcRoot = Join-Path $vsPath "VC\Tools\MSVC"
if (-not (Test-Path -Path $msvcRoot)) {
    throw "MSVC tools directory not found at '$msvcRoot'."
}

$toolVersion = (Get-ChildItem -Path $msvcRoot -Directory | Sort-Object Name -Descending | Select-Object -First 1).Name
if (-not $toolVersion) {
    throw "No MSVC toolset version found under '$msvcRoot'."
}

$runtimeArch = if ($Arch -eq "x64") { "x64" } else { "x86" }
$sanitizerRuntimePath = Join-Path $msvcRoot "$toolVersion\bin\Hostx64\$runtimeArch"
if (-not (Test-Path -Path $sanitizerRuntimePath)) {
    throw "MSVC runtime directory not found at '$sanitizerRuntimePath'."
}

$binPlatform = if ($Arch -eq "x64") { "x64" } else { "Win32" }
$fuzzExe = Join-Path $srcRoot "bin\$binPlatform\Fuzz\MFCMapi.exe"
if (-not (Test-Path -Path $fuzzExe)) {
    throw "Fuzz binary not found at '$fuzzExe'. Run 'npm run build:fuzz:$Arch' first."
}

$env:Path = "$sanitizerRuntimePath;$env:Path"

Push-Location $srcRoot
try {
    & $fuzzExe $fullCorpusPath "-artifact_prefix=$fullArtifactPrefix" "-max_total_time=$MaxTotalTime" "-print_final_stats=1" @AdditionalArgs
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
finally {
    Pop-Location
}