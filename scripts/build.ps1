# Called from codeql.ps1 during database creation
$scriptroot = $PSScriptRoot
$srcRoot = Join-Path -Path $scriptroot -ChildPath ".."

$vsPath = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -property installationPath
$config = "Release_Unicode"
$platform = "x64"

Push-Location $srcRoot
& $vsPath\MSBuild\Current\Bin\amd64\msbuild.exe /m /p:Configuration="$config" /p:Platform="$platform" /t:Clean mfcmapi.sln
& $vsPath\MSBuild\Current\Bin\amd64\msbuild.exe /m /p:Configuration="$config" /p:Platform="$platform" mfcmapi.sln
Pop-Location