$thisScript = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
. ($thisScript + '.\Compress-File.ps1')

$indir = "$Env:BUILD_SOURCESDIRECTORY\bin"
$outdir = "$Env:BUILD_ARTIFACTSTAGINGDIRECTORY\archives"
$version = $Env:BUILD_BUILDNUMBER

Write-Host "indir=$indir"
Write-Host "outdir=$outdir"
Write-Host "Version=$version"

if (!(Test-Path $outdir)) {New-Item -Path $outdir -Type Directory}

Get-ChildItem -Recurse $indir

Compress-File -Source "$indir\Win32\MrMAPI\Release\MrMAPI.exe" -Target "$outdir\MrMAPI.exe.$version.zip"
Compress-File -Source "$indir\Win32\MrMAPI\Release\MrMAPI.pdb" -Target "$outdir\MrMAPI.pdb.$version.zip"
Compress-File -Source "$indir\x64\MrMAPI\Release\MrMAPI.exe" -Target "$outdir\MrMAPI.exe.x64.$version.zip"
Compress-File -Source "$indir\x64\MrMAPI\Release\MrMAPI.pdb" -Target "$outdir\MrMAPI.pdb.x64.$version.zip"
Compress-File -Source "$indir\Win32\Release\MFCMAPI.exe" -Target "$outdir\MFCMAPI.exe.$version.zip"
Compress-File -Source "$indir\Win32\Release\MFCMAPI.pdb" -Target "$outdir\MFCMAPI.pdb.$version.zip"
Compress-File -Source "$indir\x64\Release\MFCMAPI.exe" -Target "$outdir\MFCMAPI.exe.x64.$version.zip"
Compress-File -Source "$indir\x64\Release\MFCMAPI.pdb" -Target "$outdir\MFCMAPI.pdb.x64.$version.zip"

Get-ChildItem -Recurse $outdir
