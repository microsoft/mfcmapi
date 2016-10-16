$indir = "$Env:SYSTEM_ARTIFACTSDIRECTORY"
$outdir = "$Env:SYSTEM_ARTIFACTSDIRECTORY\out"
$version = $Env:BUILD_BUILDNUMBER

Write-Host "indir=$indir"
Write-Host "outdir=$outdir"
Write-Host "Version=$version"

if (Test-Path $outdir) {Remove-Item -Path $outdir -Recurse}
if (!(Test-Path $outdir)) {New-Item -Path $outdir -Type Directory}

Compress-Archive "$indir\Win32\MrMAPI\MrMAPI.exe" "$outdir\MrMAPI.exe.$version.zip"
Compress-Archive "$indir\Win32\MrMAPI\MrMAPI.pdb" "$outdir\MrMAPI.pdb.$version.zip"
Compress-Archive "$indir\x64\MrMAPI\MrMAPI.exe" "$outdir\MrMAPI.exe.x64.$version.zip"
Compress-Archive "$indir\x64\MrMAPI\MrMAPI.exe" "$outdir\MrMAPI.pdb.x64.$version.zip"
Compress-Archive "$indir\Win32\Release\MFCMAPI.exe" "$outdir\MFCMAPI.exe.$version.zip"
Compress-Archive "$indir\Win32\Release\MFCMAPI.pdb" "$outdir\MFCMAPI.pdb.$version.zip"
Compress-Archive "$indir\x64\Release\MFCMAPI.exe" "$outdir\MFCMAPI.exe.x64.$version.zip"
Compress-Archive "$indir\x64\Release\MFCMAPI.exe" "$outdir\MFCMAPI.pdb.x64.$version.zip"
