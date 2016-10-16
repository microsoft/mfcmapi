$outdir = "$Env:SYSTEM_ARTIFACTSDIRECTORY\out"
$version = $Env:BUILD_BUILDNUMBER
$release = $($args[0])
$username = $($args[1])
$password = $($args[2])

Write-Host "Release=$release"
Write-Host "Username=$username"
Write-Host "Password=$password"
Write-Host "Version=$version"

if (Test-Path $outdir) {Remove-Item -Path $outdir -Recurse}
if (!(Test-Path $outdir)) {New-Item -Path $outdir -Type Directory}

Compress-Archive .\MrMAPI\Win32\MrMAPI.exe $outdir\MrMAPI.exe.$version.zip
Compress-Archive .\MrMAPI\Win32\MrMAPI.pdb $outdir\MrMAPI.pdb.$version.zip
Compress-Archive .\MrMAPI\x64\MrMAPI.exe $outdir\MrMAPI.exe.x64.$version.zip
Compress-Archive .\MrMAPI\x64\MrMAPI.exe $outdir\MrMAPI.pdb.x64.$version.zip
Compress-Archive .\Release\Win32\MFCMAPI.exe $outdir\MFCMAPI.exe.$version.zip
Compress-Archive .\Release\Win32\MFCMAPI.pdb $outdir\MFCMAPI.pdb.$version.zip
Compress-Archive .\Release\x64\MFCMAPI.exe $outdir\MFCMAPI.exe.x64.$version.zip
Compress-Archive .\Release\x64\MFCMAPI.exe $outdir\MFCMAPI.pdb.x64.$version.zip
