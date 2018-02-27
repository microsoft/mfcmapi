$thisScript = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
. ($thisScript + '.\Upload-ReleaseFile.ps1')
. ($thisScript + '.\New-ReleaseUri.ps1')

$indir = "$Env:SYSTEM_ARTIFACTSDIRECTORY\$Env:BUILD_DEFINITIONNAME\Archives"
$version = $Env:BUILD_BUILDNUMBER
$gitHubUsername = 'stephenegriffin'
$gitHubRepository = 'mfcmapi'
$release = $($args[0])
$gitHubApiKey = $($args[1])

Write-Host "Release=$release"
Write-Host "Version=$version"

$releaseNotes = "Build: *$version*

Full release notes at [SGriffin's blog](FIXURL).

If you just want to run the MFCMAPI or MrMAPI, get the executables. If you want to debug them, get the symbol files and the [source](https://github.com/stephenegriffin/mfcmapi).

*The 64 bit builds will only work on a machine with Outlook 2010/2013/2016 64 bit installed. All other machines should use the 32 bit builds, regardless of the operating system.*

[![Facebook Badge](http://badge.facebook.com/badge/26764016480.2776.1538253884.png)](http://www.facebook.com/MFCMAPI)

[Download stats (raw JSON)](https://api.github.com/repos/stephenegriffin/mfcmapi/releases/tags/$version)"

Write-Host "Creating $gitHubRepository/$release"

$uploadUri = New-ReleaseUri -gitHubUsername $gitHubUsername -gitHubRepository $gitHubRepository -Version $version -Release $release
Write-Host $uploadUri

Upload-ReleaseFile -Name "MFCMAPI 32 bit executable" -FileName "MFCMapi.exe" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
Upload-ReleaseFile -Name "MFCMAPI 32 bit symbol" -FileName "MFCMapi.pdb" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
Upload-ReleaseFile -Name "MFCMAPI 64 bit executable" -FileName "MFCMapi.exe.x64" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
Upload-ReleaseFile -Name "MFCMAPI 64 bit symbol" -FileName "MFCMapi.pdb.x64" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri

Upload-ReleaseFile -Name "MrMAPI 32 bit executable" -FileName "MrMAPI.exe" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
Upload-ReleaseFile -Name "MrMAPI 32 bit symbol" -FileName "MrMAPI.pdb" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
Upload-ReleaseFile -Name "MrMAPI 64 bit executable" -FileName "MrMAPI.exe.x64" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
Upload-ReleaseFile -Name "MrMAPI 64 bit symbol" -FileName "MrMAPI.pdb.x64" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
