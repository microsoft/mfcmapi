$indir = "$Env:SYSTEM_ARTIFACTSDIRECTORY\$Env:BUILD_DEFINITIONNAME\Archives"
$deploydir = "$Env:SYSTEM_DEFAULTWORKINGDIRECTORY/$Env:BUILD_DEFINITIONNAME\Deploy"
$version = $Env:BUILD_BUILDNUMBER
$project = "MFCMAPI"
$release = $($args[0])
$username = $($args[1])
$password = $($args[2])

Write-Host "Release=$release"
Write-Host "Username=$username"
Write-Host "Password=$password"
Write-Host "Version=$version"

$secstr = New-Object -TypeName System.Security.SecureString
$password.ToCharArray() | ForEach-Object {$secstr.AppendChar($_)}
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $secstr

Add-Type -Path "$deploydir\CodePlex.WebServices.Client.dll"

function Build-FileName {
  <#
  .SYNOPSIS
  Build a full filename from file and version
  .EXAMPLE
  Build-FileName -FileName "MFCMAPI.exe" -Version "16.0.0.1044"
  .PARAMETER FileName
  Name of the file
  .PARAMETER Version
  Version of the file.
  #>
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$True)]
    [string]$FileName,
    [Parameter(Mandatory=$True)]
    [string]$Version
    )
  process
  {
    return "$FileName.$Version.zip"
  }
}

function Build-ReleaseFile {
  <#
  .SYNOPSIS
  Creates a ReleaseFile for upload
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Build-ReleaseFile("MFCMAPI 32 bit executable", "MFCMapi.exe", sourcepath, Version, release, true);
  .PARAMETER Name
  Release name of the file.
  .PARAMETER FileName
  Name of the file
  .PARAMETER Sourcepath
  Location of the file.
  .PARAMETER Version
  Version of the file.
  .PARAMETER Release
  Name of the release
  #>
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$True)]
    [string]$Name,
    [Parameter(Mandatory=$True)]
    [string]$FileName,
    [Parameter(Mandatory=$True)]
    [string]$Sourcepath,
    [Parameter(Mandatory=$True)]
    [string]$Version,
    [Parameter(Mandatory=$True)]
    [string]$Release)
  process
  {
    Write-Host "Building release file for $FileName"
    $fullFileName = "$FileName.$Version.zip"
    $releaseFile = New-Object CodePlex.WebServices.Client.ReleaseFile
    $releaseFile.Name = "$Name - $Release ($Version)"
    $releaseFile.FileName = Build-FileName -FileName $FileName -Version $Version
    $releaseFile.FileType = [CodePlex.WebServices.Client.ReleaseFileType]::RuntimeBinary
    $releaseFile.FileData = [System.IO.File]::ReadAllBytes("$sourcepath\$fullFileName")
    Write-Host "Release file built"

    return $releaseFile
  }
}

function Upload-ReleaseFile {
  <#
  .SYNOPSIS
  Creates a ReleaseFile for upload
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Upload-ReleaseFile(releaseService, releaseFile, true);
  .PARAMETER ReleaseService
  Release service object
  .PARAMETER ReleaseFile
  Release file object for upload
  .PARAMETER Default
  Upload is default
  #>
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$True)]
    [CodePlex.WebServices.Client.ReleaseService]$ReleaseService,
    [Parameter(Mandatory=$True)]
    [CodePlex.WebServices.Client.ReleaseFile]$ReleaseFile,
    [switch]$Default
  )
  process
  {
    $releaseFiles = New-Object System.Collections.Generic.List[CodePlex.WebServices.Client.ReleaseFile]
    $releaseFiles.Add($ReleaseFile)

    Write-Host "Uploading $($ReleaseFile.FileName)"

    if ($Default)
    {
      Write-Host "Setting $($ReleaseFile.FileName) as default file"
      $releaseService.UploadReleaseFiles($project, $release, $releaseFiles, $ReleaseFile.FileName)
    }
    else
    {
      $releaseService.UploadReleaseFiles($project, $release, $releaseFiles)
    }
  }
}

$releaseNotes = "Build: *$version*

Full release notes at [url:SGriffin's blog|FIXURL].

If you just want to run the MFCMAPI or MrMAPI, get the executables. If you want to debug them, get the symbol files and the [url:source.|https://github.com/stephenegriffin/mfcmapi]

*The 64 bit builds will only work on a machine with Outlook 2010/2013/2016 64 bit installed. All other machines should use the 32 bit builds, regardless of the operating system.*

[image:Facebook Badge|http://badge.facebook.com/badge/26764016480.2776.1538253884.png|http://www.facebook.com/pages/MFCMAPI/26764016480]"

$releaseService = New-Object CodePlex.WebServices.Client.ReleaseService
$releaseService.Timeout = 360000 # 6 minutes
$releaseService.Credentials = $cred

Write-Host "Creating $project/$release"
$id = $releaseService.CreateARelease($project, $release, $releaseNotes, $null, [CodePlex.WebServices.Client.ReleaseStatus]::Planned, $False, $False)
Write-Host "New project id is $id"

$releaseFile = Build-ReleaseFile -Name "MFCMAPI 32 bit symbol" -FileName "MFCMapi.pdb" -Sourcepath $indir -Version $version -Release $release
Upload-ReleaseFile -ReleaseService $releaseService -ReleaseFile $releaseFile
$releaseFile = Build-ReleaseFile -Name "MFCMAPI 64 bit executable" -FileName "MFCMapi.exe.x64" -Sourcepath $indir -Version $version -Release $release
Upload-ReleaseFile -ReleaseService $releaseService -ReleaseFile $releaseFile
$releaseFile = Build-ReleaseFile -Name "MFCMAPI 64 bit symbol" -FileName "MFCMapi.pdb.x64" -Sourcepath $indir -Version $version -Release $release
Upload-ReleaseFile -ReleaseService $releaseService -ReleaseFile $releaseFile

$releaseFile = Build-ReleaseFile -Name "MrMAPI 32 bit executable" -FileName "MrMAPI.exe" -Sourcepath $indir -Version $version -Release $release
Upload-ReleaseFile -ReleaseService $releaseService -ReleaseFile $releaseFile
$releaseFile = Build-ReleaseFile -Name "MrMAPI 32 bit symbol" -FileName "MrMAPI.pdb" -Sourcepath $indir -Version $version -Release $release
Upload-ReleaseFile -ReleaseService $releaseService -ReleaseFile $releaseFile
$releaseFile = Build-ReleaseFile -Name "MrMAPI 64 bit executable" -FileName "MrMAPI.exe.x64" -Sourcepath $indir -Version $version -Release $release
Upload-ReleaseFile -ReleaseService $releaseService -ReleaseFile $releaseFile
$releaseFile = Build-ReleaseFile -Name "MrMAPI 64 bit symbol" -FileName "MrMAPI.pdb.x64" -Sourcepath $indir -Version $version -Release $release
Upload-ReleaseFile -ReleaseService $releaseService -ReleaseFile $releaseFile

# This must be last or the default file setting will be overwritten by the other uploads
$releaseFile = Build-ReleaseFile -Name "MFCMAPI 32 bit executable" -FileName "MFCMapi.exe" -Sourcepath $indir -Version $version -Release $release
Upload-ReleaseFile -ReleaseService $releaseService -ReleaseFile $releaseFile -Default