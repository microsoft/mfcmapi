$indir = "$Env:SYSTEM_ARTIFACTSDIRECTORY\$Env:BUILD_DEFINITIONNAME\Archives"
$version = $Env:BUILD_BUILDNUMBER
$gitHubUsername = 'stephenegriffin'
$gitHubRepository = 'mfcmapi'
$project = "MFCMAPI"
$release = $($args[0])
$gitHubApiKey = $($args[1])

Write-Host "Release=$release"
Write-Host "Version=$version"

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

function Create-Release {
  <#
  .SYNOPSIS
  Creates a ReleaseFile for upload
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Create-Release -gitHubUsername $gitHubUsername -gitHubRepository $gitHubRepository -Version $version -Release $release
  .PARAMETER gitHubUsername
  Github username.
  .PARAMETER gitHubRepository
  Github repository name
  .PARAMETER Version
  Version of the file.
  .PARAMETER Release
  Name of the release
  #>
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$True)]
    [string]$gitHubUsername,
    [Parameter(Mandatory=$True)]
    [string]$gitHubRepository,
    [Parameter(Mandatory=$True)]
    [string]$Version,
    [Parameter(Mandatory=$True)]
    [string]$Release
  )
  process
  {
	Write-Host "Building release for $Version"
	$releaseData = @{
	   tag_name = "$Version";
	   target_commitish = "master";
	   name = "$Release ($Version)";
	   body = $releaseNotes;
	   draft = $TRUE;
	   prerelease = $FALSE;
	}

	Write-Host "Release data built"

	$releaseParams = @{
	   Uri = "https://api.github.com/repos/$gitHubUsername/$gitHubRepository/releases";
	   Method = 'POST';
	   Headers = @{
		 Authorization = 'Basic ' + [Convert]::ToBase64String(
		 [Text.Encoding]::ASCII.GetBytes($gitHubApiKey + ":x-oauth-basic"));
	   }
	   ContentType = 'application/json';
	   Body = (ConvertTo-Json $releaseData -Compress)
	}

	$result = Invoke-RestMethod @releaseParams 
	$uploadUri = $result | Select -ExpandProperty upload_url

	return $uploadUri
  }
}

function Upload-ReleaseFile {
  <#
  .SYNOPSIS
  Creates a ReleaseFile for upload
  .DESCRIPTION
  Describe the function in more detail
  .EXAMPLE
  Upload-ReleaseFile -Name "MFCMAPI 32 bit executable" -FileName "MFCMapi.exe" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
  .PARAMETER Name
  Release name of the file.
  .PARAMETER FileName
  Name of the file
  .PARAMETER Sourcepath
  Location of the file.
  .PARAMETER Release
  Name of the release
  .PARAMETER Version
  Version of the file.
  .PARAMETER UploadURI
  Upload URI
  .PARAMETER draft
  Set to true to mark this as a draft release
  .PARAMETER Default
  Upload is default
  #>
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$True)]
    [string]$Name,
    [Parameter(Mandatory=$True)]
    [string]$FileName,
    [Parameter(Mandatory=$True)]
    [string]$Release,
    [Parameter(Mandatory=$True)]
    [string]$Sourcepath,
    [Parameter(Mandatory=$True)]
    [string]$Version,
    [Parameter(Mandatory=$True)]
    [string]$UploadURI,
    [Parameter(Mandatory=$False)]
    [switch]$Default
  )
  process
  {
    Write-Host "Uploading release file for $Name"

	$fullFileName = Build-FileName -FileName $FileName -Version $Version
	$uploadFile = Join-Path -path $Sourcepath -childpath $fullFileName
    $label = "$Name - $Release ($Version)"
	$uploadUri = $uploadUri -replace '\{\?name,label\}', "?name=$fullFileName&label=$label"
    Write-Host "Name = $Name"
    Write-Host "fullFileName = $fullFileName"
    Write-Host "Label = $Label"
    Write-Host "uploadFile = $uploadFile"
    Write-Host "uploadUri = $uploadUri"

	$uploadParams = @{
	   Uri = $UploadURI;
	   Method = 'POST';
	   Headers = @{
		 Authorization = 'Basic ' + [Convert]::ToBase64String(
		 [Text.Encoding]::ASCII.GetBytes($gitHubApiKey + ":x-oauth-basic"));
	   }
	   ContentType = 'application/zip';
	   InFile = $uploadFile
	}

	$result = Invoke-RestMethod @uploadParams    
  }
}

$releaseNotes = "Build: *$version*

Full release notes at [SGriffin's blog](FIXURL).

If you just want to run the MFCMAPI or MrMAPI, get the executables. If you want to debug them, get the symbol files and the [source](https://github.com/stephenegriffin/mfcmapi).

*The 64 bit builds will only work on a machine with Outlook 2010/2013/2016 64 bit installed. All other machines should use the 32 bit builds, regardless of the operating system.*

[![Facebook Badge](http://badge.facebook.com/badge/26764016480.2776.1538253884.png)](http://www.facebook.com/MFCMAPI)"

Write-Host "Creating $project/$release"

$uploadUri = Create-Release -gitHubUsername $gitHubUsername -gitHubRepository $gitHubRepository -Version $version -Release $release
Write-Host $uploadUri

Upload-ReleaseFile -Name "MFCMAPI 32 bit executable" -FileName "MFCMapi.exe" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
Upload-ReleaseFile -Name "MFCMAPI 32 bit symbol" -FileName "MFCMapi.pdb" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
Upload-ReleaseFile -Name "MFCMAPI 64 bit executable" -FileName "MFCMapi.exe.x64" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
Upload-ReleaseFile -Name "MFCMAPI 64 bit symbol" -FileName "MFCMapi.pdb.x64" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri

Upload-ReleaseFile -Name "MrMAPI 32 bit executable" -FileName "MrMAPI.exe" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
Upload-ReleaseFile -Name "MrMAPI 32 bit symbol" -FileName "MrMAPI.pdb" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
Upload-ReleaseFile -Name "MrMAPI 64 bit executable" -FileName "MrMAPI.exe.x64" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri
Upload-ReleaseFile -Name "MrMAPI 64 bit symbol" -FileName "MrMAPI.pdb.x64" -Release $release -Sourcepath $indir -Version $version -UploadURI $uploadUri

