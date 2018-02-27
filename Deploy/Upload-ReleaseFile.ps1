$thisScript = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
. ($thisScript + '.\New-VersionedFileName.ps1')

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

	$fullFileName = New-VersionedFileName -FileName $FileName -Version $Version
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

	[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"
	$result = Invoke-RestMethod @uploadParams
	}
}