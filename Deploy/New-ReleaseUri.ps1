function New-ReleaseUri {
	<#
	.SYNOPSIS
	Creates a Uri for a Release for upload
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
	.PARAMETER Draft
	Create release as a draft
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
	[string]$Release,
	[Parameter(Mandatory=$True)]
	[bool]$Draft
	)
	process
	{
	Write-Host "Building release for $Version"
	$releaseData = @{
		tag_name = "$Version";
		target_commitish = "master";
		name = "$Release ($Version)";
		body = $releaseNotes;
		draft = $Draft;
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

	[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"
	$result = Invoke-RestMethod @releaseParams
	$uploadUri = $result | Select -ExpandProperty upload_url

	return $uploadUri
	}
}