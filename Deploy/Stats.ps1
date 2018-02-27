$gitHubUsername = 'stephenegriffin'
$gitHubRepository = 'mfcmapi'

function Get-ReleaseStats {
	<#
	.SYNOPSIS
	Gets stats for a release
	.DESCRIPTION
	Describe the function in more detail
	.EXAMPLE
	Stats-Release
	#>
	[CmdletBinding()]
	param
	(
	)
	process
	{
	Write-Host "Getting stats"

	$releaseStats = @{
		Uri = "https://api.github.com/repos/$gitHubUsername/$gitHubRepository/releases/latest";
		Method = 'GET';
		ContentType = 'application/json';
	}

	[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"
	$result = Invoke-RestMethod @releaseStats

	return $result
	}
}

$stats = Get-ReleaseStats
$stats.assets |fl name,download_count