$gitHubUsername = 'stephenegriffin'
$gitHubRepository = 'mfcmapi'
$project = "MFCMAPI"

function Stats-Release {
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

	$result = Invoke-RestMethod @releaseStats 

	return $result
  }
}

$stats = Stats-Release
$stats.assets |fl name,download_count

