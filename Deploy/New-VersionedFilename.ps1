function New-VersionedFileName {
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