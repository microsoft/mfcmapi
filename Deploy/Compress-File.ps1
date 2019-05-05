function Compress-File {
  <#
  .SYNOPSIS
  Compress a file with logging
  .EXAMPLE
   Compress -Source "inpath\MrMAPI.exe" -Target "outpath\MrMAPI.exe.$version.zip"
  .PARAMETER Source
  Name of the source file
  .PARAMETER Target
  Name of the target file
  #>
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$True)]
    [string]$Source,
    [Parameter(Mandatory=$True)]
    [string]$Target
    )
  process
  {
    if (Test-Path $Source)
    {
      Write-Host "Compressing $Source to $Target"
      Compress-Archive $Source $Target
    }
    else
    {
      Write-Host "$Source not found: skipping"
    }
  }
}