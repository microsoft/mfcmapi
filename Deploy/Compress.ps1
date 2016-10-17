$indir = "$Env:BUILD_SOURCESDIRECTORY\bin"
$outdir = "$Env:BUILD_ARTIFACTSTAGINGDIRECTORY\archives"
$version = $Env:BUILD_BUILDNUMBER

Write-Host "indir=$indir"
Write-Host "outdir=$outdir"
Write-Host "Version=$version"

if (!(Test-Path $outdir)) {New-Item -Path $outdir -Type Directory}

function Compress {
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
    if ((Test-Path $Source))
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

gci -Recurse $indir

Compress -Source "$indir\Win32\MrMAPI\MrMAPI.exe" -Target "$outdir\MrMAPI.exe.$version.zip"
Compress -Source "$indir\Win32\MrMAPI\MrMAPI.pdb" -Target "$outdir\MrMAPI.pdb.$version.zip"
Compress -Source "$indir\x64\MrMAPI\MrMAPI.exe" -Target "$outdir\MrMAPI.exe.x64.$version.zip"
Compress -Source "$indir\x64\MrMAPI\MrMAPI.exe" -Target "$outdir\MrMAPI.pdb.x64.$version.zip"
Compress -Source "$indir\Win32\Release\MFCMAPI.exe" -Target "$outdir\MFCMAPI.exe.$version.zip"
Compress -Source "$indir\Win32\Release\MFCMAPI.pdb" -Target "$outdir\MFCMAPI.pdb.$version.zip"
Compress -Source "$indir\x64\Release\MFCMAPI.exe" -Target "$outdir\MFCMAPI.exe.x64.$version.zip"
Compress -Source "$indir\x64\Release\MFCMAPI.exe" -Target "$outdir\MFCMAPI.pdb.x64.$version.zip"

gci -Recurse $outdir
