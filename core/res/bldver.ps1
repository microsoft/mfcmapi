$Env:BUILD_BUILDNUMBER -match "(\d+)\.(\d+)\.(\d+)\.(\d+)"
$m = $matches
$file = gci "$Env:BUILD_SOURCESDIRECTORY\core\res\bldver.rc"
if($file)
{
 attrib $file -r
 $fc = Get-Content($file)
 $fc = $fc -replace "#define rmj (\d+)", ("#define rmj "+$m[1])
 $fc = $fc -replace "#define rmm (\d+)", ("#define rmm "+$m[2])
 $fc = $fc -replace "#define rup (\d+)", ("#define rup "+$m[3])
 $fc = $fc -replace "#define rmn (\d+)", ("#define rmn "+$m[4])
 $fc | Out-File $file
}