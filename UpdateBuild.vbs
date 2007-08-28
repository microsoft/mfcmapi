ForReading = 1
ForWriting = 2

set objFSO = CreateObject("Scripting.FileSystemObject")

'Clear the read only bits
Set f = objFSO.GetFile("build")
f.attributes = f.attributes - (f.attributes mod 2)

Set f = objFSO.GetFile("build.h")
f.attributes = f.attributes - (f.attributes mod 2)

set fBuild = objFSO.OpenTextFile ("build",ForReading)

buildno = fBuild.ReadAll
fBuild.close

buildno = buildno+1

set fBuild = objFSO.OpenTextFile ("build",ForWriting)
fBuild.WriteLine(buildno)
fBuild.close

set fBuildh = objFSO.OpenTextFile ("build.h",ForWriting)
fBuildh.Write("#define BUILDNO ")
fBuildh.WriteLine(buildno)

fBuildh.close
