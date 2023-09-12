$vsRoot = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -property installationPath
$vcInstallDir = Join-Path $vsRoot "VC"
$clangdir = Join-Path $vcInstallDir "Tools\Llvm\bin\clang-format.exe"

# Check if VC install directory was found
if ($null -eq $vcInstallDir) {
    Write-Host "Visual C++ installation directory not found."
}

if ($null -eq $clangdir) {
    Write-Host "clang not found."
}

Push-Location ..

Write-Host "Formatting C++ headers"
Get-ChildItem -Recurse -Filter *.h | Where-Object { $_.DirectoryName -notlike "*include*" } | ForEach-Object {
    & $clangdir -i $_.FullName
}

Write-Host "Formatting C++ sources"
Get-ChildItem -Recurse -Filter *.cpp | ForEach-Object {
    & $clangdir -i $_.FullName
}

Pop-Location