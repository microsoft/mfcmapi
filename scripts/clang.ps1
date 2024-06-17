$vsRoot = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -property installationPath
$vcInstallDir = Join-Path $vsRoot "VC"
$clang = Join-Path $vcInstallDir "Tools\Llvm\bin\clang-format.exe"

# Check if VC install directory was found
if ($null -eq $vcInstallDir) {
    Write-Host "Visual C++ installation directory not found."
}

if ($null -eq $clang) {
    Write-Host "clang not found."
}
Write-Host "clang-format found at $clang"
& $clang --version

Push-Location ..

Write-Host "Formatting C++ headers"
Get-ChildItem -Recurse -Filter *.h | Where-Object { $_.DirectoryName -notlike "*include*" } | ForEach-Object {
    & $clang -i $_.FullName
}

Write-Host "Formatting C++ sources"
Get-ChildItem -Recurse -Filter *.cpp | ForEach-Object {
    & $clang -i $_.FullName
}

Pop-Location