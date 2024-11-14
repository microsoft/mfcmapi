# find clang-format under node-modules and run it on all C++ files in the project
$projectRoot = Join-Path $PSScriptRoot ".."
$clang = Join-Path $projectRoot "\node_modules\.bin\clang-format.ps1"
$style = Join-Path $projectRoot ".clang-format"

if ($null -eq $clang) {
    Write-Host "clang not found."
}
Write-Host "clang-format found at $clang"
& $clang --version

Write-Host "Style file found at $style"

Push-Location $projectRoot

Write-Host "Formatting C++ headers"
Get-ChildItem -Recurse -Filter *.h | Where-Object { $_.DirectoryName -notlike "*include*" } | ForEach-Object {
        & $clang --style=file:$style --verbose -i $_.FullName
}

Write-Host "Formatting C++ sources"
Get-ChildItem -Recurse -Filter *.cpp | ForEach-Object {
    & $clang --style=file:$style --verbose -i $_.FullName
}

Pop-Location