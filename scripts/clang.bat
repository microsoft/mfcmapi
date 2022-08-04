pushd ..
for /r %%g in (*.h) do "%VCINSTALLDIR%\Tools\Llvm\bin\clang-format.exe" -i %%g
for /r %%g in (*.cpp) do "%VCINSTALLDIR%\Tools\Llvm\bin\clang-format.exe" -i %%g
popd