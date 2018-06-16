pushd ..
for /r %%g in (*.h) do "%VS150COMNTOOLS%\vc\vcpackages\clang-format.exe" -i %%g
for /r %%g in (*.cpp) do "%VS150COMNTOOLS%\vc\vcpackages\clang-format.exe" -i %%g
popd