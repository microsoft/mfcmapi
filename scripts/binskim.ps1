
# update nuget packages
nuget.exe restore packages.config -PackagesDirectory packages

$binskim = "packages\Microsoft.CodeAnalysis.BinSkim.1.9.5\tools\netcoreapp3.1\win-x64\BinSkim.exe"

& $binskim analyze C:\src\mfcmapi\MFCMapi.exe
& $binskim analyze C:\src\mfcmapi\bin\Win32\MrMAPI\Release\MrMAPI.exe
& $binskim analyze C:\src\mfcmapi\bin\Win32\MrMAPI\Release_Unicode\MrMAPI.exe
& $binskim analyze C:\src\mfcmapi\bin\Win32\Release\MFCMapi.exe
& $binskim analyze C:\src\mfcmapi\bin\Win32\Release_Unicode\MFCMapi.exe
& $binskim analyze C:\src\mfcmapi\bin\Win32\exampleMapiConsoleApp\Release\exampleMapiConsoleApp.exe
& $binskim analyze C:\src\mfcmapi\bin\Win32\exampleMapiConsoleApp\Release_Unicode\exampleMapiConsoleApp.exe
& $binskim analyze C:\src\mfcmapi\bin\x64\MrMAPI\Release\MrMAPI.exe
& $binskim analyze C:\src\mfcmapi\bin\x64\MrMAPI\Release_Unicode\MrMAPI.exe
& $binskim analyze C:\src\mfcmapi\bin\x64\Release\MFCMapi.exe
& $binskim analyze C:\src\mfcmapi\bin\x64\Release_Unicode\MFCMapi.exe
& $binskim analyze C:\src\mfcmapi\bin\x64\exampleMapiConsoleApp\Release\exampleMapiConsoleApp.exe
& $binskim analyze C:\src\mfcmapi\bin\x64\exampleMapiConsoleApp\Release_Unicode\exampleMapiConsoleApp.exe