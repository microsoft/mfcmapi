param(
    [string]$SdkVersion = "10.0.22621.0",
    [string]$DownloadUrl = "https://go.microsoft.com/fwlink/?linkid=2196241",
    [string]$ExpectedHash = "73FE3CC0E50D946D0C0A83A1424111E60DEE23F0803E305A8974A963B58290C0"
)

$ErrorActionPreference = "Stop"

$sdkPath = Join-Path "${env:ProgramFiles(x86)}\Windows Kits\10\Include" $SdkVersion
if (Test-Path $sdkPath) {
    Write-Host "Windows SDK $SdkVersion already installed"
    exit 0
}

$installer = Join-Path $env:TEMP "winsdksetup.exe"
$logPath = Join-Path $env:TEMP "sdk_install.log"

Write-Host "Downloading Windows 11 SDK $SdkVersion..."
Invoke-WebRequest -Uri $DownloadUrl -OutFile $installer

$actualHash = (Get-FileHash -Path $installer -Algorithm SHA256).Hash
if ($actualHash -ne $ExpectedHash) {
    Write-Error "SHA256 hash mismatch! Expected: $ExpectedHash, Got: $actualHash"
    exit 1
}
Write-Host "SHA256 verified: $actualHash"

Write-Host "Installing SDK (this may take a few minutes)..."
$proc = Start-Process -FilePath $installer -ArgumentList "/features OptionId.DesktopCPPx64 OptionId.DesktopCPPx86 OptionId.DesktopCPParm64 /quiet /norestart /log `"$logPath`"" -Wait -PassThru
if ($proc.ExitCode -ne 0) {
    Get-Content $logPath -ErrorAction SilentlyContinue | Select-Object -Last 50
    Write-Error "Windows SDK installer exited with code $($proc.ExitCode)"
    exit 1
}

if (!(Test-Path $sdkPath)) {
    Get-Content $logPath -ErrorAction SilentlyContinue | Select-Object -Last 50
    Write-Error "Windows SDK installation failed"
    exit 1
}

Write-Host "Windows SDK $SdkVersion installed successfully"
