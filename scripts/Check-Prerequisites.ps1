# Check-Prerequisites.ps1
# Validates that the development environment has all required components

param(
    [switch]$Quiet,
    [switch]$Install
)

$script:errors = @()
$script:warnings = @()

function Write-Status {
    param([string]$Message, [string]$Status, [ConsoleColor]$Color = "White")
    if (-not $Quiet) {
        Write-Host "$Message " -NoNewline
        Write-Host $Status -ForegroundColor $Color
    }
}

function Test-VSInstallation {
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswhere)) {
        $script:errors += "Visual Studio is not installed"
        Write-Status "Visual Studio 2026:" "NOT FOUND" Red
        return $null
    }

    # Look for VS 2026 (version 18.x)
    $vs = & $vswhere -version "[18.0,19.0)" -requires Microsoft.VisualStudio.Workload.NativeDesktop -property installationPath -latest
    if (-not $vs) {
        # Check if any VS is installed
        $anyVs = & $vswhere -latest -property installationVersion
        if ($anyVs) {
            $script:errors += "Visual Studio 2026 with Desktop C++ workload is required (found VS $anyVs)"
        } else {
            $script:errors += "Visual Studio 2026 with Desktop C++ workload is required"
        }
        Write-Status "Visual Studio 2026:" "NOT FOUND" Red
        return $null
    }

    $version = & $vswhere -version "[18.0,19.0)" -property installationVersion -latest
    Write-Status "Visual Studio 2026:" "$version" Green
    return $vs
}

function Test-WindowsSDK {
    param([string]$RequiredVersion = "10.0.22621.0")
    
    $sdkPath = "${env:ProgramFiles(x86)}\Windows Kits\10\Include\$RequiredVersion"
    if (Test-Path $sdkPath) {
        Write-Status "Windows SDK $RequiredVersion`:" "FOUND" Green
        return $true
    }

    $script:errors += "Windows 11 SDK ($RequiredVersion) is not installed"
    Write-Status "Windows SDK $RequiredVersion`:" "NOT FOUND" Red
    return $false
}

function Test-VC145Toolset {
    param([string]$VSPath)
    
    if (-not $VSPath) { return $false }
    
    $toolsetPath = Join-Path $VSPath "VC\Tools\MSVC"
    if (Test-Path $toolsetPath) {
        $versions = Get-ChildItem $toolsetPath -Directory | Sort-Object Name -Descending
        if ($versions) {
            $latest = $versions[0].Name
            Write-Status "VC++ Toolset (v145):" "$latest" Green
            return $true
        }
    }

    $script:errors += "VC++ v145 toolset is not installed"
    Write-Status "VC++ Toolset (v145):" "NOT FOUND" Red
    return $false
}

function Test-SpectreLibraries {
    param([string]$VSPath)
    
    if (-not $VSPath) { return $false }
    
    $toolsetPath = Join-Path $VSPath "VC\Tools\MSVC"
    $versions = Get-ChildItem $toolsetPath -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending
    if (-not $versions) { return $false }
    
    $spectrePath = Join-Path $versions[0].FullName "lib\spectre"
    if (Test-Path $spectrePath) {
        Write-Status "Spectre-mitigated libs:" "FOUND" Green
        return $true
    }

    $script:warnings += "Spectre-mitigated libraries not installed (optional but recommended)"
    Write-Status "Spectre-mitigated libs:" "NOT FOUND" Yellow
    return $false
}

function Test-NodeJS {
    $node = Get-Command node -ErrorAction SilentlyContinue
    if ($node) {
        $version = & node --version
        Write-Status "Node.js:" "$version" Green
        return $true
    }

    $script:warnings += "Node.js not installed (optional - needed for npm build scripts)"
    Write-Status "Node.js:" "NOT FOUND (optional)" Yellow
    return $false
}

function Install-MissingComponents {
    if ($script:errors.Count -eq 0) {
        Write-Host "`nAll required components are installed!" -ForegroundColor Green
        return
    }

    Write-Host "`nMissing components detected." -ForegroundColor Yellow
    
    # Check if VS Installer can help
    $vsInstaller = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vs_installer.exe"
    if (Test-Path $vsInstaller) {
        Write-Host "`nYou can install missing components using the .vsconfig file:" -ForegroundColor Cyan
        Write-Host "  1. Open Visual Studio Installer"
        Write-Host "  2. Click 'More' > 'Import configuration'"
        Write-Host "  3. Select: $PSScriptRoot\..\.vsconfig"
        Write-Host ""
        Write-Host "Or run:" -ForegroundColor Cyan
        Write-Host "  & '$vsInstaller' modify --installPath '<VS Install Path>' --config '$PSScriptRoot\..\.vsconfig'"
    } else {
        Write-Host "`nPlease install Visual Studio 2026 from:" -ForegroundColor Cyan
        Write-Host "  https://visualstudio.microsoft.com/vs/"
    }
}

# Main
if (-not $Quiet) {
    Write-Host "Checking MFCMAPI build prerequisites...`n" -ForegroundColor Cyan
}

$vsPath = Test-VSInstallation
Test-WindowsSDK -RequiredVersion "10.0.22621.0" | Out-Null
Test-VC145Toolset -VSPath $vsPath | Out-Null
Test-SpectreLibraries -VSPath $vsPath | Out-Null
Test-NodeJS | Out-Null

if (-not $Quiet) {
    Write-Host ""
}

if ($script:errors.Count -gt 0) {
    if (-not $Quiet) {
        Write-Host "ERRORS:" -ForegroundColor Red
        $script:errors | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    }
    
    if ($Install) {
        Install-MissingComponents
    }
    
    exit 1
}

if ($script:warnings.Count -gt 0 -and -not $Quiet) {
    Write-Host "WARNINGS:" -ForegroundColor Yellow
    $script:warnings | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
}

if (-not $Quiet) {
    Write-Host "`nAll required prerequisites are installed!" -ForegroundColor Green
}

exit 0
