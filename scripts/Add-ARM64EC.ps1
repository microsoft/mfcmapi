<#
.SYNOPSIS
    Adds ARM64EC platform configurations to Visual Studio C++ projects.

.DESCRIPTION
    This script adds ARM64EC platform configurations to .vcxproj files and .sln files.
    ARM64EC configurations are based on x64 configurations with platform changed to ARM64EC.

.PARAMETER ProjectPath
    Path to the project root directory (defaults to parent of script directory)

.EXAMPLE
    .\Add-ARM64EC.ps1
#>

param(
    [string]$ProjectPath = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = "Stop"

Write-Host "Adding ARM64EC configurations to mfcmapi projects..." -ForegroundColor Cyan

# Projects to update (mapistub already has ARM64EC)
$projects = @(
    "MFCMapi.vcxproj",
    "core\core.vcxproj",
    "MrMapi\MrMAPI.vcxproj",
    "UnitTest\UnitTest.vcxproj",
    "exampleMapiConsoleApp\exampleMapiConsoleApp.vcxproj"
)

# Configurations to add ARM64EC for
$configurations = @(
    "Debug",
    "Debug_Unicode",
    "Release",
    "Release_Unicode",
    "Prefast",
    "Prefast_Unicode",
    "Fuzz"
)

function Add-ARM64ECToProject {
    param([string]$vcxprojPath)
    
    Write-Host "  Processing: $vcxprojPath" -ForegroundColor Yellow
    
    [xml]$xml = Get-Content $vcxprojPath
    
    # Use XmlNamespaceManager properly
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $nsMgr.AddNamespace("ms", "http://schemas.microsoft.com/developer/msbuild/2003")
    
    # Check if ARM64EC already exists
    $existingConfigs = $xml.SelectNodes("//ms:ProjectConfiguration[contains(@Include, 'ARM64EC')]", $nsMgr)
    if ($existingConfigs.Count -gt 0) {
        Write-Host "    ARM64EC already exists, skipping" -ForegroundColor Gray
        return
    }
    
    # Find the ProjectConfigurations ItemGroup
    $configGroup = $xml.SelectSingleNode("//ms:ItemGroup[@Label='ProjectConfigurations']", $nsMgr)
    
    # Add ARM64EC ProjectConfiguration entries based on x64
    foreach ($config in $configurations) {
        $x64Config = $xml.SelectSingleNode("//ms:ProjectConfiguration[@Include='$config|x64']", $nsMgr)
        if ($x64Config) {
            $newConfig = $x64Config.Clone()
            $newConfig.SetAttribute("Include", "$config|ARM64EC")
            $newConfig.SelectSingleNode("ms:Platform", $nsMgr).InnerText = "ARM64EC"
            $configGroup.AppendChild($newConfig) | Out-Null
        }
    }
    
    # Find and clone PropertyGroups with Label="Configuration" for x64, change to ARM64EC
    $x64PropGroups = $xml.SelectNodes("//ms:PropertyGroup[@Label='Configuration' and contains(@Condition, '|x64')]", $nsMgr)
    foreach ($pg in $x64PropGroups) {
        $newPg = $pg.Clone()
        $condition = $newPg.GetAttribute("Condition")
        $newCondition = $condition -replace '\|x64', '|ARM64EC'
        $newPg.SetAttribute("Condition", $newCondition)
        $pg.ParentNode.InsertAfter($newPg, $pg) | Out-Null
    }
    
    # Find and clone ItemDefinitionGroups for x64, change to ARM64EC
    $x64ItemDefs = $xml.SelectNodes("//ms:ItemDefinitionGroup[contains(@Condition, '|x64')]", $nsMgr)
    foreach ($idg in $x64ItemDefs) {
        $newIdg = $idg.Clone()
        $condition = $newIdg.GetAttribute("Condition")
        $newCondition = $condition -replace '\|x64', '|ARM64EC'
        $newIdg.SetAttribute("Condition", $newCondition)
        
        # Update any output/intermediate directory paths that reference x64
        $allNodes = $newIdg.SelectNodes(".//*", $nsMgr)
        # Note: Most paths use $(Platform) which will automatically resolve to ARM64EC
        
        $idg.ParentNode.InsertAfter($newIdg, $idg) | Out-Null
    }
    
    # Find and clone PropertyGroups without Label that have x64 conditions (output dirs, etc)
    $x64OutputProps = $xml.SelectNodes("//ms:PropertyGroup[not(@Label) and contains(@Condition, '|x64')]", $nsMgr)
    foreach ($pg in $x64OutputProps) {
        $newPg = $pg.Clone()
        $condition = $newPg.GetAttribute("Condition")
        $newCondition = $condition -replace '\|x64', '|ARM64EC'
        $newPg.SetAttribute("Condition", $newCondition)
        $pg.ParentNode.InsertAfter($newPg, $pg) | Out-Null
    }
    
    # Save with UTF-8 encoding (with BOM for VS compatibility)
    $utf8WithBom = New-Object System.Text.UTF8Encoding($true)
    $writer = New-Object System.IO.StreamWriter($vcxprojPath, $false, $utf8WithBom)
    $xml.Save($writer)
    $writer.Close()
    
    Write-Host "    Added ARM64EC configurations" -ForegroundColor Green
}

function Add-ARM64ECToSolution {
    param([string]$slnPath)
    
    Write-Host "  Processing: $slnPath" -ForegroundColor Yellow
    
    $content = Get-Content $slnPath -Raw
    
    # Check if ARM64EC already exists
    if ($content -match "ARM64EC") {
        Write-Host "    ARM64EC already exists, skipping" -ForegroundColor Gray
        return
    }
    
    $lines = Get-Content $slnPath
    $newLines = @()
    $inSolutionConfig = $false
    $inProjectConfig = $false
    $solutionConfigsToAdd = @()
    $projectConfigsToAdd = @()
    
    foreach ($line in $lines) {
        $newLines += $line
        
        # Track sections
        if ($line -match "GlobalSection\(SolutionConfigurationPlatforms\)") {
            $inSolutionConfig = $true
        }
        elseif ($line -match "GlobalSection\(ProjectConfigurationPlatforms\)") {
            $inProjectConfig = $true
        }
        elseif ($line -match "EndGlobalSection" -and $inSolutionConfig) {
            # Insert ARM64EC solution configurations before EndGlobalSection
            foreach ($config in $configurations) {
                $newLines = $newLines[0..($newLines.Count - 2)] + @("`t`t$config|ARM64EC = $config|ARM64EC") + @($newLines[-1])
            }
            $inSolutionConfig = $false
        }
        elseif ($line -match "EndGlobalSection" -and $inProjectConfig) {
            $inProjectConfig = $false
        }
        
        # Clone x64 project configurations to ARM64EC
        if ($inProjectConfig -and $line -match "\.x64\.(ActiveCfg|Build\.0) = .+\|x64") {
            $arm64ecLine = $line -replace '\.x64\.', '.ARM64EC.' -replace '\|x64', '|ARM64EC'
            $newLines += $arm64ecLine
        }
    }
    
    # Write back
    $newLines | Set-Content $slnPath -Encoding UTF8
    Write-Host "    Added ARM64EC configurations" -ForegroundColor Green
}

# Process each project
foreach ($project in $projects) {
    $fullPath = Join-Path $ProjectPath $project
    if (Test-Path $fullPath) {
        Add-ARM64ECToProject -vcxprojPath $fullPath
    } else {
        Write-Host "  Not found: $fullPath" -ForegroundColor Red
    }
}

# Process solution file
$slnPath = Join-Path $ProjectPath "MFCMapi.sln"
if (Test-Path $slnPath) {
    Add-ARM64ECToSolution -slnPath $slnPath
} else {
    Write-Host "  Solution not found: $slnPath" -ForegroundColor Red
}

Write-Host ""
Write-Host "ARM64EC configuration complete!" -ForegroundColor Green
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Open solution in Visual Studio to verify" -ForegroundColor White
Write-Host "  2. Build ARM64EC configuration to test" -ForegroundColor White
Write-Host "  3. Run: npm run build:release:unicode:arm64ec" -ForegroundColor White
