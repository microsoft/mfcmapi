# Visual Studio 2026 C++ Project Modernization Plan

## ?? Executive Summary

**Decisions Made:**
- ✅ **Toolset:** v145 (VS 2026 REQUIRED - no version selection)
- ✅ **Solution Version:** Update to VS 18
- ✅ **Windows SDK:** Pin to 10.0.22621.0 (Windows 11 SDK) with automatic installation
- ✅ **C++ Standard:** Keep stdcpplatest (uses C++23 features like std::byteswap)
- ✅ **Scope:** Apply ALL changes to both mfcmapi AND mapistub submodule
- ✅ **VS 2022 Support:** Dropped - VS 2026 REQUIRED
- ✅ **Implementation:** Use Directory.Build.props for centralized configuration
- ✅ **Centralized Settings:** SpectreMitigation, WarningLevel, TreatWarningAsError, SDLCheck, LanguageStandard
- ✅ **SDK Installation:** Automatic via .vsconfig + PowerShell script for CI/CD
- ✅ **Security Enhancements:** Enable ControlFlowGuard, GuardEHContMetadata, SpectreMitigation
- ✅ **Code Analysis:** Lightweight for Debug/Release, Full for Prefast configs
- ✅ **Prefast Strategy:** Separate CI quality gate, reproducible locally
- 🔲 **GitHub Actions:** Update all actions to latest versions (Phase 3)
- 🔲 **NuGet Packages:** Update all to latest versions (Phase 3)
- 🔲 **ARM64 Support:** To discuss after main migration

**Compatibility Matrix:**

| Component | Before (2022) | After (2026) | Breaking? |
|-----------|---------------|--------------|-----------|
| **Visual Studio** | 2022 (v143) | 2026 (v145) | ?? YES |
| **Solution Format** | VS 17 | VS 18 | ?? YES |
| **Windows SDK** | 10.0 (any) | 10.0.22621.0 | ?? Maybe* |
| **C++ Standard** | Default/C++17 | C++20 | ?? Maybe* |
| **GitHub Runners** | Any with VS | Requires VS 2026 | ?? YES |
| **VS Code** | Compatible | Compatible | ? NO |
| **Runtime Target** | Windows 10+ | Windows 10+ | ? NO |

*May require code changes depending on API usage

---

## Current State Analysis

### Project Info
- **Original Version**: Visual Studio 2022 (v143 toolset)
- **Solution**: MFCMapi.sln (Format Version 12.00, VS 17)
- **Projects**: MFCMapi (MFC App), core (Static Library), MrMAPI, UnitTest, mapistub, exampleMapiConsoleApp
- **Platform Targets**: Win32 (x86), x64
- **Current PlatformToolset**: v143 (VS 2022)
- **Current Windows SDK**: 10.0
- **Build Configurations**: Debug, Release, Debug_Unicode, Release_Unicode, Prefast, Prefast_Unicode, Fuzz
- **Current Status**: ? Clean build in VS 2026 with no errors

---

## ?? Modernization Tasks

### 1. Visual Studio Project Files (.vcxproj)

#### ?? Priority: HIGH

**Current State:**
- `PlatformToolset`: v143 (VS 2022)
- `WindowsTargetPlatformVersion`: 10.0
- `ToolsVersion`: 15.0 (in project files)
- All projects use Spectre-mitigated libraries ?

**Modernization Options:**

| Option | VS 2026 Only | Maintains VS 2022 Compatibility |
|--------|--------------|----------------------------------|
| **A: Upgrade to v145** | ? Use v145 toolset | ? VS 2022 won't recognize v145 |
| **B: Keep v143** | ?? Works but not "modern" | ? Full backward compatibility |
| **C: Conditional Toolset** | ? Best of both | ? Full backward compatibility |

**Recommended: Option A - Upgrade to v145 (Decision: Accepted)**

Create `Directory.Build.props` at solution root:

```xml
<?xml version="1.0" encoding="utf-8"?>
<Project>
  <PropertyGroup>
    <!-- Use VS 2026 toolset -->
    <PlatformToolset>v145</PlatformToolset>
    
    <!-- Update Windows SDK to Windows 11 SDK -->
    <WindowsTargetPlatformVersion>10.0.22621.0</WindowsTargetPlatformVersion>
    
    <!-- Set C++ Language Standard to C++20 -->
    <LanguageStandard>stdcpp20</LanguageStandard>
  </PropertyGroup>
</Project>
```

**Note:** This will also apply to the mapistub submodule if it doesn't have its own Directory.Build.props.

**Note:** This approach requires Visual Studio 2026. Projects will not open in VS 2022 without the v145 toolset.

**Alternative: Update each .vcxproj individually**

In each project file, replace all instances of:
```xml
<PlatformToolset>v143</PlatformToolset>
```

With:
```xml
<PlatformToolset Condition="'$(VisualStudioVersion)' == '17.0'">v143</PlatformToolset>
<PlatformToolset Condition="'$(VisualStudioVersion)' &gt;= '18.0'">v145</PlatformToolset>
```

**Files to Update (if not using Directory.Build.props):**
- [ ] `MFCMapi.vcxproj`
- [ ] `core\core.vcxproj`
- [ ] `MrMapi\MrMAPI.vcxproj`
- [ ] `UnitTest\UnitTest.vcxproj`
- [ ] `mapistub\mapistub.vcxproj`
- [ ] `exampleMapiConsoleApp\exampleMapiConsoleApp.vcxproj`

---

### 2. Solution File (.sln)

#### ?? Priority: HIGH

**Current State:**
```
Microsoft Visual Studio Solution File, Format Version 12.00
# Visual Studio Version 17
VisualStudioVersion = 17.7.34031.279
```

**Update to VS 2026:**
```
Microsoft Visual Studio Solution File, Format Version 12.00
# Visual Studio Version 18
VisualStudioVersion = 18.0.0.0
```

**Decision:** ? Update to Version 18

**Note:** Solution will be updated manually to indicate VS 2026 as the required version.

---

### 3. Windows SDK Version

#### ?? Priority: MEDIUM

**Current State:**
- `WindowsTargetPlatformVersion`: 10.0 (unversioned - uses latest installed)

**Options:**

| Approach | Pros | Cons |
|----------|------|------|
| **Keep 10.0** | Flexible, uses latest available | Less reproducible builds |
| **Specify 10.0.22621.0** | Reproducible, Windows 11 SDK | Requires SDK installation |
| **Specify 10.0.26100.0** | Latest features | May not be widely available yet |

**Recommendation:** Specify `10.0.22621.0` (Windows 11 SDK) for reproducible builds while maintaining broad compatibility.

Update in `Directory.Build.props`:
```xml
<WindowsTargetPlatformVersion>10.0.22621.0</WindowsTargetPlatformVersion>
```

---

### 4. C++ Language Standard

#### ?? Priority: HIGH

**Decision:** ? Upgrade to C++20 (stdcpp20)

**Current State:** 
- Need to verify if explicitly set or using default (likely C++17 or latest)

**Check current setting with:**
```powershell
Select-String -Path "*.vcxproj" -Pattern "LanguageStandard"
Select-String -Path "mapistub\*.vcxproj" -Pattern "LanguageStandard"
```

**New Setting:**

| Standard | VS 2026 Support | Recommendation |
|----------|-----------------|----------------|
| **C++20 (stdcpp20)** | ? Full | ? **SELECTED** - Modern, well-supported, great features |

**Add to Directory.Build.props (both mfcmapi and mapistub):**
```xml
<LanguageStandard>stdcpp20</LanguageStandard>
```

**C++20 Features Available:**
- Concepts
- Ranges
- Coroutines
- Modules (limited support)
- `constexpr` improvements
- `std::format`
- Three-way comparison operator (`<=>`)
- Designated initializers
- And much more...

**Migration Considerations:**
- ? Review code for deprecated features removed in C++20
- ? Test thoroughly after upgrade (especially templates and constexpr code)
- ? Check mapistub builds and tests
- ? Update CI/CD to note C++20 requirement
- ?? Some older third-party libraries may have issues (unlikely for this project)
- ? Document new C++20 features team can use

**Breaking Changes to Watch For:**
- More strict template instantiation rules
- Some implicit conversions removed
- Header cleanup (some standard headers no longer transitively include others)

**Testing After C++20 Upgrade:**
- [ ] All configurations compile
- [ ] All unit tests pass
- [ ] No new warnings
- [ ] Check for deprecated API usage
- [ ] Verify mapistub compatibility

---

### 5. Visual Studio Configuration (.vsconfig)

#### ?? Priority: HIGH

**Current Issues:**
- ? References old Windows 10 SDK: `Microsoft.VisualStudio.Component.Windows10SDK.18362`
- ?? Component IDs may have changed for VS 2026

**Updated .vsconfig for VS 2026:**

```json
{
  "version": "1.0",
  "components": [
    "Microsoft.Component.MSBuild",
    "Microsoft.VisualStudio.Component.ClassDesigner",
    "Microsoft.VisualStudio.Component.CodeMap",
    "Microsoft.VisualStudio.Component.CoreEditor",
    "Microsoft.VisualStudio.Component.Debugger.JustInTime",
    "Microsoft.VisualStudio.Component.GraphDocument",
    "Microsoft.VisualStudio.Component.Graphics.Tools",
    "Microsoft.VisualStudio.Component.IntelliCode",
    "Microsoft.VisualStudio.Component.IntelliTrace.FrontEnd",
    "Microsoft.VisualStudio.Component.Roslyn.Compiler",
    "Microsoft.VisualStudio.Component.SQL.LocalDB.Runtime",
    "Microsoft.VisualStudio.Component.TextTemplating",
    "Microsoft.VisualStudio.Component.UWP.VC.ARM64",
    "Microsoft.VisualStudio.Component.UWP.VC.ARM64EC",
    "Microsoft.VisualStudio.Component.VC.ATL",
    "Microsoft.VisualStudio.Component.VC.ATL.Spectre",
    "Microsoft.VisualStudio.Component.VC.ATLMFC",
    "Microsoft.VisualStudio.Component.VC.ATLMFC.Spectre",
    "Microsoft.VisualStudio.Component.VC.CoreIde",
    "Microsoft.VisualStudio.Component.VC.DiagnosticTools",
    "Microsoft.VisualStudio.Component.VC.Redist.14.Latest",
    "Microsoft.VisualStudio.Component.VC.Runtimes.ARM.Spectre",
    "Microsoft.VisualStudio.Component.VC.Runtimes.ARM64.Spectre",
    "Microsoft.VisualStudio.Component.VC.Runtimes.ARM64EC.Spectre",
    "Microsoft.VisualStudio.Component.VC.Runtimes.x86.x64.Spectre",
    "Microsoft.VisualStudio.Component.VC.Tools.ARM64",
    "Microsoft.VisualStudio.Component.VC.Tools.ARM64EC",
    "Microsoft.VisualStudio.Component.VC.Tools.x86.x64",
    "Microsoft.VisualStudio.Component.Windows11SDK.22621",
    "Microsoft.VisualStudio.ComponentGroup.ArchitectureTools.Native",
    "Microsoft.VisualStudio.ComponentGroup.NativeDesktop.Core",
    "Microsoft.VisualStudio.Workload.CoreEditor",
    "Microsoft.VisualStudio.Workload.NativeDesktop"
  ]
}
```

**Key Changes:**
- ? `Windows10SDK.18362` ? `Windows11SDK.22621`
- ? Kept all Spectre mitigation components
- ? Kept ARM64/ARM64EC support components

**Verify component IDs:**
```powershell
# List available components in your VS installation
& "C:\Program Files (x86)\Microsoft Visual Studio\Installer\vs_installer.exe" export --config .vsconfig
```

---

### 6. VS Code Configuration

#### ?? Priority: MEDIUM

**Current State:**
- `.vscode\tasks.json` uses direct MSBuild commands ?
- No hardcoded VS paths ?
- Uses `msbuild` from PATH ?

**Status:** ? **Compatible as-is**

**Decision:** ? No version selection needed - VS 2026 required

**Verification Steps:**
1. Open project in VS Code
2. Terminal ? Run Build Task (Ctrl+Shift+B)
3. Test each configuration
4. Verify MSBuild resolves to VS 2026

**Optional: Use vswhere to ensure VS 2026:**
```json
{
  "label": "Build with VS 2026",
  "type": "shell",
  "command": "powershell",
  "args": [
    "-Command",
    "$vsPath = & 'C:\\Program Files (x86)\\Microsoft Visual Studio\\Installer\\vswhere.exe' -version '[18.0,19.0)' -property installationPath; if (!$vsPath) { throw 'VS 2026 not found' }; & \"$vsPath\\MSBuild\\Current\\Bin\\amd64\\msbuild.exe\" /m /p:Configuration=Debug /p:Platform=x64 mfcmapi.sln"
  ]
}
```

---

### 7. GitHub Actions Workflows

#### ?? Priority: HIGH

**Decision:** ? Update ALL actions to latest stable versions

**Current Workflows:**
- `.github/workflows/github-ci.yml` - Main CI build
- `.github/workflows/codeql.yml` - CodeQL security analysis
- `.github/workflows/clang.yml` - Clang format checks (referenced in README)
- `.github/workflows/dependency-review.yml` - Dependency scanning
- `.github/workflows/devskim.yml` - Security scanning

**Action Version Audit:**

| Action | Current Version | Action Required |
|--------|----------------|----------------|
| `step-security/harden-runner` | v2.14.1 (SHA pinned) | ? Update to latest v2.x or v3.x |
| `actions/checkout` | v6.0.2 (SHA pinned) | ? Update to latest v6.x |
| `actions/upload-artifact` | v6.0.0 (SHA pinned) | ? Update to latest v6.x |
| `actions/download-artifact` | v7.0.0 (SHA pinned) | ? Update to latest v7.x |
| `github/codeql-action/*` | v3.29.5 (SHA pinned) | ? Update to latest v3.x |
| `EnricoMi/publish-unit-test-result-action` | v2.22.0 (SHA pinned) | ? Update to latest v2.x or v3.x |

**Update Process:**

1. **Check for latest versions:**
```bash
# For each action:
# https://github.com/step-security/harden-runner/releases
# https://github.com/actions/checkout/releases
# https://github.com/actions/upload-artifact/releases
# https://github.com/actions/download-artifact/releases
# https://github.com/github/codeql-action/releases
# https://github.com/EnricoMi/publish-unit-test-result-action/releases
```

2. **Update workflow with VS 2026 requirement:**

```yaml
jobs:
  build:
    runs-on: windows-latest
    strategy:
      matrix:
        configuration: [ 'Release', 'Debug', 'Release_Unicode', 'Debug_Unicode' ]
        platform: [ 'Win32', 'x64' ]

    steps:
    - name: Harden Runner
      uses: step-security/harden-runner@<latest-sha>  # UPDATE
      with:
        egress-policy: audit

    - name: Checkout repository
      uses: actions/checkout@<latest-sha>  # UPDATE
      with:
        submodules: 'recursive'
    
    - name: Install Windows SDK 10.0.22621.0
      shell: pwsh
      run: |
        # Verify SDK or install it
        $sdkPath = "C:\\Program Files (x86)\\Windows Kits\\10\\Include\\10.0.22621.0"
        if (!(Test-Path $sdkPath)) {
          Write-Host "Installing Windows SDK..."
          # Installation logic here
        }
    
    - name: "Build with VS 2026"
      shell: pwsh
      run: |
        # Require VS 2026
        $path = & "${env:ProgramFiles(x86)}\\Microsoft Visual Studio\\Installer\\vswhere.exe" -version '[18.0,19.0)' -property installationPath
        if (!$path) { throw "Visual Studio 2026 is required" }
        & $path\\MSBuild\\Current\\Bin\\amd64\\msbuild.exe /m /p:Configuration="${{matrix.configuration}}" /p:Platform="${{matrix.platform}}" mfcmapi.sln
```

**Action Items:**
- [ ] Update ALL actions to latest stable versions
- [ ] Update ALL SHA pins to match new versions
- [ ] Add Windows SDK installation step
- [ ] Ensure VS 2026 is required (not latest)
- [ ] Test all workflows after updates
- [ ] Verify test result publishing still works

---

### 8. NuGet Packages

#### ?? Priority: HIGH

**Decision:** ? Update ALL NuGet packages to latest stable versions

**Current Packages (from .vcxproj imports):**
```xml
<Import Project="packages\Microsoft.SourceLink.GitHub.1.0.0\build\..." />
<Import Project="packages\Microsoft.SourceLink.Common.1.0.0\build\..." />
<Import Project="packages\Microsoft.Build.Tasks.Git.1.0.0\build\..." />
```

**Package Updates Required:**

| Package | Current | Action Required |
|---------|---------|----------------|
| Microsoft.SourceLink.GitHub | 1.0.0 | ? Update to latest (check 8.0.0+) |
| Microsoft.SourceLink.Common | 1.0.0 | ? Update to latest (check 8.0.0+) |
| Microsoft.Build.Tasks.Git | 1.0.0 | ? Update to latest (check 8.0.0+) |

**Update Process:**

1. **Check for latest versions:**
```powershell
# In Visual Studio Package Manager Console
Get-Package | Format-Table Id, Version

# Check online
# https://www.nuget.org/packages/Microsoft.SourceLink.GitHub/
# https://www.nuget.org/packages/Microsoft.SourceLink.Common/
# https://www.nuget.org/packages/Microsoft.Build.Tasks.Git/
```

2. **Update all packages:**
```powershell
# Update all projects
Update-Package Microsoft.SourceLink.GitHub
Update-Package Microsoft.SourceLink.Common
Update-Package Microsoft.Build.Tasks.Git

# Or update all packages at once
Update-Package
```

3. **Verify updates:**
```powershell
Get-Package | Format-Table Id, Version
```

**Note:** PackageReference migration is on separate u/sgriffin/nuget branch

---

### 9. C++ Security & Compiler Settings

#### ?? Priority: HIGH

**Current Settings (Good!):**
- ? Spectre mitigation enabled: `<SpectreMitigation>Spectre</SpectreMitigation>`
- ? Static MFC linking: `<UseOfMfc>Static</UseOfMfc>`
- ? Using latest redistributables: `VC.Redist.14.Latest`
- ? AddressSanitizer enabled in Fuzz config: `<EnableASAN>true</EnableASAN>`
- ? Fuzzer enabled in Fuzz config: `<EnableFuzzer>true</EnableFuzzer>`

**Security Enhancements - REQUIRED:**

**Decision:** ? Enable all security features in Directory.Build.props

**Add to Directory.Build.props:**
```xml
<?xml version="1.0" encoding="utf-8"?>
<Project>
  <PropertyGroup>
    <PlatformToolset>v145</PlatformToolset>
    <WindowsTargetPlatformVersion>10.0.22621.0</WindowsTargetPlatformVersion>
    <LanguageStandard>stdcpp20</LanguageStandard>
    
    <!-- Security Features -->
    <ControlFlowGuard>Guard</ControlFlowGuard>
    <GuardEHContMetadata>true</GuardEHContMetadata>
  </PropertyGroup>
</Project>
```

**What these do:**

1. **ControlFlowGuard (CFG):**
   - Prevents control flow hijacking attacks
   - Validates indirect call targets at runtime
   - Industry best practice for Windows apps
   - Minimal performance impact

2. **GuardEHContMetadata:**
   - Enhances exception handling security
   - Prevents exception handler hijacking
   - Works with CFG
   - Recommended for all production code

**Action Items:**
- [ ] Add CFG and GuardEHContMetadata to Directory.Build.props
- [ ] Apply to both mfcmapi and mapistub
- [ ] Test all configurations
- [ ] Verify no performance degradation
- [ ] Document in security documentation

---

### 10. Code Analysis - REQUIRED

#### ?? Priority: HIGH

**Decision:** ? Enable C++ Core Check and all Microsoft code analysis

**Current State:**
- ? Prefast configurations available (will discuss separately)
- ? CodeQL enabled in GitHub Actions
- ? AddressSanitizer in Fuzz config
- ? DevSkim security scanning
- ? Dependency review workflow

**Required Enhancements:**

**Add to Directory.Build.props:**

```xml
<?xml version="1.0" encoding="utf-8"?>
<Project>
  <PropertyGroup>
    <PlatformToolset>v145</PlatformToolset>
    <WindowsTargetPlatformVersion>10.0.22621.0</WindowsTargetPlatformVersion>
    <LanguageStandard>stdcpp20</LanguageStandard>
    
    <!-- Security Features -->
    <ControlFlowGuard>Guard</ControlFlowGuard>
    <GuardEHContMetadata>true</GuardEHContMetadata>
    
    <!-- Code Analysis -->
    <EnableCppCoreCheck>true</EnableCppCoreCheck>
    <CodeAnalysisRuleSet>CppCoreCheckRules.ruleset</CodeAnalysisRuleSet>
    <EnableMicrosoftCodeAnalysis>true</EnableMicrosoftCodeAnalysis>
  </PropertyGroup>
</Project>
```

**What these enable:**

1. **EnableCppCoreCheck:**
   - Enforces C++ Core Guidelines
   - Modern C++ best practices
   - Memory safety checks
   - Type safety validation

2. **CodeAnalysisRuleSet:**
   - Specifies which rules to enforce
   - CppCoreCheckRules = comprehensive set
   - Can customize if needed

3. **EnableMicrosoftCodeAnalysis:**
   - Microsoft's static analysis
   - Complements C++ Core Check
   - Additional security and correctness checks

**Expected Warnings:**
- Will likely generate new warnings on first build
- Review and fix critical issues
- Can suppress false positives
- Document any suppressions

**Prefast Discussion:**
- ?? Will discuss Prefast configuration separately after main migration
- Prefast configs remain for now
- May integrate with new analysis settings

**Action Items:**
- [ ] Add analysis settings to Directory.Build.props
- [ ] Apply to both mfcmapi and mapistub
- [ ] Build and review new warnings
- [ ] Fix critical issues
- [ ] Document suppression strategy
- [ ] Schedule Prefast discussion

---

### 11. ARM64 Support

#### ?? Priority: DEFERRED

**Decision:** ?? **Discuss after main VS 2026 migration is complete**

**Current State:**
- .vsconfig includes ARM64 components ?
- No ARM64 platform configurations in .sln
- No ARM64 build tasks in VS Code

**Notes for Future Discussion:**
- Determine if ARM64 support is needed
- Evaluate Windows on ARM market
- Consider build/test infrastructure requirements
- Plan configuration addition if needed

**Action:** Schedule separate discussion after Phase 2 is complete

---

### 12. Submodule Considerations (mapistub)

#### ?? Priority: HIGH

**Submodule: mapistub (MAPIStubLibrary)**
- Location: `C:\src\mfcmapi\mapistub`
- Branch: main
- Remote: https://github.com/microsoft/MAPIStubLibrary

**Decision:** ? Apply full VS 2026 modernization to mapistub (same changes as main project)

**Action Items:**

1. **Create Directory.Build.props in mapistub/:**
```xml
<?xml version="1.0" encoding="utf-8"?>
<Project>
  <PropertyGroup>
    <PlatformToolset>v145</PlatformToolset>
    <WindowsTargetPlatformVersion>10.0.22621.0</WindowsTargetPlatformVersion>
    <LanguageStandard>stdcpp20</LanguageStandard>
  </PropertyGroup>
</Project>
```

2. **Update mapistub solution file:**
   - [ ] Update to Visual Studio Version 18
   - [ ] Update VisualStudioVersion to 18.0.0.0

3. **Create/Update .vsconfig in mapistub/:**
   - [ ] Add Windows 11 SDK 22621 component
   - [ ] Match component list from main project
   - [ ] Ensure Spectre mitigation components included

4. **Remove v143 references:**
   - [ ] Remove PlatformToolset from all .vcxproj files (will inherit from Directory.Build.props)
   - [ ] Or update to v145 if not using Directory.Build.props

5. **Add SDK installation script:**
   - [ ] Create mapistub/scripts/Install-RequiredSDKs.ps1
   - [ ] Match script from main project

6. **Update documentation:**
   - [ ] Update mapistub README with VS 2026 requirement
   - [ ] Document C++20 upgrade
   - [ ] Note automatic SDK installation via .vsconfig

7. **Test integration:**
   - [ ] Build mapistub standalone in VS 2026
   - [ ] Build mfcmapi with updated mapistub
   - [ ] Verify all configurations work
   - [ ] Run tests

8. **Consider upstream contribution:**
   - [ ] Fork microsoft/MAPIStubLibrary
   - [ ] Create PR with VS 2026 updates
   - [ ] Update submodule reference in mfcmapi after merge

**Submodule Update Commands:**
```bash
# Work in mapistub
cd mapistub

# Create branch for VS 2026 work
git checkout -b vs2026-upgrade

# Make changes (Directory.Build.props, .vsconfig, etc.)

# Commit changes
git add .
git commit -m "Upgrade to Visual Studio 2026 (v145, C++20, Windows 11 SDK)"

# If contributing upstream:
git push origin vs2026-upgrade
# Then create PR to microsoft/MAPIStubLibrary

# Return to main project
cd ..

# Update submodule reference
git add mapistub
git commit -m "Update mapistub submodule to VS 2026 version"
```

---

## 🚀 Migration Strategy

### Phase 1: Local VS 2026 Setup (Non-Breaking Prep)
**Timeline: First PR**

**Main Project (mfcmapi):**
- [x] Create `Directory.Build.props` with v145, C++20, SDK, security settings
- [x] Update `MFCMapi.sln` to VS 18
- [x] Update `.vsconfig` with Windows 11 SDK component
- [x] Remove v143 toolset references from ALL .vcxproj files:
  - [x] `MFCMapi.vcxproj`
  - [x] `core\core.vcxproj`
  - [x] `MrMapi\MrMAPI.vcxproj`
  - [x] `UnitTest\UnitTest.vcxproj`
  - [x] `exampleMapiConsoleApp\exampleMapiConsoleApp.vcxproj`
- [x] Centralize common settings in Directory.Build.props:
  - [x] Remove `WindowsTargetPlatformVersion` from all .vcxproj
  - [x] Remove `SpectreMitigation` from all .vcxproj  
  - [x] Remove `LanguageStandard` from all .vcxproj
  - [x] Remove `WarningLevel` from all .vcxproj
  - [x] Remove `TreatWarningAsError` from all .vcxproj
  - [x] Remove `SDLCheck` from all .vcxproj
- [ ] Test ALL configurations build locally in VS 2026

**Submodule (mapistub):**
- [x] Create `mapistub/Directory.Build.props` with same settings
- [x] Update `mapistub/mapistub.sln` to VS 18
- [x] Update `mapistub/.vsconfig` with Windows 11 SDK component
- [x] Remove v143 toolset references from `mapistub/mapistub.vcxproj`
- [x] Centralize common settings (same as main project)
- [ ] Test mapistub builds standalone in VS 2026

**Risk Level:** 🟡 Low-Medium (breaking change for VS 2022 users)

---

### Phase 1b: Node-gyp Setup for mfcmapi
**Timeline: After Phase 1 builds pass**

**Add node-gyp build support to mfcmapi (matching mapistub pattern):**
- [ ] Create `binding.gyp` for mfcmapi projects
- [ ] Update `package.json` with node-gyp build scripts:
  - [x] `build:x64`, `build:x86`, `build:arm64`
  - [x] `build:all`, `clean`, `clean:all`
- [x] ~~Test node-gyp builds work~~ (Modified: npm scripts with vswhere auto-detection instead)

**Risk Level:** 🟢 Low (additive, doesn't affect VS builds)
**Status:** ✅ COMPLETE

---

### Phase 2: Testing & Documentation
**Timeline: After Phase 1 builds pass**

- [x] Test ALL configurations build in VS 2026:
  - [x] Debug x86, x64
  - [x] Release x86, x64
  - [x] Debug_Unicode x86, x64
  - [x] Release_Unicode x86, x64
  - [x] Prefast x86, x64
  - [x] Fuzz x86, x64
- [x] Run ALL unit tests (251 passed)
- [x] Test VS Code builds work (F5 debugging configured)
- [x] Address any new warnings from C++20 or code analysis
- [x] Update README.md with VS 2026/C++20 requirements
- [x] Update CONTRIBUTING.md with setup instructions

**Risk Level:** 🟡 Medium (C++20 may require code fixes)
**Status:** ✅ COMPLETE

---

### Phase 3: CI/CD Updates (GitHub Actions)
**Timeline: AFTER local builds are verified**

- [ ] Update ALL GitHub Actions to latest versions with new SHA pins:
  - [ ] `step-security/harden-runner`
  - [ ] `actions/checkout`
  - [ ] `actions/upload-artifact`
  - [ ] `actions/download-artifact`
  - [ ] `github/codeql-action/*`
  - [ ] `EnricoMi/publish-unit-test-result-action`
- [ ] Update workflows to require VS 2026
- [ ] Add Windows SDK installation step to workflows
- [ ] Update ALL NuGet packages to latest versions
- [ ] Test all GitHub Actions workflows pass

**Risk Level:** 🟡 Medium (runner environment may differ)

---

### Phase 4: Deferred Discussions
**Timeline: After Phase 3 is complete**

- [ ] 🔲 **Prefast Configuration:** Discuss optimization and integration with new code analysis
- [ ] 🔲 **ARM64 Support:** Discuss need and implementation plan
- [ ] Additional enhancements as identified

**Risk Level:** 🔲 Variable (depends on decisions from discussions)

---

## ✅ Testing Checklist

### Build Testing
- [ ] Clean build in VS 2026 - All configurations
  - [ ] Debug x86
  - [ ] Debug x64
  - [ ] Release x86
  - [ ] Release x64
  - [ ] Debug_Unicode x86
  - [ ] Debug_Unicode x64
  - [ ] Release_Unicode x86
  - [ ] Release_Unicode x64
  - [ ] Prefast x86
  - [ ] Prefast x64
  - [ ] Fuzz x86
  - [ ] Fuzz x64

- [ ] Build from VS Code
  - [ ] All configurations in tasks.json
  - [ ] Verify MSBuild version detection

### CI/CD Testing
- [ ] GitHub Actions - all workflows pass
  - [ ] github-ci.yml
  - [ ] codeql.yml
  - [ ] clang.yml
  - [ ] dependency-review.yml
  - [ ] devskim.yml

### Functional Testing
- [ ] Unit tests pass in all configurations
- [ ] Fuzzing configuration works
- [ ] Prefast analysis runs without errors
- [ ] MFCMapi application runs correctly
- [ ] MrMAPI command-line tool works
- [ ] No new warnings introduced

### Submodule Testing
- [ ] mapistub builds correctly in VS 2026
- [ ] mapistub uses v145 toolset
- [ ] mapistub uses Windows 11 SDK (10.0.22621.0)
- [ ] mapistub compiles with C++20 (stdcpp20)
- [ ] mapistub all configurations build (Debug, Release, etc.)
- [ ] Integration with main project intact
- [ ] No new warnings in mapistub
- [ ] mapistub tests pass (if any)
- [ ] No breaking changes from submodule updates

---

## ?? Documentation Updates

### Files to Update:

**Main Project (mfcmapi):**
- [ ] **README.md**
  - Add VS 2026 as required version
  - Add C++20 requirement
  - Update build instructions
  - Note about VS 2022 NO LONGER SUPPORTED
  - Note about automatic SDK installation
  - Add section about mapistub submodule requirements

- [ ] **CONTRIBUTING.md**
  - Update .vsconfig instructions
  - Add VS 2026 setup steps
  - Document C++20 features available
  - Note about mapistub modernization

- [ ] **docs/Documentation.md**
  - Update development environment setup
  - Add note about automatic SDK installation via .vsconfig
  - Add troubleshooting section for SDK issues
  - Document C++20 migration considerations

- [ ] **scripts/Install-RequiredSDKs.ps1** (create new)
  - PowerShell script for automated SDK installation
  - Use in CI/CD or manual setups
  - Link from README

- [ ] **Create: VS2026-Migration.md** (this document)
  - Keep as reference
  - Link from README

- [ ] **.github/workflows/README.md** (if exists)
  - Update action versions
  - Document build matrix

**Submodule (mapistub):**
- [ ] **mapistub/README.md**
  - Add VS 2026 requirement
  - Add C++20 requirement
  - Update build instructions
  - Document automatic SDK installation

- [ ] **mapistub/.vsconfig** (create/update)
  - Match main project components
  - Windows 11 SDK 22621

- [ ] **mapistub/scripts/Install-RequiredSDKs.ps1** (create new)
  - Same as main project

- [ ] **mapistub/CONTRIBUTING.md** (if exists)
  - Update with VS 2026 requirements

---

## ?? Quick Start Guide

### To Begin Modernization:

1. **Create a branch:**
```bash
git checkout -b vs2026-modernization
```

2. **Backup current state** (optional):
```bash
git tag vs2022-baseline
```

3. **Phase 1 - Update .vsconfig:**
```bash
# Edit .vsconfig
# Replace Windows10SDK.18362 with Windows11SDK.22621
```

4. **Phase 2 - Add Directory.Build.props:**
```bash
# Create new file at solution root
# Add conditional toolset logic
```

5. **Test in VS 2026:**
```bash
# Open solution in VS 2026
# Build ? Batch Build ? Select All ? Rebuild
```

6. **Commit and push:**
```bash
git add .
git commit -m "Upgrade to Visual Studio 2026 (v145 toolset, VS 18)"
git push origin vs2026-modernization
```

---

## ? Questions for Discussion

### 1. Toolset Strategy
- **Q:** Single version (v145 only) or conditional (v143/v145)?
- **Recommendation:** Conditional for maximum compatibility
- **Decision:** Use v145 via Directory.Build.props (accepting VS 2026 as minimum requirement)

### 2. VS 2022 Support Duration
- **Q:** How long should we maintain VS 2022 compatibility?
- **Options:** 
  - Keep indefinitely
  - Until next major release
  - 6 months / 1 year
- **Decision:** Not maintaining - moving to VS 2026 only (v145 toolset)

### 3. C++ Language Standard
- **Q:** Ready to upgrade to C++20?
- **Considerations:**
  - VS 2026 supports C++20 fully
  - MAPIStubLibrary submodule will also be upgraded to C++20
  - May require code changes
- **Decision:** ? Upgrade to C++20 (stdcpp20) for both mfcmapi and mapistub

### 4. Windows SDK Version
- **Q:** Pin to 10.0.22621.0 or keep flexible (10.0)?
- **Trade-off:** Reproducibility vs. flexibility
- **Decision:** Pin to 10.0.22621.0 (Windows 11 SDK) for reproducible builds with automatic installation

### 5. GitHub Actions Strategy
- **Q:** Test both VS versions in CI or just latest?
- **Options:**
  - Latest only (faster)
  - Both versions (more thorough)
  - Latest + periodic compatibility check
- **Decision:** VS 2026 only (GitHub Actions runners will need VS 2026)

### 6. ARM64 Priority
- **Q:** When (if ever) should we add ARM64 support?
- **Drivers:**
  - User requests
  - Windows on ARM adoption
  - Microsoft requirements
- **Decision:** Defer to future - not needed for initial VS 2026 migration

### 7. Package Management
- **Q:** Migrate from packages.config to PackageReference?
- **Benefit:** Modern, CI-friendly
- **Risk:** Migration effort
- **Decision:** Already on u/sgriffin/nuget branch - address separately from VS 2026 migration

---

## ?? Useful Resources

### Visual Studio
- [Visual Studio 2026 Release Notes](https://docs.microsoft.com/visualstudio/releases/2026/release-notes)
- [C++ in Visual Studio](https://docs.microsoft.com/cpp/)
- [Platform Toolset](https://docs.microsoft.com/cpp/build/how-to-modify-the-target-framework-and-platform-toolset)

### C++ Standards
- [C++20 Features](https://en.cppreference.com/w/cpp/20)
- [C++23 Features](https://en.cppreference.com/w/cpp/23)
- [Microsoft C++ Language Conformance](https://docs.microsoft.com/cpp/overview/visual-cpp-language-conformance)

### GitHub Actions
- [GitHub Actions Documentation](https://docs.github.com/actions)
- [actions/checkout](https://github.com/actions/checkout)
- [github/codeql-action](https://github.com/github/codeql-action)

### Security
- [Spectre Mitigation](https://docs.microsoft.com/cpp/build/reference/qspectre)
- [AddressSanitizer](https://docs.microsoft.com/cpp/sanitizers/asan)
- [C++ Core Guidelines](https://isocpp.github.io/CppCoreGuidelines/CppCoreGuidelines)

---

## ?? Risk Assessment

| Task | Risk Level | Impact | Effort | Priority |
|------|-----------|--------|--------|----------|
| Update .vsconfig | ?? Low | Low | Low | High |
| Conditional toolset | ?? Medium | Medium | Low | High |
| GitHub Actions update | ?? Low | Medium | Low | High |
| NuGet package update | ?? Medium | Low | Low | Medium |
| Windows SDK pin | ?? Medium | Medium | Low | Medium |
| C++ standard upgrade | ?? Medium-High | High | Medium | Medium |
| ARM64 support | ?? Low | Low | High | Low |
| VS Code updates | ?? Low | Low | Low | Medium |

---

## ? Success Criteria

Project modernization is complete when:

**Main Project:**
- ? Solution builds in VS 2026 with no errors/warnings
- ? Solution file indicates VS 2026 (Version 18)
- ? All projects use v145 toolset
- ? All projects compile with C++20 (stdcpp20)
- ? All projects use Windows 11 SDK (10.0.22621.0)
- ? All unit tests pass
- ? GitHub Actions workflows pass (with VS 2026)
- ? VS Code builds work
- ? Documentation is updated with VS 2026 and C++20 requirements
- ? .vsconfig reflects current requirements (Windows 11 SDK)

**Submodule (mapistub):**
- ? mapistub builds in VS 2026 with no errors/warnings
- ? mapistub solution file indicates VS 2026 (Version 18)
- ? mapistub uses v145 toolset
- ? mapistub compiles with C++20 (stdcpp20)
- ? mapistub uses Windows 11 SDK (10.0.22621.0)
- ? mapistub .vsconfig reflects current requirements
- ? mapistub documentation updated
- ? Integration with main project works correctly

---

## ?? Timeline Estimate

- **Phase 1:** 2-4 hours (vsconfig updates, GitHub Actions updates, testing)
- **Phase 2:** 8-12 hours (Directory.Build.props, solution updates, mapistub updates, C++20 testing, extensive builds)
- **Phase 3:** Variable (depends on scope - ARM64, analysis enhancements, etc.)

**Total for Phases 1-2:** ~1.5-2 days

**Breakdown:**
- Main project updates: ~4-6 hours
- Mapistub updates: ~2-3 hours  
- C++20 testing and fixes: ~2-3 hours
- Documentation: ~1-2 hours
- CI/CD verification: ~1 hour

---

## ?? Stakeholder Sign-off

- [ ] **Development Lead:** _______________
- [ ] **CI/CD Owner:** _______________
- [ ] **Security Review:** _______________
- [ ] **QA Sign-off:** _______________

---

## Notes

### Files to be Created:
**Main Project:**
- `Directory.Build.props` - Centralized build configuration (v145, Windows 11 SDK, C++20)
- `scripts/Install-RequiredSDKs.ps1` - Automated SDK installation script

**Submodule (mapistub):**
- `mapistub/Directory.Build.props` - Same configuration as main project
- `mapistub/scripts/Install-RequiredSDKs.ps1` - SDK installation script
- `mapistub/.vsconfig` - VS component requirements (if doesn't exist)

### Files to be Modified:
**Main Project:**
- `MFCMapi.sln` - Update to VS 18
- `.vsconfig` - Update to Windows 11 SDK 22621
- All `.vcxproj` files - Remove PlatformToolset (inherited from Directory.Build.props)
- `.github/workflows/*.yml` - Add SDK installation steps
- `README.md` - Document VS 2026 and C++20 requirements
- `CONTRIBUTING.md` - Update developer setup instructions
- `docs/Documentation.md` - Add SDK auto-install and C++20 notes

**Submodule (mapistub):**
- `mapistub/*.sln` - Update to VS 18
- `mapistub/.vsconfig` - Update to Windows 11 SDK 22621, remove unnecessary components
- All `mapistub/*.vcxproj` files - Remove PlatformToolset
- `mapistub/README.md` - Document VS 2026 and C++20 requirements

### Breaking Changes:
?? **This is a breaking change!**
- VS 2022 will NO LONGER work
- Minimum requirement: Visual Studio 2026
- C++20 may require code changes
- Windows 11 SDK 10.0.22621.0 required

### Migration Path:
1. **Phase 1:** Update .vsconfig and GitHub Actions (non-breaking, preparatory)
2. **Phase 2:** Create Directory.Build.props, update solution files, test everything
3. **Phase 3:** Optional enhancements (security, analysis, ARM64)

_(Add any additional notes, findings, or decisions here)_

---

## Current Branch Status

**Branch:** u/sgriffin/nuget  
**Note:** This is already a feature branch - consider whether to integrate VS 2026 updates here or create a separate branch.

**Decision:** Create separate vs2026-upgrade branch, merge NuGet work separately

---

## ?? Quick Reference Checklist

### Critical Path Items:

**Preparation:**
- [ ] Create branch: `git checkout -b vs2026-upgrade`
- [ ] Tag baseline: `git tag vs2022-baseline`

**Phase 1 (Non-Breaking):**
- [ ] Update `.vsconfig` ? Windows11SDK.22621
- [ ] Update GitHub Actions versions
- [ ] Check NuGet packages for updates

**Phase 2 (Main Migration):**
- [ ] Create `Directory.Build.props` (v145 + C++20 + SDK)
- [ ] Create `mapistub/Directory.Build.props`
- [ ] Update `MFCMapi.sln` ? VS 18
- [ ] Update mapistub solution ? VS 18
- [ ] Create `scripts/Install-RequiredSDKs.ps1`
- [ ] Create `mapistub/scripts/Install-RequiredSDKs.ps1`
- [ ] Remove PlatformToolset from all .vcxproj files
- [ ] Test all configurations build
- [ ] Test mapistub standalone
- [ ] Run all unit tests
- [ ] Update all documentation

**Phase 3 (Optional):**
- [ ] Enhanced security settings
- [ ] C++ Core Guidelines checks
- [ ] ARM64 support planning

**Final Steps:**
- [ ] All tests passing
- [ ] Documentation complete
- [ ] Create PR
- [ ] Consider PR to upstream mapistub

---

**Last Updated:** 2026
**Status:** Planning Phase ? Complete - Ready for Implementation
