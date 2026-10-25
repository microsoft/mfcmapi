# MFCMAPI

MFCMAPI provides access to MAPI stores to facilitate investigation of Exchange and Outlook issues and to provide developers with a sample for MAPI development.

MrMAPI is the companion command-line tool for MAPI diagnostics.

> **Caution:** MFCMAPI can modify or permanently delete mailbox data. Back up important data and only make changes you understand.

## Get MFCMAPI

### Install with WinGet (recommended)

For 64-bit Outlook, install with [WinGet](https://learn.microsoft.com/en-us/windows/package-manager/winget/):

```powershell
winget install mfcmapi
```

For **32-bit Outlook**, specify x86 instead:

```powershell
winget install mfcmapi --architecture x86
```

After installation, launch `MFCMapi` from a new terminal.

To update an existing WinGet installation:

```powershell
winget update mfcmapi
```

### Which version do I need?

Match your **Outlook installation**, not Windows: use x64 for 64-bit Outlook or x86 for 32-bit Outlook. Check under **File > Office Account > About Outlook**.

### Download portable ZIPs

**[Download the latest release](https://github.com/microsoft/mfcmapi/releases/latest)** and extract the appropriate archive:

| Your Outlook | MFCMAPI archive |
|--------------|-----------------|
| **64-bit Outlook** | `MFCMAPI.x64.exe.<version>.zip` |
| **32-bit Outlook** | `MFCMAPI.exe.<version>.zip` |

Here, `<version>` is the release version. Run `MFCMapi.exe` from the extracted folder; no installer is required.

**Looking for MrMAPI?** Download it from the same release page: `MrMAPI.x64.exe.<version>.zip` for 64-bit Outlook or `MrMAPI.exe.<version>.zip` for 32-bit Outlook.

Matching PDB archives for debugging are also available on the release page. Match the release and architecture when downloading these files.

## Dependencies

- **MAPI:** Accessing profiles and mailboxes requires an installed MAPI implementation, normally supplied by classic Outlook. New Outlook for Windows does not supply classic Outlook's MAPI implementation. Match MFCMAPI/MrMAPI's architecture to the installed MAPI provider.
- **Visual C++ runtime:** MFCMAPI and MrMAPI statically link their own C++ runtime, but Outlook's MAPI DLLs can require the Microsoft Visual C++ Redistributable. If a missing runtime such as `VCRUNTIME140_1.dll` prevents MAPI from loading, install or repair the [latest supported Visual C++ Redistributable](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist) matching Outlook's architecture (x64 or x86), not just Windows. A missing dependency can cause a "module could not be found" error even when the MAPI DLL itself exists.

## Help and feedback

For assistance using MFCMAPI, developing add-ins, or general MAPI development, consult the [documentation](docs/Documentation.md). Find a bug? Need help? Have a suggestion? Report your issues through the [issues tab](https://github.com/microsoft/mfcmapi/issues).  

## Contributing

Interested in contributing? See [Contributing](CONTRIBUTING.md) for build requirements and development setup.

## Badges

[![continuous-integration](https://github.com/microsoft/mfcmapi/actions/workflows/github-ci.yml/badge.svg)](https://github.com/microsoft/mfcmapi/actions/workflows/github-ci.yml)  
[![Clang-format](https://github.com/microsoft/mfcmapi/actions/workflows/clang.yml/badge.svg)](https://github.com/microsoft/mfcmapi/actions/workflows/clang.yml)  
[![CodeQL](https://github.com/microsoft/mfcmapi/actions/workflows/codeql.yml/badge.svg)](https://github.com/microsoft/mfcmapi/actions/workflows/codeql.yml)  
[![Dependency Review](https://github.com/microsoft/mfcmapi/actions/workflows/dependency-review.yml/badge.svg)](https://github.com/microsoft/mfcmapi/actions/workflows/dependency-review.yml)  
[![DevSkim](https://github.com/microsoft/mfcmapi/actions/workflows/devskim.yml/badge.svg)](https://github.com/microsoft/mfcmapi/actions/workflows/devskim.yml)  
[![OpenSSF Scorecard](https://api.securityscorecards.dev/projects/github.com/microsoft/mfcmapi/badge)](https://scorecard.dev/viewer/?uri=github.com%2Fmicrosoft%2Fmfcmapi)  
[![OpenSSF Best Practices](https://www.bestpractices.dev/projects/7901/badge)](https://www.bestpractices.dev/projects/7901)  
[![Facebook](https://badge.facebook.com/badge/26764016480.2776.1538253884.png)](https://www.facebook.com/MFCMAPI)