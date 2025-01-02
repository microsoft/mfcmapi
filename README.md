# mfcmapi

MFCMAPI provides access to MAPI stores to facilitate investigation of Exchange and Outlook issues and to provide developers with a sample for MAPI development.

[Latest release](https://github.com/microsoft/mfcmapi/releases/latest)  
[Release stats (raw JSON)](https://api.github.com/repos/microsoft/mfcmapi/releases/latest)  
[Pretty release stats](https://somsubhra.github.io/github-release-stats/?username=microsoft&repository=mfcmapi&page=1&per_page=5)

## Contributing

MFCMAPI depends on the [MAPI Stub Library](https://github.com/microsoft/MAPIStubLibrary). When cloning, make sure to clone submodules. See [Contributing](CONTRIBUTING.md) for more details.

## Fuzzing

MFCMAPI supports fuzzing with [libFuzzer](https://llvm.org/docs/LibFuzzer.html) and the [fsanitize](https://learn.microsoft.com/en-us/cpp/build/reference/fsanitize?view=msvc-170) switch in Visual Studio. See [fuzz.cpp](fuzz/fuzz.cpp) for details.  
To run fuzzing for this project, follow these steps:
1. **Build Fuzzing Corpus**: 
   - Open Powershell prompt
   - Run [fuzz\Build-FuzzingCorpus.ps1](fuzz\Build-FuzzingCorpus.ps1) to generate a fuzzing corpus in [fuzz/corpus](fuzz/corpus) from Smart View unit test data.

1. **Switch Solution Configuration**:
   - Open MFCMAPI.sln in Visual Studio.
   - In the toolbar, locate the **Solution Configurations** dropdown.
   - Select **Fuzz** from the list of configurations.

1. **Debug Command Line Parameters**:
   - When running the fuzzing tests, use the following command line parameters:  
`$(ProjectDir)fuzz\corpus -artifact_prefix=fuzz\artifacts\`

## Help/Feedback

For assistance using MFCMAPI, developing add-ins, or general MAPI development, consult the [documentation](docs/Documentation.md). Find a bug? Need help? Have a suggestion? Report your issues through the [issues tab](https://github.com/microsoft/mfcmapi/issues).  

## Badges

[![continuous-integration](https://github.com/microsoft/mfcmapi/actions/workflows/github-ci.yml/badge.svg)](https://github.com/microsoft/mfcmapi/actions/workflows/github-ci.yml)  
[![Clang-format](https://github.com/microsoft/mfcmapi/actions/workflows/clang.yml/badge.svg)](https://github.com/microsoft/mfcmapi/actions/workflows/clang.yml)  
[![CodeQL](https://github.com/microsoft/mfcmapi/actions/workflows/codeql.yml/badge.svg)](https://github.com/microsoft/mfcmapi/actions/workflows/codeql.yml)  
[![Dependency Review](https://github.com/microsoft/mfcmapi/actions/workflows/dependency-review.yml/badge.svg)](https://github.com/microsoft/mfcmapi/actions/workflows/dependency-review.yml)  
[![DevSkim](https://github.com/microsoft/mfcmapi/actions/workflows/devskim.yml/badge.svg)](https://github.com/microsoft/mfcmapi/actions/workflows/devskim.yml)  
[![OpenSSF
Scorecard](https://api.securityscorecards.dev/projects/github.com/microsoft/mfcmapi/badge)](https://scorecard.dev/viewer/?uri=github.com%2Fmicrosoft%2Fmfcmapi)  
[![OpenSSF Best Practices](https://www.bestpractices.dev/projects/7901/badge)](https://www.bestpractices.dev/projects/7901)  
[![Facebook](https://badge.facebook.com/badge/26764016480.2776.1538253884.png)](https://www.facebook.com/MFCMAPI)