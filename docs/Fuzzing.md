# Fuzzing MFCMAPI

MFCMAPI supports fuzzing with [libFuzzer](https://llvm.org/docs/LibFuzzer.html) and the [fsanitize](https://learn.microsoft.com/en-us/cpp/build/reference/fsanitize?view=msvc-170) switch in Visual Studio. See [fuzz.cpp](../fuzz/fuzz.cpp) for implementation details.

## Prerequisites

- Visual Studio 2026 with the Fuzz build configuration
- PowerShell for corpus generation

## Quick Start

```bash
# 1. Generate the corpus (converts hex test data to binary)
npm run fuzz:corpus

# 2. Build the fuzz configuration
npm run build:fuzz

# 3. Run from VS Code: select "Fuzz (x64)" and press F5
```

## Building the Fuzzing Corpus

The fuzzer needs binary input files. The unit test data is stored as hex strings in `.dat` files. Run the corpus builder to convert them:

```bash
npm run fuzz:corpus
```

This reads from `UnitTest/SmartViewTestData/In/*.dat` and writes binary files to `fuzz/corpus/`.

> **Note:** You must run this before fuzzing. The corpus directory won't exist until you do.

## Building the Fuzz Configuration

### From Command Line

```bash
npm run build:fuzz
```

### From Visual Studio

1. Open `MFCMapi.sln` in Visual Studio 2026
2. Select **Fuzz** from the Solution Configurations dropdown
3. Build the solution

## Running the Fuzzer

### From VS Code

1. Run `pwsh fuzz/Build-FuzzingCorpus.ps1` first (if you haven't already)
2. Open the Run and Debug panel (Ctrl+Shift+D)
3. Select "Fuzz (x64)" or "Fuzz (x86)" from the configuration dropdown
4. Press F5 to start

The fuzzer runs for 60 seconds by default.

### From Command Line

```bash
./bin/x64/Fuzz/MFCMapi.exe fuzz/corpus -artifact_prefix=fuzz/artifacts/ -max_total_time=60
```

### From Visual Studio

Set the debug command line arguments:

```
$(ProjectDir)fuzz\corpus -artifact_prefix=$(ProjectDir)fuzz\artifacts\
```

## Artifacts

When the fuzzer discovers a crash or hang, it saves the input that caused it to `fuzz/artifacts/`. These files can be used to reproduce and debug the issue.
