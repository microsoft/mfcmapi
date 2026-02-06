# CONTRIBUTING

All pull requests are welcome, there are just a few guidelines you need to follow.

When contributing to this repository, please first discuss the change by creating a new [issue](https://github.com/microsoft/mfcmapi/issues) or by replying to an existing one.

## REQUIREMENTS

* **Visual Studio 2026** with the following workloads:
  - Desktop development with C++
  - v145 toolset (VS 2026)
  - Windows 11 SDK (10.0.22621.0)
  - Spectre-mitigated libraries (recommended)
* **Node.js** (for npm-based builds - optional but recommended)

### Verify Prerequisites

Run the prerequisite check to verify your environment:
```bash
npm run prereq
```

Or use the .vsconfig file to install all required components:
1. Open Visual Studio Installer
2. Click **More** → **Import configuration**
3. Select the `.vsconfig` file from the repo root

## GETTING STARTED

* Make sure you have a [GitHub account](https://github.com/signup/free).
* Fork the repository, you can [learn about forking on Github](https://help.github.com/articles/fork-a-repo)
* [Clone the repo to your local machine](https://help.github.com/articles/cloning-a-repository/) with submodules:  
```bash
git clone --recursive https://github.com/microsoft/mfcmapi.git
```
* Install npm dependencies (for command-line builds):
```bash
npm install
```

## BUILDING

### From Visual Studio
Open `MFCMapi.sln` in Visual Studio 2026 and build.

### From VS Code
Press **F5** to build and debug. The default configuration is Debug_Unicode x64.

### From Command Line
```bash
# Build (default: Debug_Unicode x64)
npm run build

# Build specific configurations
npm run build:release          # Release_Unicode x64
npm run build:debug:x86        # Debug_Unicode x86
npm run build:debug:ansi:x64   # Debug (ANSI) x64

# Run tests
npm run test

# Clean all build outputs
npm run clean
```

The npm scripts automatically find MSBuild via vswhere, so no Developer Command Prompt is needed.

## MAKING CHANGES

* Create branch topic for the work you will do, this is where you want to base your work.
* This is usually the main branch.
* To quickly create a topic branch based on main, run  
```git checkout -b u/username/topic main```  
  * *Make sure to substitute your own name and topic in this command* *
* Once you have a branch, make your changes and commit them to the local branch.
* Run `npm run test` to verify your changes don't break existing tests.
* All submissions require a review and pull requests are how those happen. Consult
[GitHub Help](https://help.github.com/articles/about-pull-requests/) for more
information on pull requests.

## SUBMITTING CHANGES

* Push your changes to a topic branch in your fork of the repository.

## PUSH TO YOUR FORK AND SUBMIT A PULL REQUEST

At this point you're waiting on the code/changes to be reviewed.

## FUZZING

MFCMAPI supports fuzzing with libFuzzer. See [Fuzzing.md](docs/Fuzzing.md) for details.
