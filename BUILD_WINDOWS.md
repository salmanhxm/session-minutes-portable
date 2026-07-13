# Building on Windows 11

## Prerequisites

- Windows 11 x64
- Windows PowerShell 5.1
- CPython 3.12 x64 with `pip`
- Internet access only for the first dependency restore

## Clean repeatable build

```powershell
cd portable_app
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -PythonExe C:\Python312\python.exe
```

The script reads `VERSION`, verifies Python 3.12 x64, restores fully pinned build dependencies into an ABI- and requirements-specific cache, runs source tests, builds the Python engine and x64 C# launcher, writes executable metadata, creates the ZIP, and writes a `.sha256` file.

Outputs are created under `dist/`. Existing releases and build caches are not overwritten.

## Verification commands

```powershell
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile(
  (Resolve-Path .\src\session_minutes_gui.ps1), [ref]$null, [ref]$errors)
$errors

C:\Python312\python.exe -m compileall -q .\src .\tests
C:\Python312\python.exe -m unittest discover -s .\tests -p "test_*.py" -v
```

After building, run `tests\Test-PortableBuild.ps1` with a synthetic project folder or a disposable copy of local test data. Never use real case documents in CI.

## Signing

The build does not sign executables. For public distribution, sign both executables with a trusted Authenticode certificate and verify the signatures before publishing the checksum.
