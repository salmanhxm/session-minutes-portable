[CmdletBinding()]
param(
    [string]$PythonExe = 'python',
    [string]$Version,
    [switch]$SkipSourceTests
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$sourceRoot = Join-Path $root 'src'
$stableUiSource = Join-Path $sourceRoot 'session_minutes_gui.ps1'
$launcherRoot = Join-Path $root 'launcher'
$assetsRoot = Join-Path $root 'assets'
$requirementsPath = Join-Path $root 'requirements-build.txt'
$versionPath = Join-Path $root 'VERSION'
$appIcon = Join-Path $assetsRoot 'session_minutes.ico'
$appLogo = Join-Path $assetsRoot 'session_minutes_icon_source.png'

function Resolve-ExecutablePath {
    param([Parameter(Mandatory)][string]$Value)
    if (Test-Path -LiteralPath $Value -PathType Leaf) {
        return (Resolve-Path -LiteralPath $Value).Path
    }
    return (Get-Command $Value -ErrorAction Stop).Source
}

function Invoke-Checked {
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][scriptblock]$Command
    )
    Write-Host "[$Label]"
    & $Command
    if ($LASTEXITCODE -ne 0) {
        throw "$Label failed with exit code $LASTEXITCODE."
    }
}

if ([string]::IsNullOrWhiteSpace($Version)) {
    if (-not (Test-Path -LiteralPath $versionPath -PathType Leaf)) {
        throw 'VERSION is missing.'
    }
    $Version = (Get-Content -LiteralPath $versionPath -Raw).Trim()
}
if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    throw "Version must use MAJOR.MINOR.PATCH: $Version"
}
$versionParts = @($Version.Split('.') | ForEach-Object { [int]$_ })
$fileVersion = "$($versionParts[0]).$($versionParts[1]).$($versionParts[2]).0"

$required = @(
    (Join-Path $sourceRoot 'session_minutes.py'),
    $stableUiSource,
    (Join-Path $launcherRoot 'Program.cs'),
    $appIcon,
    $appLogo,
    $requirementsPath,
    $versionPath
)
foreach ($path in $required) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Missing build input: $path"
    }
}

$PythonExe = Resolve-ExecutablePath -Value $PythonExe
$pythonFactsJson = & $PythonExe -c "import json,platform,struct,sys; print(json.dumps({'major':sys.version_info.major,'minor':sys.version_info.minor,'bits':struct.calcsize('P')*8,'machine':platform.machine(),'tag':sys.implementation.cache_tag,'version':platform.python_version()}))"
if ($LASTEXITCODE -ne 0) { throw 'Unable to inspect Python.' }
$pythonFacts = $pythonFactsJson | ConvertFrom-Json
if ($pythonFacts.major -ne 3 -or $pythonFacts.minor -ne 12 -or $pythonFacts.bits -ne 64) {
    throw "Release builds require CPython 3.12 x64. Detected $($pythonFacts.version), $($pythonFacts.bits)-bit."
}

$requirementsHash = (Get-FileHash -LiteralPath $requirementsPath -Algorithm SHA256).Hash.ToLowerInvariant()
$cacheKey = "$($pythonFacts.tag)-win_x64-$($requirementsHash.Substring(0, 12))"
$toolsRoot = Join-Path $root ".build\tools\$cacheKey"
$readyMarker = Join-Path $toolsRoot '.ready'
if (-not (Test-Path -LiteralPath $readyMarker -PathType Leaf)) {
    New-Item -ItemType Directory -Force -Path $toolsRoot | Out-Null
    $oldUtf8 = $env:PYTHONUTF8
    $oldIo = $env:PYTHONIOENCODING
    $oldProgress = $env:PIP_PROGRESS_BAR
    $env:PYTHONUTF8 = '1'
    $env:PYTHONIOENCODING = 'utf-8'
    $env:PIP_PROGRESS_BAR = 'off'
    try {
        Invoke-Checked -Label 'Install pinned build dependencies' -Command {
            & $PythonExe -m pip install --disable-pip-version-check --no-input --no-color --progress-bar off --only-binary=:all: --no-deps --target $toolsRoot --requirement $requirementsPath
        }
        [System.IO.File]::WriteAllText($readyMarker, $requirementsHash, [System.Text.UTF8Encoding]::new($false))
    }
    finally {
        $env:PYTHONUTF8 = $oldUtf8
        $env:PYTHONIOENCODING = $oldIo
        $env:PIP_PROGRESS_BAR = $oldProgress
    }
}

$oldPythonPath = $env:PYTHONPATH
$env:PYTHONPATH = if ([string]::IsNullOrWhiteSpace($oldPythonPath)) { $toolsRoot } else { "$toolsRoot;$oldPythonPath" }
try {
    $dependencyFactsJson = & $PythonExe -c "import json,lxml,PyInstaller; print(json.dumps({'lxml':lxml.__version__,'pyinstaller':PyInstaller.__version__}))"
    if ($LASTEXITCODE -ne 0) { throw 'Pinned build dependency validation failed.' }
    $dependencyFacts = $dependencyFactsJson | ConvertFrom-Json
    if ($dependencyFacts.lxml -ne '6.1.0' -or $dependencyFacts.pyinstaller -ne '6.21.0') {
        throw "Unexpected build dependency versions: lxml=$($dependencyFacts.lxml), PyInstaller=$($dependencyFacts.pyinstaller)"
    }

    if (-not $SkipSourceTests) {
        Invoke-Checked -Label 'Python source tests' -Command {
            & $PythonExe -m unittest discover -s (Join-Path $root 'tests') -p 'test_*.py' -v
        }
    }

    $buildId = Get-Date -Format 'yyyyMMdd-HHmmss'
    $runRoot = Join-Path $root ".build\runs\$buildId"
    $engineDistRoot = Join-Path $runRoot 'engine-dist'
    $engineWorkRoot = Join-Path $runRoot 'engine-work'
    $specRoot = Join-Path $runRoot 'spec'
    $generatedRoot = Join-Path $runRoot 'generated'
    $releaseName = "SessionMinutesPortable-$Version-$buildId"
    $releaseRoot = Join-Path $root "dist\$releaseName"
    $zipPath = Join-Path $root "dist\$releaseName-win-x64.zip"
    New-Item -ItemType Directory -Force -Path @($runRoot,$engineDistRoot,$engineWorkRoot,$specRoot,$generatedRoot,$releaseRoot,(Join-Path $releaseRoot 'app'),(Join-Path $releaseRoot 'assets'),(Join-Path $releaseRoot 'engine')) | Out-Null

    $engineVersionFile = Join-Path $generatedRoot 'engine-version.txt'
    $versionTuple = "($($versionParts[0]), $($versionParts[1]), $($versionParts[2]), 0)"
    $engineVersionText = @"
VSVersionInfo(
  ffi=FixedFileInfo(filevers=$versionTuple, prodvers=$versionTuple, mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('040904B0', [
      StringStruct('CompanyName', 'Salman'),
      StringStruct('FileDescription', 'Session Minutes OOXML Engine'),
      StringStruct('FileVersion', '$Version'),
      StringStruct('InternalName', 'SessionMinutesEngine'),
      StringStruct('LegalCopyright', 'Copyright (c) 2026 Salman'),
      StringStruct('OriginalFilename', 'SessionMinutesEngine.exe'),
      StringStruct('ProductName', 'Session Minutes Portable'),
      StringStruct('ProductVersion', '$Version')
    ])]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"@
    [System.IO.File]::WriteAllText($engineVersionFile, $engineVersionText, [System.Text.UTF8Encoding]::new($false))

    Invoke-Checked -Label 'Build Python engine' -Command {
        & $PythonExe -m PyInstaller --noconfirm --clean --onedir --console --name SessionMinutesEngine --icon $appIcon --version-file $engineVersionFile --distpath $engineDistRoot --workpath $engineWorkRoot --specpath $specRoot (Join-Path $sourceRoot 'session_minutes.py')
    }

    $engineBuild = Join-Path $engineDistRoot 'SessionMinutesEngine'
    $engineExe = Join-Path $engineBuild 'SessionMinutesEngine.exe'
    if (-not (Test-Path -LiteralPath $engineExe -PathType Leaf)) { throw 'The compiled engine executable was not created.' }
    Copy-Item -Path (Join-Path $engineBuild '*') -Destination (Join-Path $releaseRoot 'engine') -Recurse
    # Keep the complete, visually verified 0.6.8 WinForms implementation as
    # the UI-only baseline. The current engine and launcher are still built
    # from source, so document processing and portable behavior stay current.
    Copy-Item -LiteralPath $stableUiSource -Destination (Join-Path $releaseRoot 'app\session_minutes_gui.ps1')
    Copy-Item -LiteralPath $appIcon -Destination (Join-Path $releaseRoot 'assets\session_minutes.ico')
    Copy-Item -LiteralPath $appLogo -Destination (Join-Path $releaseRoot 'assets\session_minutes_icon_source.png')

    $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
    if (-not (Test-Path -LiteralPath $csc -PathType Leaf)) { throw 'The Windows x64 C# compiler was not found.' }
    $automationAssembly = [System.Management.Automation.PSObject].Assembly.Location
    if (-not (Test-Path -LiteralPath $automationAssembly -PathType Leaf) -or [System.Reflection.AssemblyName]::GetAssemblyName($automationAssembly).Version.Major -ne 3) {
        $automationAssembly = Get-ChildItem -LiteralPath (Join-Path $env:WINDIR 'Microsoft.Net\assembly\GAC_MSIL\System.Management.Automation') -Recurse -Filter 'System.Management.Automation.dll' -File | Select-Object -First 1 -ExpandProperty FullName
    }
    if ([string]::IsNullOrWhiteSpace([string]$automationAssembly)) { throw 'System.Management.Automation.dll was not found.' }

    $assemblyInfo = Join-Path $generatedRoot 'AssemblyInfo.cs'
    $assemblyText = @"
using System.Reflection;
[assembly: AssemblyTitle("Session Minutes Portable")]
[assembly: AssemblyDescription("Arabic RTL desktop application for safe session-minutes automation")]
[assembly: AssemblyCompany("Salman")]
[assembly: AssemblyProduct("Session Minutes Portable")]
[assembly: AssemblyCopyright("Copyright (c) 2026 Salman")]
[assembly: AssemblyVersion("$fileVersion")]
[assembly: AssemblyFileVersion("$fileVersion")]
[assembly: AssemblyInformationalVersion("$Version")]
"@
    [System.IO.File]::WriteAllText($assemblyInfo, $assemblyText, [System.Text.UTF8Encoding]::new($true))
    $launcherOutput = Join-Path $releaseRoot 'SessionMinutesPortable.exe'
    Invoke-Checked -Label 'Build x64 launcher' -Command {
        & $csc /nologo /utf8output /codepage:65001 /optimize+ /platform:x64 /target:winexe "/win32icon:$appIcon" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll "/reference:$automationAssembly" "/out:$launcherOutput" (Join-Path $launcherRoot 'Program.cs') $assemblyInfo
    }

    Copy-Item -LiteralPath (Join-Path $launcherRoot 'Run-Portable.cmd') -Destination $releaseRoot
    foreach ($doc in @('README_AR.md','README.md','LICENSE','THIRD_PARTY_NOTICES.md','SECURITY.md')) {
        $docPath = Join-Path $root $doc
        if (Test-Path -LiteralPath $docPath -PathType Leaf) { Copy-Item -LiteralPath $docPath -Destination $releaseRoot }
    }

    $gitRevision = 'unversioned'
    try {
        $candidate = (& git -C $root rev-parse --verify HEAD 2>$null).Trim()
        if ($LASTEXITCODE -eq 0 -and $candidate) { $gitRevision = $candidate }
    }
    catch { }
    $buildInfo = @(
        "Version: $Version",
        "Build ID: $buildId",
        'Architecture: win-x64',
        "Python: $($pythonFacts.version)",
        "PyInstaller: $($dependencyFacts.pyinstaller)",
        "lxml: $($dependencyFacts.lxml)",
        "Requirements SHA-256: $requirementsHash",
        "Source revision: $gitRevision"
    ) -join [Environment]::NewLine
    [System.IO.File]::WriteAllText((Join-Path $releaseRoot 'BUILD-INFO.txt'), $buildInfo, [System.Text.UTF8Encoding]::new($false))

    Compress-Archive -LiteralPath $releaseRoot -DestinationPath $zipPath -CompressionLevel Optimal
    $zipHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $checksumPath = "$zipPath.sha256"
    [System.IO.File]::WriteAllText($checksumPath, "$zipHash  $([System.IO.Path]::GetFileName($zipPath))`n", [System.Text.UTF8Encoding]::new($false))

    Write-Host ''
    Write-Host 'Portable build completed:'
    Write-Host "  Folder:   $releaseRoot"
    Write-Host "  ZIP:      $zipPath"
    Write-Host "  SHA-256:  $zipHash"
    Write-Host "  Checksum: $checksumPath"
}
finally {
    $env:PYTHONPATH = $oldPythonPath
}
