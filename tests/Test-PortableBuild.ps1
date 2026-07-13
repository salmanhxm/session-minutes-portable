[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ReleaseRoot,

    [Parameter(Mandatory)]
    [string]$ProjectRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ReleaseRoot = (Resolve-Path -LiteralPath $ReleaseRoot).Path
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
$engine = Join-Path $ReleaseRoot 'engine\SessionMinutesEngine.exe'
$launcher = Join-Path $ReleaseRoot 'SessionMinutesPortable.exe'
$gui = Join-Path $ReleaseRoot 'app\session_minutes_gui.ps1'

foreach ($path in @($engine, $launcher, $gui)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Missing portable artifact: $path"
    }
}

$helpOutput = & $engine --help 2>&1
if ($LASTEXITCODE -ne 0 -or ($helpOutput -join "`n") -notmatch 'preview') {
    throw 'The compiled engine help smoke test failed.'
}

$fixtureRoot = Join-Path ([System.IO.Path]::GetTempPath()) (
    'session-minutes-portable-smoke-' + [Guid]::NewGuid().ToString('N')
)
New-Item -ItemType Directory -Path $fixtureRoot | Out-Null
$template = Get-ChildItem -LiteralPath $ProjectRoot -File -Filter '*.docx' |
    Select-Object -First 1
if ($null -eq $template) {
    throw 'No root-level DOCX template exists in the smoke-test project.'
}
Copy-Item -LiteralPath $template.FullName -Destination $fixtureRoot
$batch = Get-ChildItem -LiteralPath $ProjectRoot -Directory |
    Where-Object Name -Match '^[1-9][0-9]*$' |
    Select-Object -First 1
if ($null -eq $batch) {
    throw 'No numeric batch folder exists in the smoke-test project.'
}
Copy-Item -LiteralPath $batch.FullName -Destination $fixtureRoot -Recurse

& $engine preview --project-root $fixtureRoot
if ($LASTEXITCODE -ne 0) {
    throw 'The compiled engine preview smoke test failed.'
}
if (-not (Test-Path -LiteralPath (Join-Path $fixtureRoot 'outputs\preview\preview.json') -PathType Leaf)) {
    throw 'The compiled engine did not create preview.json.'
}

Write-Host 'Portable build smoke test passed.'
Write-Host "Temporary fixture retained at: $fixtureRoot"
