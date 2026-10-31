$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$dist = Join-Path $root 'dist'
$singleDir = Join-Path $dist 'single'
$scriptDir = Join-Path $dist 'script'
$portableDir = Join-Path $dist 'portable'

$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $csc)) {
    throw "csc.exe not found: $csc"
}

if (Test-Path -LiteralPath $scriptDir) { Remove-Item -LiteralPath $scriptDir -Recurse -Force }
if (Test-Path -LiteralPath $portableDir) { Remove-Item -LiteralPath $portableDir -Recurse -Force }
foreach ($d in @($dist, $singleDir, $scriptDir, $portableDir)) {
    [void][IO.Directory]::CreateDirectory($d)
}

function Invoke-Csc {
    param([string[]]$Arguments)
    & $csc @Arguments
    if ($LASTEXITCODE -ne 0) { throw "csc failed with exit code $LASTEXITCODE" }
}

# 1. Single-file exe
$singleExe = Join-Path $singleDir 'LocalFileCleanupAssistant.exe'
Invoke-Csc @(
    '/nologo', '/target:winexe', '/platform:anycpu', '/optimize+',
    '/reference:System.dll', '/reference:System.Windows.Forms.dll',
    "/out:$singleExe",
    "/resource:$root\app\Bootstrap.ps1,LocalFileCleanupAssistant.Bootstrap.ps1",
    "/resource:$root\app\App.ps1,LocalFileCleanupAssistant.App.ps1",
    "/resource:$root\app\Scanner.ps1,LocalFileCleanupAssistant.Scanner.ps1",
    "/resource:$root\app\MainWindow.xaml,LocalFileCleanupAssistant.MainWindow.xaml",
    (Join-Path $root 'SingleExe.cs')
)

# 2. Portable launcher
$launcherExe = Join-Path $portableDir 'FileCleanupAssistant.exe'
Invoke-Csc @(
    '/nologo', '/target:winexe', '/platform:anycpu', '/optimize+',
    '/reference:System.dll', '/reference:System.Windows.Forms.dll',
    "/out:$launcherExe",
    (Join-Path $root 'Launcher.cs')
)

# 3. Script package
Copy-Item -LiteralPath (Join-Path $root 'app') -Destination (Join-Path $scriptDir 'app') -Recurse -Force
Copy-Item -LiteralPath (Join-Path $root 'Launch.cmd') -Destination (Join-Path $scriptDir 'Launch.cmd') -Force
Copy-Item -LiteralPath (Join-Path $root 'README.md') -Destination (Join-Path $scriptDir 'README.md') -Force
Copy-Item -LiteralPath (Join-Path $root 'LICENSE') -Destination (Join-Path $scriptDir 'LICENSE') -Force
$scriptZip = Join-Path $dist 'LocalFileCleanupAssistant-script.zip'
if (Test-Path -LiteralPath $scriptZip) { Remove-Item -LiteralPath $scriptZip -Force }
Compress-Archive -Path (Join-Path $scriptDir '*') -DestinationPath $scriptZip -Force

# 4. Portable package
Copy-Item -LiteralPath (Join-Path $root 'app') -Destination (Join-Path $portableDir 'app') -Recurse -Force
Copy-Item -LiteralPath (Join-Path $root 'Launch.cmd') -Destination (Join-Path $portableDir 'Launch.cmd') -Force
Copy-Item -LiteralPath (Join-Path $root 'README.md') -Destination (Join-Path $portableDir 'README.md') -Force
Copy-Item -LiteralPath (Join-Path $root 'LICENSE') -Destination (Join-Path $portableDir 'LICENSE') -Force
$portableZip = Join-Path $dist 'LocalFileCleanupAssistant-portable.zip'
if (Test-Path -LiteralPath $portableZip) { Remove-Item -LiteralPath $portableZip -Force }
Compress-Archive -Path (Join-Path $portableDir '*') -DestinationPath $portableZip -Force

# 5. Checksums
$lines = @()
foreach ($f in @($singleExe, $scriptZip, $portableZip)) {
    if (Test-Path -LiteralPath $f) {
        $h = (Get-FileHash -Algorithm SHA256 -LiteralPath $f).Hash
        $lines += ("{0}  {1}" -f $h, (Split-Path $f -Leaf))
    }
}
$sumFile = Join-Path $dist 'SHA256.txt'
[IO.File]::WriteAllLines($sumFile, $lines, [Text.UTF8Encoding]::new($false))
$lines | ForEach-Object { Write-Host $_ }
Write-Host "Done: $dist"

