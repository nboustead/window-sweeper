<#
Window Sweeper release builder.

Packages WindowSweeper.ps1 + Run-WindowSweeper.bat + sweeper.ico + README.txt
into release/WindowSweeper-v<version>.zip using System.IO.Compression's
ZipArchive with EXPLICIT, forward-slash entry names - never
Compress-Archive, which writes backslash-separated entry paths that a
Linux-hosted unzip (or Netlify's build) reads as one literal filename
instead of a path (house rule).

Usage (from anywhere):
  powershell -File scripts\build-release.ps1
  powershell -File scripts\build-release.ps1 -Version 0.4.1
#>

param(
    [string]$Version = '0.4.0'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$repoRoot = Split-Path -Parent $PSScriptRoot
$releaseDir = Join-Path $repoRoot 'release'
$zipName = "WindowSweeper-v$Version.zip"
$zipPath = Join-Path $releaseDir $zipName

if (-not (Test-Path $releaseDir)) { New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null }
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

# Source file (repo root) -> zip entry name. Flat by design (no
# subfolders in this release), but entry names are still built explicitly
# and forward-slash-normalized rather than trusting any path-copying API.
$files = @(
    @{ Source = 'WindowSweeper.ps1'; Entry = 'WindowSweeper.ps1' },
    @{ Source = 'Run-WindowSweeper.bat'; Entry = 'Run-WindowSweeper.bat' },
    @{ Source = 'sweeper.ico'; Entry = 'sweeper.ico' },
    @{ Source = 'README.txt'; Entry = 'README.txt' }
)

foreach ($f in $files) {
    $srcPath = Join-Path $repoRoot $f.Source
    if (-not (Test-Path $srcPath)) { throw "Missing source file: $srcPath" }
}

$zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($f in $files) {
        $srcPath = Join-Path $repoRoot $f.Source
        $entryName = $f.Entry -replace '\\', '/'
        $entry = $zip.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
        $entryStream = $entry.Open()
        try {
            $bytes = [System.IO.File]::ReadAllBytes($srcPath)
            $entryStream.Write($bytes, 0, $bytes.Length)
        } finally {
            $entryStream.Dispose()
        }
    }
} finally {
    $zip.Dispose()
}

Write-Host "Built: $zipPath"
Write-Host ""
Write-Host "Entries:"
$verifyZip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
try {
    foreach ($e in $verifyZip.Entries) {
        $flag = ''
        if ($e.FullName -match '\\') { $flag = '  <-- WARNING: backslash in entry name!' }
        Write-Host "  $($e.FullName)  ($($e.Length) bytes)$flag"
    }
} finally {
    $verifyZip.Dispose()
}

$hash = Get-FileHash -Path $zipPath -Algorithm SHA256
Write-Host ""
Write-Host "SHA256: $($hash.Hash)"
Write-Host "Size: $((Get-Item $zipPath).Length) bytes"
