# Builds the Agensi upload package: dist/sheets-formula-coach.zip
# Contents: one top-level folder sheets-formula-coach/ with SKILL.md, references/, examples/ and CHANGELOG.md only
# (templates/ too, if it exists). README, LICENSE, evals/, tests/ and scripts/ are left out on purpose.
# Usage (from anywhere): pwsh ./scripts/build-agensi-zip.ps1

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root     = Split-Path -Parent $PSScriptRoot
$name     = 'sheets-formula-coach'
$distDir  = Join-Path $root 'dist'
$zipPath  = Join-Path $distDir "$name.zip"
$required = @('SKILL.md', 'references', 'examples', 'CHANGELOG.md')
$optional = @('templates')

New-Item -ItemType Directory -Force -Path $distDir | Out-Null
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

$items = @()
foreach ($i in $required) {
    if (-not (Test-Path (Join-Path $root $i))) { throw "Missing required item: $i" }
    $items += $i
}
foreach ($i in $optional) { if (Test-Path (Join-Path $root $i)) { $items += $i } }

$zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($item in $items) {
        $src   = Join-Path $root $item
        $files = if ((Get-Item $src).PSIsContainer) { Get-ChildItem -Path $src -Recurse -File } else { @(Get-Item $src) }
        foreach ($f in $files) {
            $rel   = $f.FullName.Substring($root.Length).TrimStart('\', '/').Replace('\', '/')
            $entry = "$name/$rel"
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $f.FullName, $entry, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
        }
    }
}
finally {
    $zip.Dispose()
}

$check = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
try {
    Write-Host "Built $zipPath ($((Get-Item $zipPath).Length) bytes)"
    $check.Entries | ForEach-Object { Write-Host ("  {0}  ({1} bytes)" -f $_.FullName, $_.Length) }
}
finally {
    $check.Dispose()
}
