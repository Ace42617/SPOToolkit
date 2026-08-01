# Copy SPO Toolkit adaptations onto vendor/levelup source before build.
$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$vendorSrc = Join-Path $repoRoot "vendor\levelup\src"
$overlaySrc = Join-Path $PSScriptRoot "overlay\src"

if (-not (Test-Path $vendorSrc)) { throw "Missing $vendorSrc" }
if (-not (Test-Path $overlaySrc)) { throw "Missing $overlaySrc" }

Get-ChildItem -Path $overlaySrc -Recurse -File | ForEach-Object {
  $rel = $_.FullName.Substring($overlaySrc.Length).TrimStart('\', '/')
  $dest = Join-Path $vendorSrc $rel
  New-Item -ItemType Directory -Force -Path (Split-Path $dest) | Out-Null
  Copy-Item $_.FullName $dest -Force
  Write-Host "Applied adaptation: $rel"
}
