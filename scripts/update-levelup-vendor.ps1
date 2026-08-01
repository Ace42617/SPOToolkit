# Replace vendor/levelup with a fresh shallow clone of upstream, then sync build → levelup/.
# Usage: pwsh -File scripts/update-levelup-vendor.ps1 [-Ref main]
param(
  [string]$Ref = "main",
  [string]$Repo = "https://github.com/rajyraman/Levelup-for-Dynamics-CRM.git"
)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$vendor = Join-Path $root "vendor\levelup"
$tmp = Join-Path $env:TEMP ("levelup-upstream-" + [guid]::NewGuid().ToString("n"))

Write-Host "Cloning $Repo ($Ref)…"
git clone --depth 1 --branch $Ref $Repo $tmp
if ($LASTEXITCODE -ne 0) { throw "git clone failed" }

$commit = git -C $tmp rev-parse HEAD
if (Test-Path $vendor) { Remove-Item $vendor -Recurse -Force }
New-Item -ItemType Directory -Path (Split-Path $vendor -Parent) -Force | Out-Null
robocopy $tmp $vendor /E /XD .git node_modules build release /NFL /NDL /NJH /NJS /nc /ns /np | Out-Null
$commit | Set-Content (Join-Path $vendor "UPSTREAM_COMMIT.txt")
$Repo | Set-Content (Join-Path $vendor "UPSTREAM_REPO.txt")
Remove-Item $tmp -Recurse -Force

Write-Host "Vendor updated to $commit"
& (Join-Path $PSScriptRoot "sync-levelup.ps1")
