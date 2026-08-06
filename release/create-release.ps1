# Create a GitHub release zip for the current extension version (after: gh auth login).
# 1) Bump manifest.json version  2) Commit  3) Run this from repo root or release folder
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$manifestPath = Join-Path $root "manifest.json"
$manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
$version = [string]$manifest.version
if (-not $version) { throw "manifest.json has no version" }
$tag = "v$version"
$zipName = "SPOToolkit-$version.zip"
$staging = Join-Path $env:TEMP "SPOToolkit-$version-pack"
if (Test-Path $staging) { Remove-Item $staging -Recurse -Force }
New-Item -ItemType Directory -Path $staging | Out-Null

$files = @(
  "manifest.json", "sw.js", "background.js", "content.js", "powerAppsToolkitPopout.js",
  "popup.html", "popup.js", "views.html", "views.js",
  "view-formatter.html", "view-formatter.js", "view-formatter-iframe-bridge.js",
  "sharepoint-facts.js", "filterTypeaheadLogic.js", "icon.png",
  "exportCSV.js", "permissionsMatrixExport.js", "permissionsMatrixStyles.js", "matrixScanPlan.js", "reportListPlan.js",
  "getFields.js", "getSearchSchema.js", "getColumnCreatorContext.js", "createColumn.js",
  "checkListPage.js", "getListType.js", "getRefinableMappings.js", "getViewsData.js",
  "getPageContext.js", "getPageContextFull.js", "getPageContextJson.js", "getViewFormatContext.js", "getSiteLists.js",
  "compassColumnCreator.js", "compassToolkitPanels.js", "compassUniversalSearch.js", "matrixWorkerLock.js",
  "adminCenterWaitBanner.js", "recycleBinWaitBanner.js",
  "jszip.min.js", "jszip-preload.js", "jszip-restore-define.js",
  "README.md", "OWSSVR-LISTS.md"
)
foreach ($f in $files) {
  $p = Join-Path $root $f
  if (-not (Test-Path $p)) { throw "Missing: $p" }
  Copy-Item $p $staging
}
$rulesSrc = Join-Path $root "rules"
if (Test-Path $rulesSrc) { Copy-Item $rulesSrc (Join-Path $staging "rules") -Recurse }
Copy-Item (Join-Path $root "lib") (Join-Path $staging "lib") -Recurse
$levelupSrc = Join-Path $root "levelup"
if (-not (Test-Path $levelupSrc)) { throw "Missing: $levelupSrc — run scripts/sync-levelup.ps1 first" }
Copy-Item $levelupSrc (Join-Path $staging "levelup") -Recurse

$zipPath = Join-Path $PSScriptRoot $zipName
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Push-Location $staging
try {
  Compress-Archive -Path * -DestinationPath $zipPath -CompressionLevel Optimal
} finally {
  Pop-Location
}
Remove-Item $staging -Recurse -Force

$notesPath = Join-Path $PSScriptRoot "RELEASE_NOTES_$version.md"
Write-Host "Created $zipPath"
Write-Host "Version: $version  Tag: $tag"
if (-not (Test-Path $notesPath)) {
  Write-Warning "Missing notes file: $notesPath"
}
Write-Host "Next: git tag $tag; git push origin main; git push origin $tag"
Write-Host "gh release create `"$tag`" --title `"SPO Developer Toolkit $version`" --notes-file `"$notesPath`" `"$zipPath`""
