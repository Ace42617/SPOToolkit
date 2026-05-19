# Create a GitHub release for the current extension version (after: gh auth login).
# 1) Bump manifest.json version  2) Commit  3) Run this from repo root or release folder
$ErrorActionPreference = "Stop"
$version = "1.0.6.1"
$tag = "v$version"
$zipName = "SPOToolkit-$version.zip"
$root = Split-Path $PSScriptRoot -Parent
$staging = Join-Path $env:TEMP "SPOToolkit-$version-pack"
if (Test-Path $staging) { Remove-Item $staging -Recurse -Force }
New-Item -ItemType Directory -Path $staging | Out-Null

$files = @(
  "manifest.json", "background.js", "content.js",
  "recycleBinWaitBanner.js", "adminCenterWaitBanner.js",
  "popup.html", "popup.js", "views.html", "views.js",
  "view-formatter.html", "view-formatter.js",
  "sharepoint-facts.js", "filterTypeaheadLogic.js", "icon.png",
  "exportCSV.js", "getFields.js", "getSearchSchema.js", "checkListPage.js", "getListType.js",
  "getRefinableMappings.js", "getViewsData.js", "getPageContext.js", "getPageContextJson.js", "getViewFormatContext.js", "getSiteLists.js",
  "jszip.min.js", "jszip-preload.js", "jszip-restore.js",
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

$zipPath = Join-Path $PSScriptRoot $zipName
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Push-Location $staging
try {
  Compress-Archive -Path * -DestinationPath $zipPath -CompressionLevel Optimal
} finally {
  Pop-Location
}
Remove-Item $staging -Recurse -Force

Write-Host "Created $zipPath"
Write-Host "Next: git tag $tag; git push origin main; git push origin $tag"
Write-Host "gh release create `"$tag`" --title `"SPO Developer Toolkit $version`" --notes-file `"$PSScriptRoot\RELEASE_NOTES_$version.md`" `"$zipPath`""
