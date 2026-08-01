# Build Level Up (vendor/levelup) and sync patched artifacts into levelup/ for the extension.
# Update flow:
#   1) Pull/replace vendor/levelup from https://github.com/rajyraman/Levelup-for-Dynamics-CRM
#   2) Run:  pwsh -File scripts/sync-levelup.ps1
# Do not hand-edit files under levelup/ — they are regenerated here.
# Optional: -SkipInstall  (reuse existing node_modules)
# Optional: -SkipBuild    (re-patch last vendor/levelup/build only)

param(
  [switch]$SkipInstall,
  [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$vendor = Join-Path $root "vendor\levelup"
$build = Join-Path $vendor "build"
$dest = Join-Path $root "levelup"

if (-not (Test-Path (Join-Path $vendor "package.json"))) {
  throw "Missing vendor/levelup. Clone https://github.com/rajyraman/Levelup-for-Dynamics-CRM into vendor/levelup first."
}

# Apply SPO Toolkit overlays (context-aware actions, etc.) before build.
$applyAdaptations = Join-Path $PSScriptRoot "levelup-adaptations\apply.ps1"
if (Test-Path $applyAdaptations) {
  & $applyAdaptations
}

Push-Location $vendor
try {
  if (-not $SkipInstall) {
    if (Test-Path "package-lock.json") { npm ci } else { npm install }
  }
  if (-not $SkipBuild) {
    npm run build
  }
} finally {
  Pop-Location
}

if (-not (Test-Path (Join-Path $build "background.js"))) {
  throw "Level Up build missing at $build"
}

if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
New-Item -ItemType Directory -Path $dest | Out-Null

# Copy runtime assets (skip source maps to keep the extension pack lean)
$copyNames = @(
  "background.js", "content.js", "levelup-extension.js",
  "popup.js", "popup.html", "sidebar.js", "sidebar.html", "sidebar.css",
  "actions.js", "CustomCommandsService.js", "manifest.json"
)
foreach ($name in $copyNames) {
  $from = Join-Path $build $name
  if (-not (Test-Path $from)) { throw "Missing build artifact: $name" }
  Copy-Item $from (Join-Path $dest $name)
}
Copy-Item (Join-Path $build "icons") (Join-Path $dest "icons") -Recurse
Copy-Item (Join-Path $vendor "LICENSE") (Join-Path $dest "LICENSE")
Copy-Item (Join-Path $vendor "PRIVACY.md") (Join-Path $dest "PRIVACY.md") -ErrorAction SilentlyContinue
if (Test-Path (Join-Path $vendor "UPSTREAM_COMMIT.txt")) {
  Copy-Item (Join-Path $vendor "UPSTREAM_COMMIT.txt") (Join-Path $dest "UPSTREAM_COMMIT.txt")
}
if (Test-Path (Join-Path $vendor "UPSTREAM_REPO.txt")) {
  Copy-Item (Join-Path $vendor "UPSTREAM_REPO.txt") (Join-Path $dest "UPSTREAM_REPO.txt")
}

function Patch-File([string]$Path, [hashtable]$Replacements) {
  $text = Get-Content -Raw -LiteralPath $Path
  foreach ($key in $Replacements.Keys) {
    $text = $text.Replace($key, $Replacements[$key])
  }
  Set-Content -LiteralPath $Path -Value $text -NoNewline -Encoding utf8
}

# Remap Level Up root-relative asset paths into the levelup/ namespace.
Patch-File (Join-Path $dest "content.js") @{
  'getURL("levelup-extension.js")' = 'getURL("levelup/levelup-extension.js")'
  "getURL('levelup-extension.js')" = "getURL('levelup/levelup-extension.js')"
}

Patch-File (Join-Path $dest "background.js") @{
  'files:["content.js"]' = 'files:["levelup/content.js"]'
  'path:"sidebar.html"' = 'path:"levelup/sidebar.html"'
  # Prefer popup-driven open; keep side panel enabled without auto-stealing focus.
  'openIfDynamics:!0' = 'openIfDynamics:!1'
  'title:"Open Level Up Sidebar"' = 'title:"Open Power Apps toolkit side panel"'
  'title:"Level Up"' = 'title:"SPO Dev Toolkit"'
  # Prefer model-driven / Power Apps hosts for the context-menu entry (popup still gates on Xrm).
  'contexts:["page"]' = 'contexts:["page"],documentUrlPatterns:["*://*.dynamics.com/*","*://*.powerapps.com/*","*://*.powerplatform.com/*"]'
  # Open the in-page popout (SharePoint-style) instead of the Chromium side panel.
  'if("levelup-open"===t.menuItemId&&e?.id)try{await chrome.sidePanel.open({tabId:e.id})}catch(a){}' =
    'if("levelup-open"===t.menuItemId&&e?.id)try{await chrome.tabs.sendMessage(e.id,{type:"SPOToolkitOpenPowerAppsPopout"})}catch(a){try{await chrome.scripting.executeScript({target:{tabId:e.id},files:["powerAppsToolkitPopout.js"]}),await chrome.tabs.sendMessage(e.id,{type:"SPOToolkitOpenPowerAppsPopout"})}catch(b){}}'
}

# Express Mode was a fixed 320px card — stretch it to the SPO popup width and purple-brand the ribbon.
Patch-File (Join-Path $dest "popup.js") @{
  'width:"320px"' = 'width:"100%"'
  'maxWidth:"320px"' = 'maxWidth:"100%"'
  'linear-gradient(135deg, #1976d2, #42a5f5)' = 'linear-gradient(135deg, #742774, #5a1d5a)'
  'EXPRESS MODE' = 'POWER APPS'
  'Sidebar Modes' = 'Open expanded'
}

# SPO visual shell: DM Sans + denser chrome; keep Power Platform purple.
$themeCss = @'
/* SPO Toolkit overlay for Power Apps tools (generated — do not edit by hand) */
@import url("https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap");

:root {
  --spo-levelup-purple: #742774;
  --spo-levelup-purple-dark: #5a1d5a;
  --spo-levelup-purple-mid: #8e4a8e;
  --spo-levelup-purple-soft: #f3e8f8;
  --spo-levelup-bg: #f3f4fb;
  --spo-levelup-surface: #ffffff;
  --spo-levelup-border: #e8dff0;
  --spo-levelup-border-mid: #d4c2de;
  --spo-levelup-txt: #1c1f4a;
  --spo-levelup-txt-mid: #5a5f8a;
  --spo-r-md: 9px;
  --spo-r-lg: 13px;
}

html.spo-theme-dark {
  --spo-levelup-purple: #b794f6;
  --spo-levelup-purple-dark: #8a4fbc;
  --spo-levelup-purple-mid: #dbb2ff;
  --spo-levelup-purple-soft: #3b2a4a;
  --spo-levelup-bg: #0f1117;
  --spo-levelup-surface: #181b23;
  --spo-levelup-border: #4a3a58;
  --spo-levelup-border-mid: #6b5578;
  --spo-levelup-txt: #e8eaef;
  --spo-levelup-txt-mid: #9ea3c8;
}

html, body {
  font-family: "DM Sans", system-ui, sans-serif !important;
  -webkit-font-smoothing: antialiased;
  margin: 0 !important;
  height: 100% !important;
  max-height: 100% !important;
  overflow: hidden;
}

body {
  background: var(--spo-levelup-bg) !important;
  color: var(--spo-levelup-txt);
  width: 100% !important;
  max-width: 100% !important;
  min-height: 100% !important;
  box-sizing: border-box;
}

#popup-root, #root {
  min-height: 100% !important;
  height: 100% !important;
  max-height: 100% !important;
  width: 100%;
  box-sizing: border-box;
}

#root > div {
  height: 100% !important;
  max-height: 100% !important;
  min-height: 0 !important;
}

#popup-root > * {
  width: 100% !important;
  max-width: 100% !important;
  border-radius: 0 !important;
  box-shadow: none !important;
  min-height: 100%;
  box-sizing: border-box;
}

.MuiPaper-root,
.MuiCard-root {
  border-radius: var(--spo-r-md) !important;
  box-shadow: none !important;
}

.MuiButton-containedPrimary,
.MuiButton-root.MuiButton-contained {
  border-radius: var(--spo-r-md) !important;
  text-transform: none !important;
  font-weight: 600 !important;
  box-shadow: none !important;
}

.MuiButton-outlined {
  border-radius: var(--spo-r-md) !important;
  border-width: 1.5px !important;
  text-transform: none !important;
  font-weight: 600 !important;
}

.MuiChip-root {
  border-radius: 999px !important;
  font-weight: 600 !important;
}

.MuiTab-root {
  text-transform: none !important;
  font-weight: 600 !important;
  min-height: 40px !important;
}

.MuiTypography-root,
.MuiButton-root,
.MuiLink-root,
.MuiInputBase-root,
.MuiMenuItem-root {
  font-family: "DM Sans", system-ui, sans-serif !important;
}

/* Scrollbars — match SPO popup */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb {
  background: var(--spo-levelup-border-mid);
  border-radius: 10px;
}

/* Top chrome — matches SPO popup / View Manager header actions */
.sidebar-chrome-bar {
  border-bottom: 1px solid var(--spo-levelup-border) !important;
  background: var(--spo-levelup-surface) !important;
  box-shadow: 0 4px 12px rgba(28, 31, 74, 0.05);
}
html.spo-theme-dark .sidebar-chrome-bar {
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.28);
}

/* Early Riser / Night Owl pill — mirrors popup.html .dark-toggle-btn */
.spo-dark-toggle-btn {
  width: 52px !important;
  height: 24px !important;
  min-width: 52px !important;
  border: none !important;
  border-radius: 12px !important;
  cursor: pointer;
  position: relative;
  flex-shrink: 0;
  overflow: hidden;
  padding: 0 !important;
  margin: 0 !important;
  appearance: none;
  -webkit-appearance: none;
  background: linear-gradient(90deg, #f5e6c8 0%, #e8d4a8 35%, #3d2a5c 65%, #1a0a2e 100%) !important;
  box-shadow: inset 0 1px 2px rgba(0,0,0,.2), 0 0 0 1px rgba(255,255,255,.15) !important;
  transition: box-shadow .2s ease;
  color: transparent !important;
  font-size: 0 !important;
  line-height: 0 !important;
  text-decoration: none !important;
}
.spo-dark-toggle-btn:hover {
  background: linear-gradient(90deg, #f5e6c8 0%, #e8d4a8 35%, #3d2a5c 65%, #1a0a2e 100%) !important;
  border: none !important;
  box-shadow: inset 0 1px 2px rgba(0,0,0,.25), 0 0 0 1px rgba(255,255,255,.25) !important;
  text-decoration: none !important;
}
.spo-dark-toggle-btn:focus-visible {
  outline: 2px solid #fff;
  outline-offset: 2px;
}
.spo-dark-toggle-btn::before {
  content: '';
  position: absolute;
  left: 5px;
  top: 50%;
  transform: translateY(-50%);
  width: 14px;
  height: 14px;
  background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23b8860b' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='12' cy='10' r='4'/%3E%3Cpath d='M12 2v2M12 18v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M18 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41'/%3E%3Cpath d='M5 14h14'/%3E%3C/svg%3E") no-repeat center / contain;
  opacity: .95;
  pointer-events: none;
}
.spo-dark-toggle-btn::after {
  content: '';
  position: absolute;
  right: 5px;
  top: 50%;
  transform: translateY(-50%);
  width: 14px;
  height: 14px;
  background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23c9b8e8' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z'/%3E%3C/svg%3E") no-repeat center / contain;
  opacity: .9;
  pointer-events: none;
}
.spo-dark-toggle-knob {
  position: absolute;
  top: 2px;
  left: 2px;
  width: 20px;
  height: 20px;
  border-radius: 50%;
  background: radial-gradient(circle at 30% 30%, #fff 0%, #f0ebe0 50%, #e0d8c8 100%);
  box-shadow: 0 1px 3px rgba(0,0,0,.35), 0 0 0 1px rgba(255,255,255,.5);
  transition: transform .25s cubic-bezier(0.4, 0, 0.2, 1);
  pointer-events: none;
  display: flex;
  align-items: center;
  justify-content: center;
}
.spo-dark-toggle-knob::before {
  content: '';
  width: 11px;
  height: 11px;
  background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23c4952a' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='12' cy='9' r='3'/%3E%3Cpath d='M12 1v1.5M12 19.5V21M4.22 4.22l1.06 1.06M18.72 18.72l1.06 1.06M1 12h1.5M19.5 12H21M4.22 19.78l1.06-1.06M18.72 5.28l1.06-1.06'/%3E%3Cpath d='M4 14h16'/%3E%3C/svg%3E") no-repeat center / contain;
}
.spo-dark-toggle-btn--night .spo-dark-toggle-knob {
  transform: translateX(28px);
  background: radial-gradient(circle at 30% 30%, #c9b8e8 0%, #5c4d7a 50%, #2d2345 100%);
  box-shadow: 0 1px 3px rgba(0,0,0,.5), 0 0 0 1px rgba(255,255,255,.15);
}
.spo-dark-toggle-btn--night .spo-dark-toggle-knob::before {
  width: 10px;
  height: 10px;
  background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23e8e0f0' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z'/%3E%3C/svg%3E") no-repeat center / contain;
}

/* Section headers denser + SPO-like */
.MuiCard-root .MuiCardContent-root {
  font-family: "DM Sans", system-ui, sans-serif;
}

#root > * {
  min-height: 100%;
}
'@
Set-Content -LiteralPath (Join-Path $dest "spo-theme.css") -Value $themeCss.Trim() -Encoding utf8

function Inject-ThemeLink([string]$HtmlPath) {
  $html = Get-Content -Raw -LiteralPath $HtmlPath
  if ($html -notmatch 'spo-theme\.css') {
    $html = $html.Replace('</head>', "    <link rel=`"stylesheet`" href=`"spo-theme.css`" />`r`n  </head>")
  }
  $html = $html.Replace('<title>Level Up Commands</title>', '<title>Power Apps Toolkit</title>')
  $html = $html.Replace('<title>Level Up for Dynamics 365/Power Apps</title>', '<title>Power Apps Toolkit</title>')
  # Fit Express Mode inside the wider SPO popup shell
  $html = $html.Replace("width: 350px;", "width: 100%;")
  Set-Content -LiteralPath $HtmlPath -Value $html -NoNewline -Encoding utf8
}

Inject-ThemeLink (Join-Path $dest "popup.html")
Inject-ThemeLink (Join-Path $dest "sidebar.html")

# Narrow Level Up's own manifest is unused at runtime; keep for reference / license audit.
$refManifest = Get-Content -Raw (Join-Path $dest "manifest.json") | ConvertFrom-Json
$refNote = [ordered]@{
  note = "Reference only. SPOToolkit manifest.json is authoritative."
  upstreamName = $refManifest.name
  upstreamVersion = $refManifest.version
  syncedAt = (Get-Date).ToString("o")
}
$refNote | ConvertTo-Json | Set-Content (Join-Path $dest "SYNC_INFO.json") -Encoding utf8

$readme = @"
# Level Up (vendored build)

Built artifacts from [Level Up for Dynamics 365/Power Apps](https://github.com/rajyraman/Levelup-for-Dynamics-CRM) (MIT).

**Do not edit these files by hand.** Regenerate with:

``````powershell
pwsh -File scripts/sync-levelup.ps1
``````

Upstream source lives in vendor/levelup/.
"@
# Normalize fence markers (here-string uses doubled backticks for a single backtick fence)
$readme = $readme.Replace('``````', '```')
Set-Content -LiteralPath (Join-Path $dest "README.md") -Value $readme.Trim() -Encoding utf8

Write-Host "Synced Level Up -> $dest"
Write-Host "Upstream version: $($refManifest.version)"
Write-Host "Reload the unpacked extension after sync."
