# Release notes – SPO Developer Toolkit **1.0.7.0**

Major update on **1.0.6.1**: in-page compass toolkit, site permissions matrix export, worker-tab orchestration, and performance/refactor cleanup.

## Permissions matrix export

- **Site-wide permissions matrix** with subsite picker, group expansion, inherited permissions, folder sharing links, and multi-sheet XLSX output.
- **Worker-tab flow**: export runs on the original tab (locked); a **new tab** shows live progress in the compass Reports panel and auto-closes the worker when finished.
- **Progress console** with bar, elapsed time, scrollable log, cancel, and stale/reconnect handling via background session state.
- **Failed to fetch** fixes: matrix fetches run in the SharePoint page context on the worker tab (no offscreen document).

## Compass toolkit (in-page launcher)

- **Reports**, **Columns**, **Universal Settings Search**, and related panels from the on-page compass menu.
- **Universal search** hotkey: `Ctrl+Shift+,` (Mac: `Command+Shift+,`).
- **Column creator** panel with shared schema helpers (`lib/column*.mjs`).
- Status bar progress colors aligned with compass theme (light/dark).

## Worker lock overlay

- Full-screen lock on the worker tab during matrix export with live progress.
- **CRT Snake** mini-game: auto-play until you press an arrow key; resumes auto after idle; separate Auto High / Your High scores.

## View formatter & other

- View formatter iframe bridge and declarativeNetRequest rules (reload extension after update).
- Second-stage recycle bin wait banner and admin center wait banner.
- Popup export progress console; matrix prefs persisted in `chrome.storage.local`.

## Refactor & cleanup

- Removed unused **offscreen document** stack and dead fetch proxy code.
- Single compass export progress UI path (session storage + `compassToolkitPanels.js`).
- Dropped unused `offscreen` and `downloads` manifest permissions.
- **Experimental** fork: fixed `../lib/` imports in popup.

## Install

1. Download **`SPOToolkit-1.0.7.0.zip`** from this release.
2. Unzip and load unpacked in `chrome://extensions` or `edge://extensions` (Developer mode).
3. Reload the extension if updating from an earlier build.
