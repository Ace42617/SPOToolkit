# SPO Dev Toolkit

A Chrome / Edge extension for people who live in **SharePoint Online** (and often wander into **Power Apps / Dynamics**). It puts the admin pages, exports, and view tools you usually hunt for behind a floating compass on the page — plus a purple Level Up toolkit when you’re on a model-driven app.

**Current release:** [v1.1.0.0](https://github.com/Ace42617/SPOToolkit/releases/tag/v1.1.0.0) — see [release notes](release/RELEASE_NOTES_1.1.0.0.md).

---

## Install

1. Open `edge://extensions` or `chrome://extensions`
2. Turn on **Developer mode**
3. **Load unpacked** and pick the folder that contains `manifest.json`

For a packaged build, grab the zip from [Releases](https://github.com/Ace42617/SPOToolkit/releases), unzip it, and load that folder the same way. Reload the extension after updates so declarative network rules (View formatter preview) pick up changes.

---

## SharePoint: the compass

On `*.sharepoint.com`, a floating launcher opens an in-page panel. Tabs:

| Tab | What it’s for |
|-----|----------------|
| **Site Contents** | Lists, libraries, and subsites for the current web — open the default view, jump into View Manager, list settings, or permissions |
| **Quick Links** | Deep links for the current site, current user, page modes, recycle bins, and tenant admin surfaces, with a sticky filter |
| **Page Props** | Flattened `_spPageContextInfo` / page context — filter by name or value, copy values, refresh |
| **Columns** | On a list/library page: display name, internal name, crawled property, type; open column settings; create columns |
| **Reports** | Export list/library data, folder counts, path lengths, permissions matrix |
| **Refinables** | Jump to refinable managed property pages in search schema |
| **Views** | Open **View Manager** for the current list (edit/create views, columns, sort/filter/group, export column names) |

Also built into the compass experience:

- **Universal Search** — find tabs, quick links, and page info quickly (`Ctrl+Shift+,` / `Cmd+Shift+,`, or the search control in the panel)
- **Theme** — Early Riser / Night Owl
- **Settings** — launcher visibility & size, remember last tab, animation speed, optional keyboard shortcuts per tab
- **View formatter** — split preview + JSON editor for a list view’s `CustomFormatter` (from the popup / `{ }` flow); save patches the view over REST when the original list tab is still available

Row actions on Site Contents: View Manager, list settings, and list permissions.

---

## Reports

From the **Reports** tab (or the popup on a list/library page):

- **Download list/library** — CSV or Excel (XLSX); optional column picker and version history; pick one or more lists (multiple selections download as a zip)
- **Folder & item counts** — nested folder counts (document libraries)
- **Path length report** — long path audit (document libraries)
- **Permissions matrix** — same list/library picker as other reports; scans selected lists (and their sites) with inheritance and “given through” detail (Excel)

Exports use JSZip for XLSX/zip bundles and show progress in the panel.

Companion PowerShell (optional, under `scripts/`) can export or remediate permissions matrix work outside the browser — see those scripts’ own comments.

---

## Popup

Click the extension icon for a compact UI that follows where you are:

- **On SharePoint** — shortcuts into the same toolkit surfaces (exports, View Manager, View formatter, etc.)
- **On Power Apps / Dynamics** — purple **Power Apps** mode with **Open popout** for the in-page Level Up panel
- **Elsewhere** — a short notice plus **recently visited SharePoint sites** (from browser history and prior use) so you can jump back in

Feedback from the popup header goes to a Microsoft Form.

---

## Power Apps / Dynamics (Level Up)

On Power Platform hosts (`powerapps.com`, `dynamics.com`, and related), the extension loads a purple in-page popout adapted from [Level Up for Dynamics 365/Power Apps](https://github.com/rajyraman/Levelup-for-Dynamics-CRM) (MIT): form tools, navigation, impersonation, debugging, favorites, and custom commands. The side panel path remains available via the browser’s side panel UI.

Do **not** hand-edit `levelup/*`. Refresh upstream under `vendor/levelup`, then run:

```bash
npm run sync-levelup
```

(or `pwsh -File scripts/sync-levelup.ps1`). That rebuilds `levelup/`, remaps assets, and applies the SPO theme overlay while keeping the Power Platform purple look.

---

## Other page helpers

- **Recycle bin wait banner** — helps the “second stage recycle bin” quick link finish navigating after the first-stage page loads (see [docs/SECOND-STAGE-RECYCLE-BIN-TESTING.md](docs/SECOND-STAGE-RECYCLE-BIN-TESTING.md))
- **Admin center wait banner** — similar orchestration for certain admin-center hops
- **View formatter iframe bridge** + `rules/view-formatter-iframe.json` — lets many SharePoint views render inside the formatter preview

If a recent SharePoint site list is empty when you’re off SharePoint, the popup needs the **`history`** permission (already declared) and some prior visits to `*.sharepoint.com`.

---

## Development

- **Node.js 18+** — from the repo root: `npm test`
- View Manager REST logic lives in **`lib/viewsDataCore.mjs`** (unit tested) and is mirrored in injected **`getViewsData.js`** — keep those in sync after changes
- Owssvr vs REST fallback notes: [OWSSVR-LISTS.md](OWSSVR-LISTS.md)
- Ship a release: bump `manifest.json`, then `release/create-release.ps1` (see [release/CREATE-RELEASE-INSTRUCTIONS.md](release/CREATE-RELEASE-INSTRUCTIONS.md))

### Layout (high level)

| Area | Where |
|------|--------|
| Manifest / SW | `manifest.json`, `sw.js` → `background.js` + Level Up background |
| Compass UI | `content.js`, `compassToolkitPanels.js`, `compassUniversalSearch.js`, `compassColumnCreator.js` |
| Popup | `popup.html`, `popup.js`, `lib/popupUi.mjs` |
| View Manager / formatter | `views.html` / `views.js`, `view-formatter.html` / `view-formatter.js` |
| Injected page scripts | `exportCSV.js`, `permissionsMatrixExport.js`, `get*.js`, `createColumn.js`, … |
| Power Apps popout | `powerAppsToolkitPopout.js`, `levelup/` |
| Level Up source / sync | `vendor/levelup/`, `scripts/sync-levelup.ps1`, `scripts/levelup-adaptations/` |
| Shared libs / tests | `lib/`, `test/` |

---

## Credits

Power Apps / Dynamics tools incorporate [Level Up for Dynamics 365/Power Apps](https://github.com/rajyraman/Levelup-for-Dynamics-CRM) by Natraj Yegnaraman ([MIT](levelup/LICENSE)). Upstream copyright and license text are preserved in `levelup/LICENSE` and `vendor/levelup/LICENSE`.

## License

Use and modify as you like for the SPOToolkit portions. Level Up code remains under its MIT license as noted above.
