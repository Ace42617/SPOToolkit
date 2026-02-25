# SharePoint Toolkit

Chrome/Edge extension for SharePoint: export list or library data, run reports, manage views, and access quick links.

## Install

1. Open `edge://extensions` or `chrome://extensions`
2. Enable **Developer mode**
3. Click **Load unpacked** and select this folder

## Use

1. Open a SharePoint list or library view (or any SharePoint page for quick links)
2. Click the extension icon
3. **Quick links** – Site settings, Recycle bin, Admin center, etc.
4. **Page Properties** – View/copy `_spPageContextInfo` from the current page
5. **Columns** – Browse list columns (name, internal name, type); shown when on a list/library page
6. **Views** – Edit or create list views (columns, sort, filter, group by); shown when on a list/library page
7. **Reports** – Pick a report and click Export

## Reports

- **Download list/library contents** – CSV or Excel (XLSX/XLS), optional column selection and versions
- **Folder count** – Nested folder counts (document libraries only)
- **Path lengths** – Path length report (document libraries only)
- **List/library permissions** – Who has access and at what level
- **Permissions matrix** – Item-level matrix with inheritance and “Given through” (explicit vs inherited)

## Settings

- **Page size** – Items per request (500–20,000)
- **Include versions** – Toggle for version history in main export

## Files

| File | Purpose |
|------|--------|
| `manifest.json` | Extension config and permissions |
| `popup.html` / `popup.js` | Popup UI (tabs, reports, column picker, settings) |
| `content.js` | Injects scripts into SharePoint page; progress bar; message handling |
| `exportCSV.js` | Export and report logic (REST, owssvr fallback, permissions, matrix) |
| `checkListPage.js` | Injected: detects list/library page |
| `getFields.js` | Injected: list fields for column picker |
| `getSearchSchema.js` | Injected: search schema columns |
| `getViewsData.js` | Injected: views and view details |
| `saveView.js` | Injected: create/update list view |
| `jszip*.js` | XLSX export (preload, library, restore) |

See **OWSSVR-LISTS.md** for owssvr vs REST fallback behavior.

## License

Use and modify as you like.
