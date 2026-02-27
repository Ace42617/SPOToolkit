# SPO Toolkit 1.0.3

## Installation

1. Open **edge://extensions** or **chrome://extensions**
2. Enable **Developer mode**
3. Click **Load unpacked** and select the folder containing the extension files (the folder that contains `manifest.json`, `popup.html`, etc.)

**If you downloaded the .zip:** unzip it first, then select the unzipped folder in step 3.

## What's new in 1.0.3

- **View Manager (full page)** – Dedicated View Manager tab opens a full-page editor: manage list views via REST (list views, create/update/delete/duplicate), view settings (name, personal, row limit, scope), columns (search, add/remove all, drag to reorder), sort, filter (type inferred from column), group by. Set as default view. Layout: 50/50 columns vs settings, vertical fill.
- **Popup** – Views tab removed; View Manager tab is a single “View Manager” label and “Open View Manager” button.
- **Filters** – View Manager filter conditions no longer ask for data type; type is inferred from the selected column (like standard SharePoint).
- **Refactor** – Removed unused `saveView.js`; content script uses shared `injectAndWait` for getPageContext; simplified REST handler; View Manager CAML uses `escapeAttr` helper.

## What's included

- Quick links, Page Properties, Columns, Reports
- View Manager (full page): open from popup when on a list/library page
- Reports: list/library export (CSV/Excel), folder count, path lengths, permissions, permissions matrix

See the [README](https://github.com/Ace42617/SPOToolkit/blob/main/README.md) for full details.
