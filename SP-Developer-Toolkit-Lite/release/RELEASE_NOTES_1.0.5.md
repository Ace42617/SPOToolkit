# SPO Toolkit 1.0.5

## Installation

1. Open **edge://extensions** or **chrome://extensions**
2. Enable **Developer mode**
3. Click **Load unpacked** and select the folder containing the extension files (the folder that contains `manifest.json`, `popup.html`, etc.)

**If you downloaded the .zip:** unzip it first, then select the unzipped folder in step 3.

## What's new in 1.0.5

- **Page Properties on modern pages** – Page Properties now works on site home and other modern SharePoint pages. When `_spPageContextInfo` is not available, the extension falls back to fetching the current page with `?as=json` (in page context) and displays the flattened context (e.g. Web.WebAbsoluteUrl, Web.Title). Context is also read from all frames so it works when the list is in an iframe.
- **Columns tab** – Removed the Managed Property column. The inferred crawled property column is now labeled **Crawled Property (Inferred)**. For lookup columns whose internal name uses `_x003a__x0020_` (e.g. `Advocate_x003a__x0020_Email`), the inferred crawled property now uses a colon: `ows_Advocate:_x0020_Email`.
- **Export** – Temporary RPC view used for list/library export is always deleted when done (success or error), using a try/finally so cleanup runs even if export throws. View creation retries on 403/503 are no longer shown to the user.

## What's included

- Quick links, Page Properties, Columns, View Manager, Reports
- Page Properties: works on list/library, classic, and modern pages (including site home)
- Reports: list/library export (CSV/Excel), folder count, path lengths, permissions matrix

See the [README](https://github.com/Ace42617/SPOToolkit/blob/main/README.md) for full details.
