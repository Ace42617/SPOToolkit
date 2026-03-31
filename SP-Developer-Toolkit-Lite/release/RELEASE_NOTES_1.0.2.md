# SPO Toolkit 1.0.2

## Installation

1. Open **edge://extensions** or **chrome://extensions**
2. Enable **Developer mode**
3. Click **Load unpacked** and select the folder containing the extension files (the folder that contains `manifest.json`, `popup.html`, etc.)

**If you downloaded the .zip:** unzip it first, then select the unzipped folder in step 3.

## What's new in 1.0.2

- **List export with versions** – When "Include versions" is checked, the export no longer drops items when some files have very large version counts. Versions are fetched per item with full pagination, and if version fetch fails for an item, the current item is still included so row count never drops below the non-versioned export.
- **403 / security validation** – Better handling when SharePoint returns 403 or "security validation invalid":
  - **RPC view creation**: Longer initial wait (6s), more retries (5) with longer backoff, and a fresh request digest on retries so view creation is more reliable on first run in a session.
  - **View filter (ViewQuery)**: Uses a fresh digest every time and retries up to 3 times on 403/security-validation errors.
  - **View field setup**: RemoveAllViewFields and addViewField now use a fresh digest and retry on 403. If more than half of the field setup calls fail, the export stops with a clear error instead of continuing and producing a blank or wrong export.
- **Empty export guard** – If view setup fails and the export would produce no rows (or only headers), the extension now shows an error and does not download a blank file.
- **XLSX numbers** – Numeric values in Excel (XLSX) exports are now written as numbers so Excel no longer shows "Number Stored as Text" for integer and decimal columns.

## What's included

- Quick links, Page Properties, Columns, Views
- Reports: list/library export (CSV/Excel), folder count, path lengths, permissions, permissions matrix (single list or whole site)

See the [README](https://github.com/Ace42617/SPOToolkit/blob/main/README.md) for full details.
