# SPO Toolkit 1.0.1

## Installation

1. Open **edge://extensions** or **chrome://extensions**
2. Enable **Developer mode**
3. Click **Load unpacked** and select the folder containing the extension files (the folder that contains `manifest.json`, `popup.html`, etc.)

**If you downloaded the .zip:** unzip it first, then select the unzipped folder in step 3.

## What's new in 1.0.1

- **Export filenames** – All exports now use `sitename_listname_datetime.extension` (e.g. `MySite_Documents_2025-02-20_14-30.csv`).
- **XLSX export** – Lookup/user fields (e.g. Author, Editor) no longer show raw `;#` values; display values match CSV.
- **RPC view** – Waits for the page to be ready and retries view creation on 403/503 to reduce "create RPC view failed" errors.
- **Permissions matrix – whole site** – New option **Include whole site** adds a report across all lists and libraries with columns **List/Library/Page name** and **Type** (List, Library, Page).

## What's included

- Quick links, Page Properties, Columns, Views
- Reports: list/library export (CSV/Excel), folder count, path lengths, permissions, permissions matrix (single list or whole site)

See the [README](https://github.com/Ace42617/SPOToolkit/blob/main/README.md) for full details.
