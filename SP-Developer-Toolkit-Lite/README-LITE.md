# SP Developer Toolkit Lite

Read-only fork of SPO Dev Toolkit. Same read functionality; **no write functionality**.

## What's included (read-only)

- **Quick Links** – navigate to SharePoint sites, user links, page modes, tenant admin
- **Page Props** – view page context / ContextInfo
- **Columns** – search schema / managed metadata columns (read)
- **Reports** – Export list/library to CSV/Excel, Folder & Item Counts, Path Length Report, Permissions Matrix (no column picker; export uses the existing RPC view’s columns)
- **Site Contents launcher** – bottom-right launcher to open Site Contents / Subsites (read-only list data)

## What's removed

- **View management** – No Views tab; no creating, editing, or deleting list views
- **Refinables tab** – No refinable managed property configuration
- **No PATCH/POST/DELETE** – Content script only allows GET requests (no digest, no view create/update/delete)
- **Column picker** – The “Choose Columns” button is removed in Lite. You cannot select columns from the extension; the export uses whatever columns the existing RPC view has.
- **Reports (list/library)** – Do **not** create or modify an RPC view. You must **create an "RPC" view yourself** on each list/library (e.g. Scope: RecursiveAll, include the columns you want). Export uses that existing view and does a **single-page** export (one owssvr fetch). No view creation, no view deletion, no view query or column changes.

## RPC view for reports

For **Export List/Library**, **Folder & Item Counts**, and **Path Length Report**, the extension uses a **shared** view named **"RPC"** (any scope, e.g. RecursiveAll 1 or 2).

- Create the view manually: **List Settings > Views > Create view**. Name it **RPC**, save as a **shared** view, and set **Scope** to **RecursiveAll** (show all items) if you want all items/folders. Lite does not validate Scope or filter.
- Lite uses this view as-is and does **one page** of export (view’s row limit). For larger lists, use the full toolkit or create a view with a high row limit.
- If owssvr returns 404, Lite tries the list’s **default view** and, if that works, exports using it and tells you.
- If both RPC and default view return 404, the site may not support the legacy owssvr endpoint; use the full SPO Dev Toolkit or PowerShell/Graph.

## Loading the extension

Load the **SP-Developer-Toolkit-Lite** folder as an unpacked extension in `chrome://extensions` (or your browser’s equivalent). Do not load both the full toolkit and Lite on the same profile to avoid conflicts.
