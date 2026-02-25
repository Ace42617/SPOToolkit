# owssvr.dll and SharePoint Lists – Troubleshooting

## Why owssvr can fail with lists

owssvr.dll is a legacy RPC handler. It behaves differently across list types and tenants:

| Issue | Cause | Fix in this extension |
|-------|-------|------------------------|
| **404** | Some sites return 404 when using a custom View with owssvr (list-dependent) | Automatically falls back to REST API |
| **0 rows** | ViewQuery (ID range filter) not applied by owssvr for some list types | Falls back to REST API |
| **XML parse error** | Unexpected or malformed owssvr response | Falls back to REST API |

## REST fallback

When owssvr fails on the first page (404, 0 rows, or parse error), the extension automatically falls back to the **REST API** (`/_api/web/lists(...)/items`). REST works reliably for both lists and document libraries.

You’ll see progress messages such as:
- `owssvr 404 – falling back to REST API…`
- `owssvr returned 0 rows – falling back to REST API…`
- `owssvr parse error – falling back to REST API…`

## Lists vs document libraries

- **Document libraries** (BaseTemplate 101): Typically work well with owssvr; FileLeafRef, FSObjType, etc. are available.
- **Custom lists** (BaseTemplate 100): May not support owssvr in the same way; REST fallback is used when owssvr fails.

## If export still fails

1. **Refresh the SharePoint page** (F5) and try again.
2. **Check the browser console** (F12 → Console) for error details.
3. **Verify permissions** – you need read access to the list.
4. **REST 400** – If REST fails with 400, a requested field may not exist on the list (e.g. FSObjType on some lists).
