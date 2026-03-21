# Changes since v1.0.6

## Refactor (View Manager + Site Contents) — 2025-02-27

### View Manager data loading

- **`lib/viewsDataCore.mjs`** — Single place for testable logic: site URL resolution (`webAbsoluteUrl` from extension vs `_spPageContextInfo`), forced list ID vs current list, OData **`@odata.nextLink`** paging for `/views` and `/fields`, view query parsing, board-scaffold field filtering, stub fields for view editor.
- **`getViewsData.js`** — Injected script refactored to an async `run()` with the same pipeline as the core module (classic IIFE, no bundler). **`SP-Developer-Toolkit-Lite/getViewsData.js`** kept identical (except header comment).
- **`npm test`** — `node --test` on `test/viewsDataCore.test.mjs` (requires Node 18+). **`package.json`** added with `engines.node` and test script only (no build step).

### Content script

- **`attachSpViewsParamsScript(message)`** — Centralizes JSON injection for `#sp-views-params` (`viewId`, `listId`, `webAbsoluteUrl`) used by `getViewsData`. Applied in root **`content.js`** and **`SP-Developer-Toolkit-Lite/content.js`**.

### Site Contents panel (dark theme on admin pages)

- **Dark panel overrides** — Explicit light text colors with `!important` where SharePoint host CSS overrode inherited colors; `color-scheme: dark` and `isolation: isolate` on the dark panel (root + Lite `content.js`).

### `views.js` (extension page)

- **`contextWebAbsoluteUrl`** — Set in `showContext()`; passed with `getViewsData` messages so the injected script targets the correct web when View Manager is opened off-list.

### Documentation

- **`README.md`** — Development & tests section; `lib/viewsDataCore.mjs` row in file table.

---

## Earlier entries (pre-refactor snapshot)

### Site Contents launcher (floating button + panel)

- **Floating button** (bottom-right on SharePoint pages): compass icon, light neutral background; dark mode uses dark background with light icon. Hover: single spin; click: triple spin then opens panel. Toggle in popup Settings: "Show lists & libraries icon on SharePoint pages."
- **Panel**: Filter row above column headers; columns Name, Type, Items, Modified, Settings (slider icon per row linking to list/library settings). Auto-sort by Type on load. Header shows **Site name | list/library name | Contents** (site name bold, 14px); below it a subtle **tenant name - Tenant URL** line (11px, deduped on theme switch). Header and table content left-aligned (12px). Recycle bin icon: simple trash can. Per-row settings icon: slider (same as "Current list/library settings" in header).

### Popup

- **Settings screen**: Gear in header opens Settings; Back returns to main. Settings include "Show lists & libraries icon on SharePoint pages" checkbox and version at bottom. Popup fade-in on open; when Site Contents panel is opened from the page, popup fades out and closes.
- **Messages**: Popup notifies content script on open/close (`SPOToolkitPopupOpened` / `SPOToolkitPopupClosed`) so launcher can roll down/up; content script sends `SPOToolkitPanelOpened` when panel opens.

### Manifest

- **web_accessible_resources**: Added `getSiteLists.js` and `icon.png` for content script (Site Contents launcher and compass icon).

### New file

- **getSiteLists.js**: Injected on SharePoint pages to fetch site title, tenant name, and lists (Id, Title, DefaultViewUrl, BaseTemplate, ItemCount, LastItemModifiedDate); posts `SPCSVSiteListsResult` and `SPCSVSubsitesResult`. Used by content.js to build the Site Contents panel.

### Refactor

- **content.js**: Tenant line logic unified into `applyTenantLine(container, text)` so a single helper updates or creates the tenant/URL line and removes duplicates (fixes duplicate line when switching themes).
