# Release notes – SPO Developer Toolkit **1.0.6**

*Compared to **1.0.5**.*

## Highlights

- **Quick Links** – Filter bar, **site content types** (`SiteAdmin.aspx#/contentTypes`), **tenant** content types in admin center, sticky filter while scrolling.
- **Columns** – Click a column **display name** to open **column settings** (`FldEdit.aspx`) in a new tab; **Type** column left-aligned; **refresh** icon on the toolbar.
- **Page Properties** – **Refresh** as an icon button (same style as Columns).
- **Popup** – **Send feedback** in the header (Microsoft Forms); **Settings** gear (e.g. lists launcher); codebase uses **`lib/popupUi.mjs`** for shared column/link/filter helpers.
- **When not on SharePoint** – Dedicated panel with **recent SharePoint sites** (history + extension usage); **`history`** permission added.
- **View Manager** – **Export column names to CSV**; REST pipeline aligned with tested **`lib/viewsDataCore.mjs`**; layout and field UI polish.
- **Page experience** – **Site Contents** launcher affordance on SharePoint pages (when enabled in Settings).
- **Misc** – **Refinables** tab label (was “Refinable Props”); Quick Links re-entry guard; `npm test` includes filter-typeahead tests.

---

## Quick Links

- **Filter links…** – Live filter on link title, URL, and section heading (Current Site, Current User, Page Modes, Tenant Admin). Toolbar stays visible while scrolling (`position: sticky`).
- **Site content types** – Opens **`/_layouts/15/SiteAdmin.aspx#/contentTypes`** (modern site content type gallery).
- **Tenant content types** – SharePoint admin center hash route to content types.
- **SharePoint Admin Settings** – Deep link to site in admin center when site ID is available.
- Duplicate tiles no longer appear when reopening the popup or restoring the tab (re-entry guard).

## Columns (list / library)

- Lists **fields** from the current list (REST): display name, internal name, inferred crawled property name, type.
- **Display name** is a link to classic sharePoint column edit (**`FldEdit.aspx`**) for that list field, in a **new** tab.
- Toolbar uses a **refresh** icon (reload column list) instead of a “Load” label.
- **Type** column and header are **left-aligned**.

## Page Properties

- **Refresh** is an icon button (reloads `_spPageContextInfo` / context payload).

## Popup & settings

- **Feedback** – Header link to the toolkit feedback form (`forms.cloud.microsoft`).
- **Settings** (gear) – e.g. **Show lists & libraries icon** on SharePoint pages; version string.
- **Dark mode** toggle unchanged (Early Riser / Night Owl).

## Not on a SharePoint tab

- Tabs are hidden; message explains how to proceed.
- **Recently visited SharePoint sites** – Dropdown from browser history and sites where the extension was used; deduped by site base; choosing an entry navigates the current tab.

## View Manager (`views.html`)

- **Export column names to CSV** in the actions area (filename pattern includes site/library).
- Sort/filter row styling and hover behavior aligned with the popup.
- View/load logic kept in sync with unit-tested **`lib/viewsDataCore.mjs`**.

## Site Contents launcher

- Optional **lists & libraries** launcher on SharePoint pages (toggle in Settings).

## Technical / developer

- **`lib/popupUi.mjs`** – `listColumnSettingsUrl`, quick-links DOM filter, column filter matching; covered by **`test/popupUi.test.mjs`**.
- **Popup script** is an ES module (`<script type="module" src="popup.js">`) loading **`./lib/popupUi.mjs`** (Lite build uses **`../lib/popupUi.mjs`**).
- **`getSearchSchema.js`** – Returns **`siteUrl`** and **`listId`** with the column list for FldEdit links.

## Cleanup

- Removed unused **`searchQuery.js`** and **`views_from_ct.html`**.

---

## Install (load unpacked)

1. Download **`SPOToolkit-1.0.6.zip`** from this release (or clone this repo).
2. Unzip if needed.
3. In **edge://extensions** or **chrome://extensions**, enable **Developer mode** → **Load unpacked** → select the folder that contains **`manifest.json`**.
