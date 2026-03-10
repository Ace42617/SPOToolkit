# Changes since last commit (v1.0.6)

## Site Contents launcher (floating button + panel)

- **Floating button** (bottom-right on SharePoint pages): compass icon, light neutral background; dark mode uses dark background with light icon. Hover: single spin; click: triple spin then opens panel. Toggle in popup Settings: "Show lists & libraries icon on SharePoint pages."
- **Panel**: Filter row above column headers; columns Name, Type, Items, Modified, Settings (slider icon per row linking to list/library settings). Auto-sort by Type on load. Header shows **Site name | list/library name | Contents** (site name bold, 14px); below it a subtle **tenant name - Tenant URL** line (11px, deduped on theme switch). Header and table content left-aligned (12px). Recycle bin icon: simple trash can. Per-row settings icon: slider (same as "Current list/library settings" in header).

## Popup

- **Settings screen**: Gear in header opens Settings; Back returns to main. Settings include "Show lists & libraries icon on SharePoint pages" checkbox and version at bottom. Popup fade-in on open; when Site Contents panel is opened from the page, popup fades out and closes.
- **Messages**: Popup notifies content script on open/close (`SPOToolkitPopupOpened` / `SPOToolkitPopupClosed`) so launcher can roll down/up; content script sends `SPOToolkitPanelOpened` when panel opens.

## Manifest

- **web_accessible_resources**: Added `getSiteLists.js` and `icon.png` for content script (Site Contents launcher and compass icon).

## New file

- **getSiteLists.js**: Injected on SharePoint pages to fetch site title, tenant name, and lists (Id, Title, DefaultViewUrl, BaseTemplate, ItemCount, LastItemModifiedDate); posts `SPCSVSiteListsResult` and `SPCSVSubsitesResult`. Used by content.js to build the Site Contents panel.

## Refactor

- **content.js**: Tenant line logic unified into `applyTenantLine(container, text)` so a single helper updates or creates the tenant/URL line and removes duplicates (fixes duplicate line when switching themes).
