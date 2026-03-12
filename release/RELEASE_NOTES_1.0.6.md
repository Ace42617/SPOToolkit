# Release notes – 1.0.6

## Summary of updates

### Not on SharePoint experience
- When the current tab is **not** a SharePoint Online page, the popup now hides all tabs and shows a single message panel.
- Message: *"Navigate to a SharePoint Online site or select a recently visited site below."*
- **Recently visited SharePoint sites** dropdown: populated from browser history and from sites where the extension was opened; one entry per site (deduped by site base URL). Labels show site name (e.g. `hr — 823cwd.sharepoint.com`).
- Selecting a site navigates the current tab to that URL and closes the popup; reopening on that tab shows the full toolkit.

### View Manager (views.html / views.js)
- **Export column names to CSV**: button in the actions bar; filename `sitename-libraryname-columns.csv`.
- Dark toggle and "Back to SharePoint" link aligned with popup styling.
- Sort/Filter spacing and per-field hover (border, background, green text); select/input hover styling.
- Removed row-level hover on sort/filter rows.

### Popup
- **Field hover**: Toolbar text inputs and all selects use the same hover style (border, background, green text) in both popup and View Manager.
- **Quick Links duplicate fix**: Re-entry guard prevents duplicate "Edit user profile", "Login as another user", "MaintenanceMode", etc., when the popup is shown or tab is restored.
- **Refinables tab**: Renamed from "Refinable Props" to "Refinables".

### Permissions
- Added **history** permission to support the recently visited SharePoint sites list from browser history.

### Cleanup and refactor
- Removed unused files: **searchQuery.js**, **views_from_ct.html** (temporary).
- Popup: re-entry guard for Quick Links; refinable mappings HTML built with template literals; `var` → `let`/`const` and minor loop simplifications in popup.js.

---

*No tab content sliding animation; tab bar uses animated underline (0.16s) and instant content switch.*
