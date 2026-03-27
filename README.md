# SharePoint Developer Toolkit

Chrome/Edge extension for **SharePoint Online**: export list or library data, reports, **View Manager**, refinable managed-property shortcuts, and **Quick Links**.

## Install

1. Open `edge://extensions` or `chrome://extensions`
2. Enable **Developer mode**
3. Click **Load unpacked** and select the folder that contains **`manifest.json`**

*(Release builds: unzip the attached `.zip` from [Releases](https://github.com/Ace42617/SPOToolkit/releases) and load the unzipped folder.)*

## Use

1. Open a SharePoint page (list/library view for exports, View Manager, and full **Columns** tab).
2. Click the extension icon.

### When you are not on SharePoint

The popup shows a short message and a **Recently visited SharePoint sites** dropdown (from history and past extension use). Pick a site to navigate the current tab there.

### Quick Links

- **Current Site** – Settings, Site contents, **Site content types** (modern site admin `#/contentTypes`), Recycle bin, People, Storage metrics, **SharePoint admin** site details (when site ID is known).
- **Current User** – Edit profile, sign in as another user.
- **Page Modes** – Maintenance mode, WebView / WebViewList, disable SPFx third-party code, web part maintenance.
- **Tenant Admin** – Admin center, tenant settings, user profiles, term store, search administration, API access, Teams admin, app catalog(s).
- **Filter** – Sticky search box at the top filters by link text, URL, or section title.

### Page Properties

- Flattened **page / list context** (from `_spPageContextInfo` and related payloads).
- Filter by name or **by value**; copy values.
- **Refresh** (icon) reloads context from the page.

### Columns *(list or library page)*

- Table of **display name**, **internal name**, **crawled property** name, **type**.
- Filter columns with the search field.
- **Display name** opens **column settings** (`FldEdit.aspx`) in a **new** tab.
- **Refresh** (icon) reloads the field list.

### Views

- Opens the **View Manager** full page: edit/create views, columns, sort, filter, group; **export column names to CSV**.

### Refinables

- Jump to **refinable managed property** pages (e.g. `RefinableString00`) for search schema work.

### Reports

- **Download list/library** – CSV or Excel (XLSX), optional column picker and versions.
- **Folder count**, **Path lengths** (libraries).
- **Permissions matrix** – Item-level matrix with inheritance and “Given through”.

## Settings *(gear in popup)*

- **Show lists & libraries icon on SharePoint pages** – Site Contents–style launcher.
- Extension version.
- **Dark mode** – Header toggle (Early Riser / Night Owl).

### Other permissions

- **`history`** – Powers the “recent SharePoint sites” list when you are not on a SharePoint tab.

## Feedback

- **Send feedback** (speech bubble) in the popup header links to the Microsoft Forms feedback flow.

## Reports (detail)

- **Download list/library contents** – CSV or Excel (XLSX); default XLSX; optional column selection and version history.
- **Folder count** – Nested folder counts (document libraries).
- **Path lengths** – Path length report (document libraries).
- **Permissions matrix** – Item-level matrix with inheritance and explicit vs inherited.

## Files (main extension)

| Area | Files |
|------|--------|
| Config | `manifest.json` |
| Popup | `popup.html`, `popup.js` (ES module), `lib/popupUi.mjs` (shared column / quick-links helpers) |
| Full-page views | `views.html`, `views.js`, `filterTypeaheadLogic.js` |
| Page bridge | `content.js`, `background.js` |
| Injected scripts | `exportCSV.js`, `getFields.js`, `getSearchSchema.js`, `getViewsData.js`, `checkListPage.js`, `getListType.js`, `getRefinableMappings.js`, `getPageContext.js`, `getPageContextJson.js`, `getSiteLists.js`, … |
| ZIP / XLSX | `jszip*.js` |
| Facts / feedback constant | `sharepoint-facts.js` |

- **View Manager REST** logic lives in **`lib/viewsDataCore.mjs`** (unit tests) and is mirrored in page-injected **`getViewsData.js`**.

See **OWSSVR-LISTS.md** for owssvr vs REST fallback behavior.

## Development & tests

- Install [Node.js](https://nodejs.org/) **18+**, then from the repo root run **`npm test`** (`node --test` over `test/**/*.mjs` and related tests).
- After changing **`lib/viewsDataCore.mjs`**, keep **`getViewsData.js`** (and Lite copy if present) in sync.

## License

Use and modify as you like.
