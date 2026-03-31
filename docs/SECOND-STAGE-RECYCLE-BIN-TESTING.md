# Testing the “Second Stage Recycle Bin” quick link

The extension cannot log into your tenant from this repo, so validation is manual on **your** SharePoint Online site. Automated checks here only cover **URL string construction** (`npm test` → `test/recycleBinUrls.test.mjs`).

## 1. Tenant / permissions

1. Use any **SharePoint Online** site collection where your account is a **Site Collection Administrator** (or Global/SharePoint admin).  
   - Free option: [Microsoft 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program) (E5 dev tenant).
2. Second-stage recycle bin is **admin-only**. Without that role, SharePoint may redirect or show access denied even if the URL is correct.

## 2. Load the extension (unpacked)

1. In Chrome or Edge: **Extensions** → enable **Developer mode** → **Load unpacked**.
2. Choose the extension folder:
   - Full toolkit: repo root (folder that contains `manifest.json`).
   - Lite: `SP-Developer-Toolkit-Lite` (if that is your manifest location).

## 3. Manual checks

### A. Site collection root URL

1. Open a **subsite** under a site collection (not only the root web), e.g.  
   `https://<tenant>.sharepoint.com/sites/<site>/subsite`
2. Open the extension popup → **Quick links** → **Second Stage Recycle Bin**.
3. Before clicking, **right‑click the link → Copy link address** (or hover and note the URL).
4. Expected:
   - Host and path match the **site collection root** (same as opening the root site in the browser), **not** the subsite URL.
   - Final admin URL (also used for **Copy link address**):  
     `/_layouts/15/AdminRecycleBin.aspx?view=5#view=13`  
     at the **site collection root** (not `/_layouts/15/RecycleBin.aspx`).
   - **Clicking** the quick link navigates the **current tab** to `RecycleBin.aspx` for the current web, then — after that page finishes loading — to the admin URL above in the **same tab**.

### B. Page you land on

1. Click the quick link.
2. You should land on the **site collection admin recycle bin** UI, **second stage** (items deleted from the end-user recycle bin), not the personal first-stage bin only.

If you still see the wrong view:

1. In the same browser session, open **Site settings** → **Site collection administration** → **Recycle bin**, scroll to the bottom, and use the built-in **second-stage recycle bin** link once.
2. Compare the URL in the address bar after that navigation to the copied quick-link URL (same site root + same layout page pattern).

## 4. Automated tests (URL building only)

From the repo root (Node 18+):

```bash
npm test
```

This runs `test/recycleBinUrls.test.mjs` and does **not** call SharePoint.

## 5. If `_spPageContextInfo` is missing

Some minimal or non-SharePoint pages may not inject page context. The popup then falls back to `/_api/site/rootweb?$select=Url`, then to the URL parsed from the tab. If you test only from a **standard modern team/communication site page** or a **list/library view**, context is usually present.

---

## 6. How to “set it up” so the AI in Cursor can help debug (no shared passwords)

The assistant **cannot** log into Microsoft 365 with your account. There is no supported way to hand over credentials so it can “keep trying until it works” inside your tenant. You can still make progress by giving it **evidence** from your browser or from Cursor’s browser tools.

### A. Paste URLs and context (fastest)

Do these in order and paste all four into the chat (you may redact the hostname if you keep the **path and query identical**, e.g. replace `contoso` with `tenant`):

1. **Quick link target** — In the extension popup, right‑click **Second Stage Recycle Bin** → **Copy link address**.
2. **Known-good URL** — Use Site settings → Site collection Recycle bin → bottom link **second-stage recycle bin**. After it opens, copy the **full address bar** URL (including `#…` if present).
3. **Page context** — On the SharePoint tab where you opened the popup, press **F12** → **Console**, run:

   ```js
   ({
     href: location.href,
     web: window._spPageContextInfo?.webAbsoluteUrl,
     site: window._spPageContextInfo?.siteAbsoluteUrl
   })
   ```

   Copy the printed object (expand/copy as JSON if needed).

4. **What you see** — One line: e.g. “lands on first-stage”, “access denied”, “blank page”, “wrong site”.

With (1) vs (2) the assistant can see whether the extension is building the wrong root, wrong layout page, or wrong query/hash.

### B. Cursor Browser MCP (if your project has it)

If Cursor’s **browser** integration is enabled for your workspace:

1. You open a browser tab from Cursor, **you** sign in to SharePoint (the assistant does not get your password).
2. You navigate until you are on **second stage recycle bin** using the **official UI** link.
3. Ask the assistant to take a **snapshot** of the page; it can read the **current URL** from the tooling output.
4. Repeat after opening the **quick link** in a new tab (same session).

Then the assistant can compare URLs character-by-character and adjust the extension.

### C. What does *not* work

- Sending **passwords**, **client secrets**, or **refresh tokens** in chat (unsafe; do not do this).
- Expecting the assistant to run Playwright against `login.microsoftonline.com` with your credentials from this environment (not supported as a secure pattern).

### D. Optional: HAR export

In Edge/Chrome **DevTools** → **Network**, reproduce: click the broken quick link once, then the working footer link once → **Export HAR** → share only the **document** request URLs (or the full HAR if your org allows it). That shows redirects and final `Location` headers.
