# Release notes – SPO Developer Toolkit **1.1.1.0**

Follow-up to **1.1.0.0**: Everything Bagel report, Site Contents tree, export worker polish, and repo cleanup.

## Reports

- **Everything Bagel** — one Excel workbook for the selected sites/lists/libraries with Permissions Matrix, Group Members, Sharing Links, All Items, **Folder Counts**, and **Path Lengths**.
- Same list picker and matrix settings as the permissions matrix; always downloads `.xlsx`.

## Site Contents

- Expandable **subsites** in Site Contents (soft insert/remove, depth labels).
- Subsite rows include site settings and site permissions shortcuts.
- New `getWebContents.js` REST helper for nested web contents.

## Export worker tab

- Everything Bagel gets hyped banner copy (legendary whole-schmear / pack-a-lunch wait messaging).
- Optional **Send me bragging rights** checkbox (off by default): on successful export, opens the default mail app with Auto-Snake scores and **To:** set to your SharePoint email.

## Other

- Export console / progress UX polish; column creator Cancel; Quick Links scroll aligned with Site Contents.
- Removed unused Experimental/Lite trees, docs, and local test clutter from the main package path.

## Install

1. Download **`SPOToolkit-1.1.1.0.zip`** from this release.
2. Unzip and load unpacked in `chrome://extensions` or `edge://extensions` (**Developer mode**).
3. If updating from an earlier build, click **Reload** on the extension card.

## Credits

Power Apps / Dynamics tools incorporate [Level Up for Dynamics 365/Power Apps](https://github.com/rajyraman/Levelup-for-Dynamics-CRM) by Natraj Yegnaraman (MIT). See `levelup/LICENSE`.
