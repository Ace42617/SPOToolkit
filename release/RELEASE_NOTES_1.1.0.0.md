# Release notes – SPO Developer Toolkit **1.1.0.0**

Major update on **1.0.7.0**: **Power Apps / Dynamics Level Up** tools in-page, SharePoint compass header polish, and export/matrix reliability fixes.

## Power Apps / Dynamics mode

- Popup switches to a purple **Power Apps mode** on Power Platform / Dynamics hosts.
- **In-page popout** (rocket FAB) mirrors the SharePoint compass launcher: shared position/size, fade open/close, purple header with app name + org/URL, display-mode menu, Early Riser / Night Owl toggle, and close **×**.
- Full toolkit UI from [Level Up for Dynamics 365/Power Apps](https://github.com/rajyraman/Levelup-for-Dynamics-CRM) (MIT): form actions, navigation, impersonation, debugging, favorites, and custom commands.
- Sidebar sections share equal header heights, collapse cleanly, and support **drag-to-reorder** (persisted).
- Rebuild path: `npm run sync-levelup` (overlays in `scripts/levelup-adaptations/`, output in `levelup/`).

## SharePoint compass

- Header layout aligned with Power Apps chrome (icons, fonts, theme pill).
- Close **×** on the compass panel (same control style as Power Apps).
- Smooth panel fade-in / fade-out kept as the reference animation for both surfaces.

## Exports & matrix

- JSZip load path hardened for SharePoint AMD (`jszip-preload` / `jszip-restore` / `jszip-restore-define`).
- Permissions matrix / report list plan updates and worker-lock tweaks.
- Column internal-name prediction encodes punctuation the SharePoint way (e.g. `A:B` → `A_x003a_B`).

## Install

1. Download **`SPOToolkit-1.1.0.0.zip`** from this release.
2. Unzip and load unpacked in `chrome://extensions` or `edge://extensions` (**Developer mode**).
3. If updating from an earlier build, click **Reload** on the extension card.

## Credits

Power Apps / Dynamics tools incorporate [Level Up for Dynamics 365/Power Apps](https://github.com/rajyraman/Levelup-for-Dynamics-CRM) by Natraj Yegnaraman (MIT). See `levelup/LICENSE`.
