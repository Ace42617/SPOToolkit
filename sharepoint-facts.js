/**
 * SharePoint facts – shown when the header compass icon is clicked.
 *
 * HOW TO EDIT:
 * - Add a new fact: add an object to the array below.
 * - Update a fact: change the "text" and/or "sourceUrl".
 * - Remove the Learn More button: omit "sourceUrl" or set it to "".
 *
 * FORMAT (one per line in the array):
 *   { text: "Your fact here.", sourceUrl: "https://learn.microsoft.com/..." }   ← with Learn More button
 *   { text: "Fact with no link." }   ← no Learn More button (omit sourceUrl)
 *   "Plain string fact."   ← also allowed; no Learn More button
 *   { text: "...", openUrl: "https://..." }   ← opens openUrl in a new tab when this fact is picked (optional)
 */
window.SHAREPOINT_FACTS = [
  { text: "SharePoint Online runs on the Microsoft 365 cloud and uses the same infrastructure as Teams and OneDrive.", sourceUrl: "https://learn.microsoft.com/en-us/sharepoint/introduction" },
  { text: "Document libraries in SharePoint support co-authoring: multiple people can edit the same file at once.", sourceUrl: "https://learn.microsoft.com/en-us/sharepoint/introduction" },
  { text: "SharePoint list views can be personalized per user without changing the shared view.", sourceUrl: "https://learn.microsoft.com/en-us/microsoft-365/community/creating-useful-views-in-lists-libraries" },
  { text: "The SharePoint REST API and _api endpoint work with the same authentication as the browser session.", sourceUrl: "https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints" },
  { text: "SharePoint Framework (SPFx) lets you build web parts and extensions that run inside SharePoint and Teams.", sourceUrl: "https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview" },
  { text: "List and library IDs in SharePoint are GUIDs; they appear in the URL when you open a view's settings or use the 'Page Props' tab of this extension.", sourceUrl: "https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints" },
  { text: "Version history is available on document libraries and can be enabled on lists for major/minor versions.", sourceUrl: "https://learn.microsoft.com/en-us/sharepoint/faqs-for-versions" },
  { text: "SharePoint search uses the same index as Microsoft Search; refiners map to managed properties.", sourceUrl: "https://learn.microsoft.com/en-us/sharepoint/" },
  { text: "Power Automate and Power Apps can use SharePoint as a data source with connectors.", sourceUrl: "https://learn.microsoft.com/en-us/connectors/sharepointonline/" },
  { text: "Site collections have a recycle bin and a second-stage recycle bin before permanent deletion.", sourceUrl: "https://learn.microsoft.com/en-us/sharepoint/" },
  { text: "Never trust a file named 'Final.' There will be a Final_Final_v2." },
  { text: "The recycle bin is not a backup strategy. It's just a second chance at regret." },
  { text: "A view without a filter is just a cry for help." },
  { text: "If the URL has %20 in it, you already know what happened." },
  { text: "If users can type in it, they will." },
  { text: "If a file name contains the word 'NEW,' it is at least three years old." },
  { text: "Breaking inheritance is easy. Fixing it is character building." },
  { text: "If you ignore the 5,000 Item threshold, it will not ignore you." },
  { text: "Documentation is the difference between architecture and folklore." },
  { text: "Never gonna give you up... 🎵", openUrl: "https://ia801509.us.archive.org/10/items/Rick_Astley_Never_Gonna_Give_You_Up/Rick_Astley_Never_Gonna_Give_You_Up.mp4" },
];

/** Microsoft Forms link for extension feedback (popup, page UI, content script). */
window.SPOTOOLKIT_FEEDBACK_URL = "https://forms.cloud.microsoft/r/HtmARYViZP";
