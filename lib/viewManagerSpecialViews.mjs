/**
 * View Manager can only rebuild standard list-view CAML (Where / OrderBy / GroupBy)
 * plus ViewFields. Calendar, board, gallery, and Gantt views store a different
 * query/layout; PATCHing ViewQuery and RemoveAllViewFields would destroy them.
 */

const SPECIAL_VIEW_TYPE2 = {
  kanban: "board",
  tiles: "gallery",
  gallery: "gallery",
  moderncalendar: "calendar",
  calendar: "calendar",
  gantt: "gantt",
};

export function classifySpecialView(input) {
  const src = input || {};
  const t2 = String(src.viewType2 || "").trim().toLowerCase();
  if (t2 && SPECIAL_VIEW_TYPE2[t2]) return SPECIAL_VIEW_TYPE2[t2];

  const query = String(src.viewQuery || "");
  if (/DateRangesOverlap/i.test(query)) return "calendar";

  const cal = src.calendarSettings;
  if (cal != null && String(cal).trim() !== "") return "calendar";

  const viewData = String(src.viewData || "");
  if (/Type\s*=\s*"Gantt/i.test(viewData) || /Gantt(?:Title|Start|End)/i.test(viewData)) {
    return "gantt";
  }
  return "";
}

export function specialViewSaveBlockMessage(kind) {
  if (!kind) return "";
  return (
    "View Manager cannot save " +
    kind +
    " views. Saving would replace this view's query and columns with a standard list-view definition, which breaks the calendar/board layout. Select a standard list view, or choose Create new view."
  );
}
