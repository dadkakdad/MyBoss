/** The works list is newest first. Reopen that one after a refresh. */
export function workToReopen(works) {
  if (!Array.isArray(works) || !works.length) return null;
  return works[0];
}

/** Label for a saved workbook in Account and in the Refresh menu. */
export function workMenuLabel(work) {
  if (!work?.name) return "";
  return work.sheetName ? `${work.name} · ${work.sheetName}` : work.name;
}

/**
 * Every saved workbook stays selectable from Refresh after newer uploads.
 * The works list is newest first; the open workbook is marked current.
 */
export function refreshSavedEntries(works, currentWorkId = "") {
  if (!Array.isArray(works)) return [];
  return works
    .filter((work) => work && work.id)
    .map((work) => ({
      id: work.id,
      label: workMenuLabel(work),
      current: work.id === currentWorkId,
    }));
}

/** File bytes and slides travel together on the first save. Later saves can send slides only. */
export function saveBody({ workId, fileName, fileBase64, sheetName, slides, active, includeFile }) {
  const body = {
    sheetName: sheetName || "",
    active: Number.isInteger(active) && active >= 0 ? active : 0,
  };
  if (slides !== undefined) body.slides = Array.isArray(slides) ? slides : [];
  if (includeFile || !workId) {
    body.name = fileName || "workbook.xlsx";
    if (fileBase64) body.dataBase64 = fileBase64;
  }
  return body;
}
