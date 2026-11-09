/**
 * Helpers for the in-browser spreadsheet editor. Pure functions, no DOM.
 * A "draft" is an array of rows, each row an array of at least COLUMN_LABELS.length cells.
 */

import { COLUMN_LABELS, buildOrg, flagOn } from "./org.js";

/** Text shown in an input for a cell. Dates become yyyy-mm-dd. */
export function cellText(value) {
  if (value == null) return "";
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? "" : value.toISOString().slice(0, 10);
  return String(value);
}

export function padRow(row) {
  const copy = Array.isArray(row) ? row.slice() : [];
  while (copy.length < COLUMN_LABELS.length) copy.push("");
  return copy;
}

/** Copy of the sheet matrix that is safe to edit: every row padded to the full column set. */
export function draftFromMatrix(matrix) {
  return (Array.isArray(matrix) ? matrix : []).map(padRow);
}

export function blankDraftRow() {
  return padRow([]);
}

/** A vacancy row: surname starts with TBH_, which the org builder treats as "to be hired". */
export function vacancyDraftRow() {
  const row = blankDraftRow();
  row[0] = "TBH_";
  return row;
}

export function isFlagOn(value) {
  return flagOn(value);
}

export function flagValue(on) {
  return on ? "1" : "";
}

/** Drop rows past the last one that holds any text, so the sheet does not grow blank tails. */
export function trimTrailingBlankRows(draft, firstRow = 0) {
  let end = draft.length;
  while (end > firstRow && draft[end - 1].slice(0, COLUMN_LABELS.length).every((cell) => cellText(cell).trim() === "")) {
    end -= 1;
  }
  return draft.slice(0, end);
}

/** Map of 1-based sheet row to messages, from the issues buildOrg reported. */
export function problemRows(org) {
  const map = new Map();
  for (const issue of org?.issues || []) {
    if (!issue.row) continue;
    if (!map.has(issue.row)) map.set(issue.row, []);
    map.get(issue.row).push(issue.message);
  }
  return map;
}

/** First blocking message from buildOrg, or "" when the sheet can be drawn. */
export function sheetProblem(org) {
  if (!org?.fatal) return "";
  return org.issues.find((issue) => issue.level === "error")?.message || "This sheet has errors.";
}

/**
 * Which tab to open after the checks ran. `usable` is the tab to import: `prefer` when it passes,
 * otherwise the first tab that passes. When none passes, `usable` is null and `shown` is the tab to
 * explain (still `prefer` when it exists, so the user's choice is not lost).
 * @param {{ name: string, problem: string }[]} sheets
 */
export function pickSheet(sheets, prefer = "") {
  const list = sheets || [];
  const preferred = list.find((sheet) => sheet.name === prefer);
  const usable = (preferred && !preferred.problem ? preferred : list.find((sheet) => !sheet.problem)) || null;
  return { usable, shown: usable || preferred || list[0] || null };
}

/**
 * Check a tab both ways: first row as data, then first row as titles. A real title row never passes as
 * a person (its manager names are not in the sheet), so "passes as data" means there is no title row.
 * When neither passes, the titled reading is reported, since templates normally have titles.
 */
export function checkSheet(matrix) {
  const asData = buildOrg(matrix, { hasTitles: false });
  if (!asData.fatal) return { org: asData, hasTitles: false, problem: "" };
  const titled = buildOrg(matrix, { hasTitles: true });
  return { org: titled, hasTitles: true, problem: sheetProblem(titled) };
}
