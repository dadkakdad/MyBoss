/**
 * Helpers for the in-browser spreadsheet editor. Pure functions, no DOM.
 * A "draft" is an array of rows, each row an array of at least COLUMN_LABELS.length cells.
 * A copied tab is a new sheet in the same workbook: same rows, its own name, so another tree
 * can be drawn for a different viewer without changing the tab it came from.
 */

import { COLUMN_LABELS, buildOrg, flagOn } from "./org.js";

/** Excel's limit for a sheet name. */
export const SHEET_NAME_MAX = 31;

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

const SHEET_NAME_FORBIDDEN = /[:\\/?*[\]]/;

/** "" when `name` can be a new Excel tab. `existingNames` are the tabs already in the workbook. */
export function sheetNameError(name, existingNames = []) {
  const text = String(name ?? "").trim();
  if (!text) return "Enter a name for the new tab.";
  if (text.length > SHEET_NAME_MAX) return `A tab name can be at most ${SHEET_NAME_MAX} characters.`;
  if (SHEET_NAME_FORBIDDEN.test(text)) return "A tab name cannot contain : \\ / ? * [ ].";
  if (/[\u0000-\u001f]/.test(text)) return "A tab name cannot contain a line break.";
  if (text.startsWith("'") || text.endsWith("'")) return "A tab name cannot start or end with an apostrophe.";
  const taken = (existingNames || []).some((item) => String(item).toLowerCase() === text.toLowerCase());
  if (taken) return "A tab with that name already exists.";
  return "";
}

/** A legal base for "Name (2)". Characters Excel rejects become spaces. */
function sheetNameBase(sourceName) {
  const cleaned = String(sourceName || "")
    .replace(/[:\\/?*[\]]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^'+|'+$/g, "")
    .trim();
  return cleaned || "Sheet";
}

/** Next free "Name (2)", "Name (3)", … that fits in Excel's 31 characters. */
export function nextSheetCopyName(existingNames, sourceName) {
  const names = existingNames || [];
  const base = sheetNameBase(sourceName);
  for (let n = 2; n < 1000; n += 1) {
    const suffix = ` (${n})`;
    let head = base;
    if (head.length + suffix.length > SHEET_NAME_MAX) head = head.slice(0, SHEET_NAME_MAX - suffix.length).trimEnd();
    if (!head) head = "Sheet";
    const candidate = `${head}${suffix}`;
    if (!sheetNameError(candidate, names)) return candidate;
  }
  return "Sheet (2)";
}

function cloneCell(value) {
  return value instanceof Date ? new Date(value.getTime()) : value;
}

/** Independent rows: editing the copy does not change the tab it came from. Dates stay dates. */
export function cloneSheetMatrix(matrix) {
  const draft = draftFromMatrix(matrix).map((row) => row.map(cloneCell));
  return trimTrailingBlankRows(draft, 0);
}

/**
 * The tab to add. `matrix`, when given, is the rows to copy (the editor draft); otherwise the
 * stored rows of `sourceName`. The source list is not modified.
 */
export function copiedSheet({ sheets, sourceName, requestedName, matrix }) {
  const list = Array.isArray(sheets) ? sheets : [];
  const source = list.find((sheet) => sheet.name === sourceName);
  if (!source && matrix == null) return { error: "Choose a tab to copy." };
  const name = String(requestedName ?? "").trim();
  const error = sheetNameError(name, list.map((sheet) => sheet.name));
  if (error) return { error };
  const rows = matrix != null ? matrix : source.matrix;
  return { sheet: { name, matrix: cloneSheetMatrix(rows) } };
}

/** The copy sits just after the tab it came from. The original list is left as it was. */
export function withSheetCopy(sheets, sourceName, sheet) {
  const list = Array.isArray(sheets) ? sheets.slice() : [];
  const index = list.findIndex((item) => item.name === sourceName);
  list.splice(index >= 0 ? index + 1 : list.length, 0, sheet);
  return list;
}
