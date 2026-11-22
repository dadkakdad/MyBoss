import assert from "node:assert/strict";
import test from "node:test";
import { buildOrg } from "../js/org.js";
import {
  blankDraftRow,
  cellText,
  copiedSheet,
  draftFromMatrix,
  flagValue,
  isFlagOn,
  nextSheetCopyName,
  problemRows,
  renamedSheet,
  sheetNameError,
  slideSheetNameAfterRename,
  trimTrailingBlankRows,
  vacancyDraftRow,
  withSheetCopy,
} from "../js/sheet.js";

const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];

test("draft rows are padded copies, not the original rows", () => {
  const matrix = [["Ada", "Lovelace"]];
  const draft = draftFromMatrix(matrix);
  assert.equal(draft[0].length, 17);
  draft[0][0] = "Changed";
  assert.equal(matrix[0][0], "Ada");
});

test("a vacancy row starts with TBH_ and is read as to-be-hired once named", () => {
  const vacancy = vacancyDraftRow();
  assert.equal(vacancy[0], "TBH_");
  vacancy[0] = "TBH_SE";
  vacancy[1] = "One";
  vacancy[4] = "Ada";
  vacancy[5] = "Lovelace";
  const org = buildOrg([header, ["Ada", "Lovelace", "CEO", "Office", ...Array(9).fill("")], vacancy]);
  assert.equal(org.fatal, false);
  assert.equal([...org.people.values()].filter((person) => person.tbh).length, 1);
});

test("flags round-trip through the editor", () => {
  assert.equal(isFlagOn(flagValue(true)), true);
  assert.equal(isFlagOn(flagValue(false)), false);
  assert.equal(isFlagOn(1), true);
  assert.equal(isFlagOn(""), false);
});

test("dates show as ISO text and blanks as empty", () => {
  assert.equal(cellText(new Date(Date.UTC(2025, 2, 1))), "2025-03-01");
  assert.equal(cellText(null), "");
  assert.equal(cellText(420), "420");
});

test("trailing blank rows are trimmed but the title row stays", () => {
  const draft = [header.slice(), blankDraftRow(), ["Ada", "L", ...Array(11).fill("")], blankDraftRow(), blankDraftRow()];
  assert.equal(trimTrailingBlankRows(draft, 1).length, 3);
  assert.equal(trimTrailingBlankRows([blankDraftRow()], 1).length, 1);
});

test("problem rows map the issue to the sheet row", () => {
  const rows = [header, ["Ada", "Lovelace", "CEO", "Office", ...Array(9).fill("")], ["Bob", "Ross", "Eng", "Lab", "Zed", "Nobody", ...Array(7).fill("")]];
  const org = buildOrg(rows);
  assert.equal(org.fatal, true);
  assert.equal(problemRows(org).has(3), true);
});

test("the first tab that passes is opened, the preferred one when it still passes", async () => {
  const { pickSheet, sheetProblem } = await import("../js/sheet.js");
  const sheets = [
    { name: "Raw", problem: "Row 7 is missing a surname or a name." },
    { name: "Org", problem: "" },
    { name: "Old", problem: "" },
  ];
  assert.equal(pickSheet(sheets).usable.name, "Org");
  assert.equal(pickSheet(sheets, "Old").usable.name, "Old");
  assert.equal(pickSheet(sheets, "Raw").usable.name, "Org", "a broken preference falls through to a usable tab");
  const broken = sheets.map((sheet) => ({ ...sheet, problem: "bad" }));
  assert.equal(pickSheet(broken).usable, null);
  assert.equal(pickSheet(broken).shown.name, "Raw");
  assert.equal(pickSheet(broken, "Old").shown.name, "Old", "the tab the user picked stays shown");
  assert.deepEqual(pickSheet([]), { usable: null, shown: null });
  assert.equal(sheetProblem({ fatal: false, issues: [] }), "");
  assert.equal(sheetProblem({ fatal: true, issues: [{ level: "error", message: "No CEO found." }] }), "No CEO found.");
  assert.equal(sheetProblem(null), "");
});

test("a tab is read with or without a title row, whichever passes", async () => {
  const { checkSheet } = await import("../js/sheet.js");
  const ada = ["Ada", "Lovelace", "CEO", "Office", "", "", "", "", "", "", "", "", ""];
  const grace = ["Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace", "", "", "", "", "", "", ""];
  const titles = ["Employee Surname", "Employee Name", "Job", "Team", "Reports To Surname", "Reports To Name", "", "", "", "", "", "", ""];
  const titled = checkSheet([titles, ada, grace]);
  assert.equal(titled.hasTitles, true);
  assert.equal(titled.problem, "");
  assert.equal(titled.org.companySize, 2);
  const bare = checkSheet([ada, grace]);
  assert.equal(bare.hasTitles, false);
  assert.equal(bare.org.companySize, 2, "the first row is a person, not dropped as a title");
  const broken = checkSheet([titles, ["", "Lovelace", "CEO", "Office", "", "", "", "", "", "", "", "", ""]]);
  assert.equal(broken.hasTitles, true);
  assert.match(broken.problem, /missing a surname/);
});

test("a new tab name has to be unique and legal for Excel", () => {
  assert.equal(sheetNameError("", ["Org"]), "Enter a name for the new tab.");
  assert.equal(sheetNameError("   ", []), "Enter a name for the new tab.");
  assert.equal(sheetNameError("Board", ["Org"]), "");
  assert.match(sheetNameError("org", ["Org"]), /already exists/);
  assert.match(sheetNameError(`${"A".repeat(32)}`, []), /31 characters/);
  assert.equal(sheetNameError("A".repeat(31), []), "");
  assert.match(sheetNameError("Team/Board", []), /cannot contain/);
  assert.match(sheetNameError("Team?Board", []), /cannot contain/);
  assert.match(sheetNameError("Board\n2", []), /line break/);
  assert.match(sheetNameError("'Board", []), /apostrophe/);
  assert.match(sheetNameError("Board'", []), /apostrophe/);
  assert.equal(sheetNameError("O'Brien", []), "");
});

test("renaming a tab keeps its rows and follows Excel's name rules", () => {
  const sheets = [
    { name: "Advanced", matrix: [header, ["Ada", "Lovelace"]] },
    { name: "Board", matrix: [] },
  ];
  const same = renamedSheet({ sheets, sourceName: "Advanced", requestedName: "  Advanced  " });
  assert.equal(same.unchanged, true);
  assert.equal(same.name, "Advanced");
  assert.equal(sheets[0].name, "Advanced", "deciding the name does not rewrite the workbook");
  assert.equal(sheets[0].matrix[1][0], "Ada");

  const renamed = renamedSheet({ sheets, sourceName: "Advanced", requestedName: " Leadership " });
  assert.equal(renamed.error, undefined);
  assert.equal(renamed.from, "Advanced");
  assert.equal(renamed.name, "Leadership");
  assert.equal(sheets[0].name, "Advanced");

  const cased = renamedSheet({ sheets, sourceName: "Advanced", requestedName: "advanced" });
  assert.equal(cased.error, undefined);
  assert.equal(cased.name, "advanced");

  assert.match(renamedSheet({ sheets, sourceName: "Advanced", requestedName: "board" }).error, /already exists/);
  assert.equal(renamedSheet({ sheets, sourceName: "Advanced", requestedName: "  " }).error, "Enter a name for the tab.");
  assert.equal(renamedSheet({ sheets, sourceName: "Missing", requestedName: "Team" }).error, "Choose a tab to rename.");
  assert.match(renamedSheet({ sheets, sourceName: "Advanced", requestedName: "Team/Board" }).error, /cannot contain/);
  assert.match(renamedSheet({ sheets, sourceName: "Advanced", requestedName: "A".repeat(32) }).error, /31 characters/);

  assert.equal(slideSheetNameAfterRename("Advanced", "Advanced", "Leadership", false), "Leadership");
  assert.equal(slideSheetNameAfterRename("Board", "Advanced", "Leadership", true), "Board");
  assert.equal(slideSheetNameAfterRename("", "Advanced", "Leadership", true), "Leadership");
  assert.equal(slideSheetNameAfterRename(undefined, "Advanced", "Leadership", false), undefined);
});

test("the suggested copy name is Name (2), then the next free number, and it always fits", () => {
  assert.equal(nextSheetCopyName(["Advanced"], "Advanced"), "Advanced (2)");
  assert.equal(nextSheetCopyName(["Advanced", "Advanced (2)"], "Advanced"), "Advanced (3)");
  assert.equal(nextSheetCopyName(["Advanced", "advanced (2)"], "Advanced"), "Advanced (3)");
  assert.equal(nextSheetCopyName(["Other"], "Team/Board"), "Team Board (2)");
  assert.equal(nextSheetCopyName(["Team Board (2)"], "Team/Board"), "Team Board (3)");
  assert.equal(nextSheetCopyName([], "???"), "Sheet (2)");
  const long = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  const suggested = nextSheetCopyName([long], long);
  assert.equal(suggested.length <= 31, true);
  assert.equal(sheetNameError(suggested, [long]), "");
  assert.match(suggested, /\(2\)$/);
});

test("copying a tab keeps the same people on independent rows", () => {
  const hired = new Date(Date.UTC(2026, 10, 16));
  const ada = ["Ada", "Lovelace", "CEO", "Office", "", "", "", "", "Moscow", hired];
  const grace = ["Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace"];
  const sheets = [{ name: "Advanced", matrix: [header, ada, grace, blankDraftRow()] }];
  const draft = [header, ada, grace, ["Bob", "Ross", "Art", "Studio", "Ada", "Lovelace"]];
  const copy = copiedSheet({ sheets, sourceName: "Advanced", requestedName: "Board", matrix: draft });
  assert.equal(copy.error, undefined);
  assert.equal(copy.sheet.name, "Board");
  assert.equal(copy.sheet.matrix.length, 4, "the editor draft is copied, including the extra person");
  assert.equal(sheets[0].matrix.length, 4, "the source list still has its own rows");
  assert.notEqual(copy.sheet.matrix[1], ada);
  assert.notEqual(copy.sheet.matrix[1][9], hired);
  assert.equal(copy.sheet.matrix[1][9].getTime(), hired.getTime());
  const org = buildOrg(copy.sheet.matrix);
  assert.equal(org.fatal, false);
  assert.equal(org.companySize, 3);
  copy.sheet.matrix[1][0] = "Changed";
  assert.equal(ada[0], "Ada");
  assert.equal(sheets[0].matrix[1][0], "Ada");

  const stored = copiedSheet({ sheets, sourceName: "Advanced", requestedName: "Team" });
  assert.equal(stored.sheet.matrix.length, 3, "blank rows at the end of the stored tab are dropped");
  assert.equal(buildOrg(stored.sheet.matrix).companySize, 2);

  assert.equal(copiedSheet({ sheets, sourceName: "Missing", requestedName: "Board" }).error, "Choose a tab to copy.");
  assert.match(copiedSheet({ sheets, sourceName: "Advanced", requestedName: "advanced" }).error, /already exists/);
  assert.match(copiedSheet({ sheets, sourceName: "Advanced", requestedName: "  " }).error, /Enter a name/);

  const ordered = withSheetCopy(sheets, "Advanced", copy.sheet);
  assert.deepEqual(ordered.map((item) => item.name), ["Advanced", "Board"]);
  assert.deepEqual(sheets.map((item) => item.name), ["Advanced"]);
  assert.deepEqual(withSheetCopy([{ name: "A" }, { name: "B" }], "A", { name: "A (2)" }).map((item) => item.name), ["A", "A (2)", "B"]);
  assert.deepEqual(withSheetCopy([], "Missing", { name: "New" }).map((item) => item.name), ["New"]);
});

test("a copied tab name can be stored in an xlsx workbook", async () => {
  const XLSX = await import("xlsx");
  const name = nextSheetCopyName(["Advanced"], "Advanced");
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([header, ["Ada", "Lovelace", "CEO", "Office"]]), "Advanced");
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([header, ["Ada", "Lovelace", "CEO", "Office"]]), name);
  const out = XLSX.write(book, { bookType: "xlsx", type: "array" });
  const again = XLSX.read(out, { type: "array" });
  assert.deepEqual(again.SheetNames, ["Advanced", name]);
});
