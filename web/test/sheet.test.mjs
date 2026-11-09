import assert from "node:assert/strict";
import test from "node:test";
import { buildOrg } from "../js/org.js";
import {
  blankDraftRow,
  cellText,
  draftFromMatrix,
  flagValue,
  isFlagOn,
  problemRows,
  trimTrailingBlankRows,
  vacancyDraftRow,
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
