import assert from "node:assert/strict";
import test from "node:test";
import { refreshSavedEntries, saveBody, workMenuLabel, workToReopen } from "../js/persist.js";

test("reopening uses the newest saved workbook", () => {
  assert.equal(workToReopen(null), null);
  assert.equal(workToReopen([]), null);
  assert.equal(workToReopen([{ id: "newest" }, { id: "older" }]).id, "newest");
});

test("saved workbook labels show the file and sheet name", () => {
  assert.equal(workMenuLabel(null), "");
  assert.equal(workMenuLabel({ name: "team.xlsx" }), "team.xlsx");
  assert.equal(workMenuLabel({ name: "team.xlsx", sheetName: "Advanced" }), "team.xlsx · Advanced");
});

test("Refresh lists every saved workbook so older uploads stay selectable", () => {
  assert.deepEqual(refreshSavedEntries(null), []);
  assert.deepEqual(refreshSavedEntries([]), []);
  assert.deepEqual(
    refreshSavedEntries(
      [
        { id: "newest", name: "new.xlsx", sheetName: "Sheet1" },
        { id: "older", name: "old.xlsx", sheetName: "Advanced" },
        { id: "", name: "broken.xlsx" },
      ],
      "newest",
    ),
    [
      { id: "newest", label: "new.xlsx · Sheet1", current: true },
      { id: "older", label: "old.xlsx · Advanced", current: false },
    ],
  );
});

test("the first save stores the imported file and the slides together", () => {
  const body = saveBody({
    workId: "",
    fileName: "team.xlsx",
    fileBase64: "QQ==",
    sheetName: "Advanced",
    slides: [{ id: "s", shapes: [{ type: "card", text: "Ada" }] }],
    active: 0,
    includeFile: true,
  });
  assert.equal(body.name, "team.xlsx");
  assert.equal(body.dataBase64, "QQ==");
  assert.equal(body.sheetName, "Advanced");
  assert.equal(body.slides[0].shapes[0].text, "Ada");
  assert.equal(body.active, 0);
});

test("a later save can update the slides without sending the file again", () => {
  const body = saveBody({
    workId: "abc",
    fileName: "team.xlsx",
    fileBase64: "QQ==",
    sheetName: "Advanced",
    slides: [{ id: "s", shapes: [] }, { id: "s2", shapes: [] }],
    active: 1,
    includeFile: false,
  });
  assert.equal("dataBase64" in body, false);
  assert.equal("name" in body, false);
  assert.equal(body.slides.length, 2);
  assert.equal(body.active, 1);
});
