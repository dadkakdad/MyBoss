import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import {
  formatMoscowUploadTime,
  refreshSavedEntries,
  saveBody,
  workMenuLabel,
  workToReopen,
} from "../js/persist.js";

test("a reload reopens the newest saved workbook for the signed-in account", () => {
  const source = readFileSync(new URL("../js/app.js", import.meta.url), "utf8");
  const boot = source.slice(source.lastIndexOf("async function restoreAccount"));
  assert.match(boot, /restoreLatest\(/);
  const reopen = source.slice(source.indexOf("async function restoreLatest"), source.indexOf("async function submitAccount"));
  assert.match(reopen, /workToReopen\(/);
});

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

test("upload times are shown in Moscow local time", () => {
  assert.equal(formatMoscowUploadTime(null), "");
  assert.equal(formatMoscowUploadTime("not-a-date"), "");
  // 2026-11-13 10:05 UTC → 13:05 Europe/Moscow
  assert.equal(formatMoscowUploadTime("2026-11-13T10:05:00.000Z"), "13.11.2026, 13:05");
});

test("Refresh lists every saved workbook so older uploads stay selectable", () => {
  assert.deepEqual(refreshSavedEntries(null), []);
  assert.deepEqual(refreshSavedEntries([]), []);
  assert.deepEqual(
    refreshSavedEntries(
      [
        { id: "newest", name: "new.xlsx", sheetName: "Sheet1", createdAt: "2026-11-13T10:05:00.000Z" },
        { id: "older", name: "old.xlsx", sheetName: "Advanced", createdAt: "2026-11-12T08:00:00.000Z" },
        { id: "", name: "broken.xlsx" },
      ],
      "newest",
    ),
    [
      {
        id: "newest",
        label: "new.xlsx · Sheet1",
        when: "13.11.2026, 13:05",
        current: true,
      },
      {
        id: "older",
        label: "old.xlsx · Advanced",
        when: "12.11.2026, 11:00",
        current: false,
      },
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
