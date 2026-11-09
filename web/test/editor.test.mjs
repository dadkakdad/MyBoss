import assert from "node:assert/strict";
import test from "node:test";
import { readFileSync } from "node:fs";
import { SLIDE } from "../js/layout.js";
import {
  anchorNotes,
  boxesIntersect,
  cardNoteRecords,
  contentFrame,
  costCenterText,
  hideCostCenterStamps,
  noteBox,
  noteHeightFor,
  noteTypeSize,
  NOTE_COLORS,
  offsetSlideTargets,
  placeCardNotes,
  prependCharts,
  removeSlide,
  scaleShapes,
  selectionBounds,
  showCostCenterStamps,
  stampSize,
  stickNotes,
} from "../js/editor.js";

function slide(shapes, title = "") {
  return { title, shapes };
}

test("new charts are prepended and old slide numbers move back", () => {
  const existing = [
    slide([{ type: "card", x: 0, y: 0, w: 10, h: 10 }, { type: "link", text: "See slide 2", target: 2 }], "Kept"),
    slide([{ type: "nav", text: "Up to Ada (slide 1)", target: 1 }]),
  ];
  const fresh = [slide([{ type: "title", x: 0, y: 0, w: 10, h: 10 }]), slide([])];
  const deck = prependCharts(existing, fresh);
  assert.equal(deck.length, 4);
  assert.equal(deck[0], fresh[0]);
  assert.equal(deck[2].title, "Kept");
  assert.equal(deck[2].shapes[1].target, 4);
  assert.equal(deck[2].shapes[1].text, "See slide 4");
  assert.equal(deck[3].shapes[0].target, 3);
  assert.equal(deck[3].shapes[0].text, "Up to Ada (slide 3)");
});

test("the untouched blank canvas is replaced by the first chart", () => {
  const blank = slide([]);
  const fresh = [slide([{ type: "card", x: 1, y: 1, w: 2, h: 2 }])];
  const deck = prependCharts([blank], fresh);
  assert.deepEqual(deck, fresh);
});

test("a named empty slide is kept when a chart is added in front", () => {
  const named = slide([], "Notes");
  const fresh = [slide([{ type: "card", x: 0, y: 0, w: 1, h: 1 }])];
  const deck = prependCharts([named], fresh);
  assert.equal(deck.length, 2);
  assert.equal(deck[1].title, "Notes");
});

test("deleting a slide drops links to it and renumbers the rest", () => {
  const deck = [
    slide([{ type: "link", text: "See slide 2", target: 2 }, { type: "card", linkSlide: 3 }]),
    slide([{ type: "card", x: 0, y: 0, w: 1, h: 1 }]),
    slide([{ type: "nav", text: "Up to Grace (slide 2)", target: 2 }]),
  ];
  const next = removeSlide(deck, 1);
  assert.equal(next.length, 2);
  assert.equal(next[0].shapes.length, 1);
  assert.equal(next[0].shapes[0].type, "card");
  assert.equal(next[0].shapes[0].linkSlide, 2);
  assert.equal(next[1].shapes.length, 0);
});

test("inserting a slide shifts only the numbers after the insert point", () => {
  const deck = [
    slide([{ type: "link", text: "See slide 1", target: 1 }]),
    slide([{ type: "link", text: "See slide 2", target: 2 }]),
  ];
  deck.splice(1, 0, slide([]));
  offsetSlideTargets(deck, 1, 1);
  assert.equal(deck[0].shapes[0].target, 1);
  assert.equal(deck[2].shapes[0].target, 3);
  assert.equal(deck[2].shapes[0].text, "See slide 3");
});

test("content can hang off the slide page", () => {
  const frame = contentFrame([
    { type: "card", x: -30, y: 500, w: 80, h: 80 },
    { type: "wire", x1: 10, y1: 10, x2: 1000, y2: 20 },
  ]);
  assert.ok(frame.x < 0);
  assert.ok(frame.y === 0);
  assert.ok(frame.w > SLIDE.width);
  assert.ok(frame.h > SLIDE.height);
  assert.ok(frame.y + frame.h > 580);
});

test("marquee hits a shape that only overlaps the rectangle", () => {
  const shape = { x: 40, y: 40, w: 20, h: 20 };
  const box = { x: shape.x, y: shape.y, w: shape.w, h: shape.h };
  assert.equal(boxesIntersect(box, { x: 50, y: 50, w: 30, h: 30 }), true);
  assert.equal(boxesIntersect(box, { x: 80, y: 80, w: 10, h: 10 }), false);
});

test("scale grows a selection away from the anchor", () => {
  const shapes = [
    { type: "card", x: 10, y: 20, w: 40, h: 10 },
    { type: "wire", x1: 10, y1: 20, x2: 50, y2: 20, weight: 2 },
  ];
  const applied = scaleShapes(shapes, 2, { x: 10, y: 20 });
  assert.equal(applied, 2);
  assert.equal(shapes[0].x, 10);
  assert.equal(shapes[0].y, 20);
  assert.equal(shapes[0].w, 80);
  assert.equal(shapes[0].h, 20);
  assert.equal(shapes[1].x2, 90);
  assert.equal(shapes[1].weight, 4);
  const bounds = selectionBounds(shapes);
  assert.equal(bounds.w, 80);
});

test("the ribbon exposes slide tools and cost centers", () => {
  const html = readFileSync(new URL("../index.html", import.meta.url), "utf8");
  assert.match(html, /id="cost-centers"/);
  assert.match(html, /id="tool-select"/);
  assert.match(html, /id="tool-scale"/);
  assert.match(html, /id="filmstrip"/);
  const ribbon = html.slice(html.indexOf('<header class="ribbon">'), html.indexOf("</header>"));
  assert.match(ribbon, /id="account"[^>]*>Log in</);
  const status = html.slice(html.indexOf('<div class="status">'), html.indexOf('<div class="sheet-dialog"'));
  assert.equal(status.includes('id="account"'), false);
  assert.match(html, /class="empty-slide"/);
  const css = readFileSync(new URL("../css/app.css", import.meta.url), "utf8");
  assert.match(css, /\.holder[\s\S]*aspect-ratio:\s*960\s*\/\s*540/);
});

test("cost-center stamps sit on the card corner and toggle off", () => {
  assert.equal(costCenterText(""), "?");
  assert.equal(costCenterText("0"), "?");
  assert.equal(costCenterText(" 420 "), "420");
  const people = new Map([["ada", { costCenter: "420" }], ["blank", { costCenter: "" }]]);
  const slides = [
    slide([
      { type: "card", key: "ada", x: 100, y: 80, w: 90, h: 40 },
      { type: "card", key: "blank", x: 10, y: 10, w: 70, h: 30 },
      { type: "title", x: 0, y: 0, w: 10, h: 10 },
    ]),
  ];
  let n = 0;
  const count = showCostCenterStamps({ people }, slides, () => `cc-${(n += 1)}`);
  assert.equal(count, 2);
  const stamps = slides[0].shapes.filter((shape) => shape.type === "cc");
  assert.deepEqual(
    stamps.map((shape) => [shape.text, shape.x, shape.y, shape.w, shape.h]),
    [
      ["420", 164, 109, 26, 11],
      ["?", 54, 29, 26, 11],
    ],
  );
  const long = stampSize("СТАЖЕР", 70);
  assert.ok(long.w > 26 && long.w <= 70, "grows with the text");
  assert.equal(stampSize("0", 70).w, 26, "short text keeps the usual size");
  const tight = stampSize("ПРОЕКТНЫЙ ОФИС", 50);
  assert.equal(tight.w, 50);
  assert.ok(tight.fs < 7.04 && tight.fs >= 4.5, "type shrinks when the card is too narrow");
  hideCostCenterStamps(slides);
  assert.equal(slides[0].shapes.some((shape) => shape.type === "cc"), false);
  assert.equal(slides[0].shapes.some((shape) => shape.type === "card"), true);
});

test("slide text is copied in reading order, one line per shape", async () => {
  const { shapesText } = await import("../js/editor.js");
  const people = new Map([
    ["ada", { surname: "Ada [x]", name: "Lovelace", job: "CEO", location: "London" }],
    ["tbd", { surname: "TBH_a", name: "TBH_a", job: "Analyst", tbh: true }],
  ]);
  const text = shapesText({ people }, [
    { type: "card", key: "tbd", x: 200, y: 100, w: 80, h: 40 },
    { type: "card", key: "ada", x: 10, y: 100, w: 80, h: 40 },
    { type: "title", name: "Ada Lovelace", department: "Office", x: 0, y: 0, w: 10, h: 10 },
    { type: "wire", x1: 0, y1: 0, x2: 5, y2: 5 },
    { type: "comment", text: "Works until 1 May", x: 10, y: 150, w: 50, h: 10 },
  ]);
  assert.equal(text, "Ada Lovelace\nOffice\nAda Lovelace — CEO — London\nTo be hired — Analyst\nWorks until 1 May");
});

test("a slide copies as Markdown: title, statistics, nested structure, data table", async () => {
  const { slideMarkdown } = await import("../js/editor.js");
  const person = (key, surname, name, job, extra = {}) => ({ key, surname, name, job, department: "Lab", location: "", solidKey: "", dottedKey: "", solidName: "", dottedName: "", stats: { subordinates: 1 }, ...extra });
  const boss = person("ada", "Ada", "Lovelace", "CEO", { stats: { subordinates: 3 } });
  const grace = person("grace", "Grace", "Hopper", "Eng", { solidKey: "ada", solidName: "Ada Lovelace", location: "London", newcomer: true, newcomerDate: "16.11" });
  const ken = person("ken", "Ken", "Thompson", "Eng | Ops", { solidKey: "ada", solidName: "Ada Lovelace", dottedKey: "grace", dottedName: "Grace Hopper", termination: "30.12.2026" });
  const shapes = [
    { type: "title", name: "Ada Lovelace", department: "Office", x: 0, y: 0, w: 10, h: 10 },
    { type: "stats", heading: "Department statistics", lines: ["Ada has 2 direct reports", "Team Size – 3 people"], x: 0, y: 0, w: 10, h: 10 },
    { type: "card", key: "ada", person: boss, x: 100, y: 100, w: 80, h: 40, head: true, managerType: "Solid" },
    { type: "card", key: "grace", person: grace, x: 50, y: 200, w: 80, h: 40, managerType: "Solid" },
    { type: "note", key: "grace", text: "call her", x: 64, y: 230, w: 110, h: 28 },
    { type: "card", key: "ken", person: ken, x: 150, y: 200, w: 80, h: 40, managerType: "Solid" },
    { type: "card", key: "ken", person: ken, x: 60, y: 300, w: 80, h: 40, managerType: "Dotted" },
    { type: "hiddenlist", text: "Hidden – 1 person: Zed Last", x: 0, y: 500, w: 100, h: 10 },
  ];
  const md = slideMarkdown({ people: new Map() }, shapes, { number: 1, total: 2 });
  assert.match(md, /^# Ada Lovelace\n_Office_\n\nSlide 1 of 2/);
  assert.match(md, /\*\*Department statistics\*\*\n\n- Ada has 2 direct reports\n- Team Size – 3 people/);
  assert.match(md, /## Structure\n\n- \*\*Ada Lovelace\*\* — CEO _\(team of 3\)_\n  - \*\*Grace Hopper\*\* — Eng _\(London; newcomer, joins 16.11; call her\)_\n    - \*\*Ken Thompson\*\* — Eng \| Ops _\(dotted line; works until 30.12.2026\)_\n  - \*\*Ken Thompson\*\*/);
  assert.match(md, /\| Grace Hopper \| Eng \| Lab \| London \| Ada Lovelace \| 1 \| newcomer, joins 16.11; call her \|/);
  assert.match(md, /\| Name \| Job \| Department \| Location \| Reports to \| Team \| Notes \|/);
  assert.match(md, /\| Ken Thompson \| Eng \\\| Ops \| Lab \| — \| Grace Hopper \| 1 \| dotted line; works until 30.12.2026 \|/);
  assert.match(md, /> Hidden – 1 person: Zed Last\n$/);
  assert.equal(slideMarkdown({ people: new Map() }, [{ type: "comment", text: "Works until 1 May", x: 0, y: 0, w: 10, h: 10 }]), "Works until 1 May");
});

test("a note is an orange rectangle on the card, and a second one stacks under it", () => {
  const card = { id: "c1", type: "card", key: "ada", size: "Medium", x: 100, y: 40, w: 80, h: 42.3 };
  const first = noteBox(card, 0);
  const second = noteBox(card, 1);
  assert.equal(NOTE_COLORS[0], "#E8891A");
  assert.equal(noteTypeSize(card), 9);
  assert.equal(noteTypeSize({ size: "Small", h: 35.5 }), 7);
  assert.equal(noteTypeSize({ size: "Large", h: 54.5 }), 11);
  assert.equal(first.fs, 9);
  assert.ok(first.h < 20, "the rectangle is one line of the card name");
  assert.ok(first.w >= card.w);
  assert.ok(first.y < card.y + card.h, "overlaps the card");
  assert.ok(first.y + first.h > card.y + card.h, "hangs below the card");
  assert.ok(first.x + first.w > card.x + card.w, "hangs to the right");
  assert.equal(second.x, first.x);
  assert.ok(second.y >= first.y + first.h);
  assert.ok(noteHeightFor("one\ntwo\nthree", first.w, first.fs) > noteHeightFor("one", first.w, first.fs));
});

test("notes follow their card, and a dragged note keeps its new offset", () => {
  const card = { id: "c1", type: "card", key: "ada", x: 10, y: 20, w: 80, h: 40 };
  const box = noteBox(card, 0);
  const note = { id: "n1", type: "note", key: "ada", cardId: "c1", text: "hello", ...box };
  const shapes = [card, note];
  card.x = 50;
  card.y = 80;
  stickNotes(shapes);
  assert.equal(note.x, 50 + box.dx);
  assert.equal(note.y, 80 + box.dy);

  note.x += 12;
  note.y += 3;
  anchorNotes(shapes, new Set(["n1"]));
  assert.equal(note.dx, box.dx + 12);
  assert.equal(note.dy, box.dy + 3);
  card.x += 5;
  anchorNotes(shapes, new Set(["c1"]));
  assert.equal(note.x, card.x + note.dx);
  assert.equal(note.dx, box.dx + 12);
});

test("generating a chart stamps the note from the slide you were editing", () => {
  const older = {
    shapes: [
      { id: "c-old", type: "card", key: "ada", x: 0, y: 0, w: 80, h: 40 },
      { id: "n1", type: "note", key: "ada", text: "old", fill: "#E8891A", dx: 14, dy: 30, w: 110, h: 28, x: 14, y: 30 },
    ],
  };
  const current = {
    shapes: [
      { id: "c-now", type: "card", key: "ada", x: 0, y: 0, w: 80, h: 40 },
      { id: "n1", type: "note", key: "ada", text: "Василиса Премудрая?", fill: "#F6C445", dx: 20, dy: 36, w: 120, h: 32, x: 20, y: 36 },
    ],
  };
  const records = cardNoteRecords([older, current], 1);
  const fresh = [{ id: "c-new", type: "card", key: "ada", x: 200, y: 100, w: 90, h: 42 }];
  placeCardNotes(fresh, records);
  placeCardNotes(fresh, records);
  const notes = fresh.filter((shape) => shape.type === "note");
  assert.equal(notes.length, 1);
  assert.equal(notes[0].id, "n1");
  assert.equal(notes[0].text, "Василиса Премудрая?");
  assert.equal(notes[0].fill, "#F6C445");
  assert.equal(notes[0].x, 200 + 20);
  assert.equal(notes[0].y, 100 + 36);
  assert.equal(notes[0].cardId, "c-new");
});
