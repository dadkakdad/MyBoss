import assert from "node:assert/strict";
import test from "node:test";
import { readFileSync } from "node:fs";
import { SLIDE } from "../js/layout.js";
import {
  boxesIntersect,
  contentFrame,
  costCenterText,
  hideCostCenterStamps,
  offsetSlideTargets,
  prependCharts,
  removeSlide,
  scaleShapes,
  selectionBounds,
  showCostCenterStamps,
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
      ["420", 140, 100, 50, 30],
      ["?", 30, 20, 50, 30],
    ],
  );
  hideCostCenterStamps(slides);
  assert.equal(slides[0].shapes.some((shape) => shape.type === "cc"), false);
  assert.equal(slides[0].shapes.some((shape) => shape.type === "card"), true);
});
