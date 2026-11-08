import assert from "node:assert/strict";
import test from "node:test";
import { readFileSync } from "node:fs";
import XLSX from "xlsx";
import {
  buildOrg,
  cardStyle,
  directReports,
  displayName,
  headcountLabel,
  isManager,
  peopleInCategory,
  solidTree,
  statisticsLines,
} from "../js/org.js";
import { CARDS, GUIDE, SLIDE, buildSlides, chooseFit, layoutLeaves, makeMetrics } from "../js/layout.js";

const samplePath = new URL("../../MyBoss-Sample data.xlsx", import.meta.url);

function matrixFrom(sheetName) {
  const workbook = XLSX.read(readFileSync(samplePath), { type: "buffer", cellDates: true });
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
}

function row(surname, name, job, dept, ms, mn, ds = "", dn = "", location = "", end = "", out = "", important = "") {
  return [surname, name, job, dept, ms, mn, ds, dn, location, end, out, important, ""];
}

test("sample workbook imports on every sheet", () => {
  for (const name of ["Simplest", "Real-life", "Advanced"]) {
    const org = buildOrg(matrixFrom(name));
    assert.equal(org.fatal, false, `${name}: ${org.issues.map((issue) => issue.message).join(" | ")}`);
    assert.ok(org.ceo, name);
    assert.ok(org.companySize > 5, name);
    const ceo = org.people.get(org.ceo);
    assert.ok(ceo.stats.subordinates >= org.companySize - 2, `${name} headcount ${ceo.stats.subordinates}`);
    assert.ok(headcountLabel(ceo).length > 0);
  }
});

test("CEO, managers, and letter filters follow the ribbon", () => {
  const org = buildOrg(matrixFrom("Advanced"));
  assert.equal(peopleInCategory(org, "CEO").length, 1);
  const managers = peopleInCategory(org, "Managers");
  assert.ok(managers.length > 1);
  assert.ok(managers.every((key) => isManager(org, key)));
  const letter = org.people.get(org.ceo).surname.slice(0, 1).toUpperCase();
  assert.ok(peopleInCategory(org, letter).includes(org.ceo));
});

test("a sheet with fewer than 13 columns is rejected", () => {
  const org = buildOrg([
    ["Employee surname", "Employee name"],
    ["Ada", "Lovelace"],
  ]);
  assert.equal(org.fatal, true);
  assert.match(org.issues[0].message, /enough columns/);
});

test("missing manager, two CEOs, and a loop are rejected", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const missing = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", "Eng", "Lab", "Nobody", "Home"),
  ]);
  assert.equal(missing.fatal, true);

  const two = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", "CEO", "Lab", "", ""),
  ]);
  assert.equal(two.fatal, true);

  const loop = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", "Eng", "Lab", "Alan", "Turing"),
    row("Alan", "Turing", "Eng", "Lab", "Grace", "Hopper"),
  ]);
  assert.equal(loop.fatal, true);
});

test("duplicate names that manage someone are rejected", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace"),
    row("Grace", "Hopper", "Eng", "Field", "Ada", "Lovelace"),
    row("Ken", "Thompson", "Eng", "Lab", "Grace", "Hopper"),
  ]);
  assert.equal(org.fatal, true);
  assert.match(org.issues.map((issue) => issue.message).join(" "), /unique/);
});

test("headcount skips people still to be hired and counts external staff", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", "", "", "", "Russia, Moscow"),
    row("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace", "", "", "Russia, Moscow", "", "", "1"),
    row("TBH_SE", "One", "Eng", "Lab", "Grace", "Hopper"),
    row("Ken", "Thompson", "Eng", "Lab", "Grace", "Hopper", "", "", "Berlin", "", "1"),
  ]);
  assert.equal(org.fatal, false, org.issues.map((i) => i.message).join(" | "));
  const grace = [...org.people.values()].find((person) => person.surname === "Grace");
  assert.equal(grace.stats.subordinates, 2);
  assert.equal(grace.stats.outsource, 1);
  assert.equal(grace.stats.cards, 3);
  assert.equal(displayName([...org.people.values()].find((person) => person.tbh)), "One");
  assert.equal(headcountLabel(grace), "2 (1)");
});

test("a person who is both important and external is counted twice and drawn as a split card", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace", "", "", "Berlin", "", "1", "1"),
  ]);
  assert.equal(org.fatal, false, org.issues.map((issue) => issue.message).join(" | "));
  const grace = [...org.people.values()].find((person) => person.surname === "Grace");
  const ada = org.people.get(org.ceo);
  assert.equal(cardStyle(grace), "split");
  assert.equal(grace.stats.important, 1);
  assert.equal(grace.stats.outsource, 1);
  assert.equal(grace.stats.ordinary, 0);
  assert.equal(ada.stats.important, 1);
  assert.equal(ada.stats.outsource, 1);
  assert.equal(ada.stats.ordinary, 1);
});

test("an empty department is drawn as cards and not as a shade", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", "Eng", "", "Ada", "Lovelace"),
  ]);
  const slides = buildSlides(org, org.ceo, "infinity");
  const shades = slides.flatMap((slide) => slide.shapes.filter((shape) => shape.type === "shade").map((shape) => shape.text));
  assert.equal(shades.includes(""), false);
  assert.equal(shades.includes("<no_department_specified>"), false);
  const grace = [...org.people.values()].find((person) => person.surname === "Grace");
  assert.ok(slides.some((slide) => slide.shapes.some((shape) => shape.type === "card" && shape.key === grace.key)));
});

test("a full chart places every person, including a split card and a blank department", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace"),
    row("Ken", "Thompson", "Eng", "", "Grace", "Hopper", "Ada", "Lovelace", "", "", "1", "1"),
    row("Lin", "Worker", "Eng", "Lab", "Ken", "Thompson"),
    row("TBH_SE", "One", "Eng", "Lab", "Grace", "Hopper"),
  ]);
  assert.equal(org.fatal, false, org.issues.map((issue) => issue.message).join(" | "));
  const slides = buildSlides(org, org.ceo, "infinity");
  const seen = new Set();
  for (const slide of slides) {
    for (const shape of slide.shapes) if (shape.type === "card" && shape.key) seen.add(shape.key);
  }
  for (const key of solidTree(org, org.ceo)) assert.ok(seen.has(key), key);
  assert.equal(seen.size, org.people.size);
});

test("a dotted chain that never reaches the CEO stops the import", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace"),
    row("Bob", "Isolated", "Eng", "Lab", "Bob", "Isolated"),
    row("Ken", "Thompson", "Eng", "Lab", "Grace", "Hopper", "Bob", "Isolated"),
  ]);
  assert.equal(org.fatal, true);
  assert.equal(org.issues.length, 1);
});

test("letter lists are not capped at 1000", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const rows = [header, row("Ada", "Lovelace", "CEO", "Office", "", "")];
  for (let i = 0; i < 3; i += 1) rows.push(row(`A${i}`, "Person", "Eng", "Lab", "Ada", "Lovelace"));
  const org = buildOrg(rows);
  assert.equal(org.fatal, false, org.issues.map((issue) => issue.message).join(" | "));
  assert.equal(peopleInCategory(org, "AZ").length, org.order.length);
});

test("the generated head is Large and drops 15 points", () => {
  const org = wideOrg(1, 3);
  const slides = buildSlides(org, org.ceo, "clever");
  const head = slides[0].shapes.find((shape) => shape.type === "card" && shape.head);
  assert.equal(head.w, CARDS.Large.w);
  const drop = slides[0].shapes.find(
    (shape) =>
      shape.type === "wire" &&
      Math.abs(shape.x1 - shape.x2) < 1 &&
      Math.abs(Math.min(shape.y1, shape.y2) - (head.y + head.h)) < 0.6 &&
      Math.abs(Math.abs(shape.y2 - shape.y1) - 15) < 0.6,
  );
  assert.ok(drop, "15 point head tick");
});

test("dotted-line reports hang off both managers but only expand under the solid line", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace"),
    row("Ken", "Thompson", "Eng", "Lab", "Grace", "Hopper", "Ada", "Lovelace"),
  ]);
  const ada = org.ceo;
  const grace = directReports(org, ada).solid[0];
  assert.equal(directReports(org, ada).dotted.length, 1);
  assert.equal(directReports(org, grace).solid.length, 1);
  assert.equal(isManager(org, ada), true);
});

function wideOrg(teams, peoplePerTeam) {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const rows = [header, row("Ada", "Lovelace", "CEO", "Office", "", "")];
  for (let team = 0; team < teams; team += 1) {
    const lead = `Lead${team}`;
    rows.push(row(lead, "Manager", "Director", `Dept ${team % 3}`, "Ada", "Lovelace"));
    for (let person = 0; person < peoplePerTeam; person += 1) {
      rows.push(row(`P${team}_${person}`, "Member", "IC", `Dept ${team % 3}`, lead, "Manager"));
    }
  }
  return buildOrg(rows);
}

test("a wide organisation folds onto one slide inside the PowerPoint guide", () => {
  const org = wideOrg(8, 6);
  assert.equal(org.fatal, false);
  const fit = chooseFit(org, org.ceo);
  assert.ok(fit.box.w <= GUIDE.width + 1, `width ${fit.box.w}`);
  assert.ok(fit.box.h <= GUIDE.height + 1, `height ${fit.box.h}`);
  const slides = buildSlides(org, org.ceo, "clever");
  assert.equal(slides.length, 1);
  const cards = slides[0].shapes.filter((shape) => shape.type === "card");
  assert.ok(cards.some((card) => card.key === org.ceo && card.head));
  assert.ok(cards.length > 1);
});

test("full structure gives a folded team its own slide and links back", () => {
  const org = wideOrg(10, 8);
  const slides = buildSlides(org, org.ceo, "infinity");
  assert.ok(slides.length > 1, `slides ${slides.length}`);
  const seen = new Set();
  for (const slide of slides) {
    for (const shape of slide.shapes) {
      if (shape.type === "card") seen.add(shape.key);
    }
  }
  for (const key of org.people.keys()) assert.ok(seen.has(key), key);
  const linked = slides[0].shapes.some((shape) => shape.type === "link");
  assert.equal(linked, true);
});

test("a company of 100 or more starts collapsed and still lays out", () => {
  const org = wideOrg(10, 10);
  assert.ok(org.people.get(org.ceo).stats.subordinates >= 100);
  const fit = chooseFit(org, org.ceo);
  const slides = buildSlides(org, org.ceo, "infinity");
  assert.ok(slides.length >= 1);
  assert.ok(fit.box.w < 5000);
});

test("every sample sheet builds a complete chart", () => {
  for (const name of ["Simplest", "Real-life", "Advanced"]) {
    const org = buildOrg(matrixFrom(name));
    assert.equal(org.fatal, false, name);
    assert.equal(org.people.get(org.ceo).surname, "Goldberg");
    const slides = buildSlides(org, org.ceo, "infinity");
    const seen = new Set();
    for (const slide of slides) {
      const keys = new Set(slide.shapes.filter((shape) => shape.type === "card").map((shape) => shape.key));
      for (const key of keys) seen.add(key);
      const missing = directReports(org, slide.rootKey).all.filter((key) => !keys.has(key));
      assert.deepEqual(missing, [], `${name} ${slide.rootKey}`);
      assert.ok(slide.shapes.some((shape) => shape.type === "stats"));
      assert.ok(slide.shapes.some((shape) => shape.type === "legend"));
      for (const shape of slide.shapes) {
        if (shape.type === "wire") continue;
        if (["title", "legend", "nav", "stats"].includes(shape.type)) continue;
        assert.ok(shape.x >= -1 && shape.x + shape.w <= SLIDE.width + 1, `${name} ${shape.type}`);
        assert.ok(shape.y + shape.h <= SLIDE.height + 1, `${name} ${shape.type} bottom`);
      }
    }
    assert.equal(seen.size, org.people.size, `${name} missing people`);
  }
});

test("Advanced statistics and Larsen’s departments match the template", () => {
  const org = buildOrg(matrixFrom("Advanced"));
  const ceo = org.people.get(org.ceo);
  assert.deepEqual(statisticsLines(ceo, directReports(org, org.ceo).all.length), [
    "Jossef has 5 direct reports",
    "Team Size – 51 people",
    "46 ordinary employees",
    "2 non-ordinary employees",
    "3 outsourcers",
  ]);
  const slides = buildSlides(org, org.ceo, "infinity");
  const larsen = slides.find((slide) => org.people.get(slide.rootKey).surname === "Larsen");
  assert.ok(larsen, "Larsen has a slide");
  const shades = larsen.shapes.filter((shape) => shape.type === "shade").map((shape) => shape.text).sort();
  assert.deepEqual(shades, ["Development", "Research"]);
  const smith = [...org.people.values()].find((person) => person.surname === "Smith");
  assert.match(smith.termination, /July/);
});

test("dates and dotted-line surnames stay clear of the cards", () => {
  const org = buildOrg(matrixFrom("Advanced"));
  const slides = buildSlides(org, org.ceo, "infinity");
  for (const slide of slides) {
    const cards = new Map();
    for (const shape of slide.shapes) {
      if (shape.type === "card") cards.set(shape.key, shape);
    }
    for (const shape of slide.shapes) {
      const card = cards.get(shape.key);
      if (!card || shape === card) continue;
      if (shape.type === "comment") {
        assert.ok(shape.h >= 16, `comment height ${shape.h}`);
        assert.ok(shape.y >= card.y + card.h + 2, `${shape.text} overlaps its card`);
        assert.match(shape.text, /Works until \d/);
      }
      if (shape.type === "bubble") {
        assert.ok(shape.y + shape.h <= card.y - 3, `${shape.text} overlaps ${card.key}`);
      }
    }
  }
  const kim = [...org.people.values()].find((person) => person.surname === "Kim" && person.name === "Jim");
  const kimSlide = slides.find((slide) => slide.shapes.some((shape) => shape.type === "comment" && shape.key === kim.key));
  assert.ok(kimSlide, "Kim's end date is on a slide");
  const note = kimSlide.shapes.find((shape) => shape.type === "comment" && shape.key === kim.key);
  assert.match(note.text, /November/);
});

test("sample charts stay inside the slide", () => {
  for (const name of ["Simplest", "Real-life", "Advanced"]) {
    const org = buildOrg(matrixFrom(name));
    for (const slide of buildSlides(org, org.ceo, "infinity")) {
      for (const shape of slide.shapes) {
        if (shape.type === "wire") {
          for (const value of [shape.x1, shape.x2]) {
            assert.ok(value >= -1 && value <= SLIDE.width + 1, `${name} wire ${value}`);
          }
        } else if (shape.type !== "title" && shape.type !== "legend" && shape.type !== "nav") {
          assert.ok(shape.x >= -1, `${name} ${shape.type} x ${shape.x}`);
          assert.ok(shape.x + shape.w <= SLIDE.width + 1, `${name} ${shape.type} right ${shape.x + shape.w}`);
        }
      }
    }
  }
});

test("manual row layout keeps every person", () => {
  const org = wideOrg(1, 5);
  const lead = directReports(org, org.ceo).managers[0];
  const people = directReports(org, lead).all;
  const box = layoutLeaves(org, people, lead, makeMetrics("Small", 8), "4");
  const cards = box.items.filter((item) => item.type === "card");
  assert.equal(cards.length, people.length);
  assert.ok(box.w > 0 && box.h > 0);
});

test("one direct report hangs straight down from the manager", () => {
  const org = wideOrg(1, 1);
  const lead = directReports(org, org.ceo).managers[0];
  const person = directReports(org, lead).all[0];
  const metrics = makeMetrics("Small", 10);
  const box = layoutLeaves(org, [person], lead, metrics, "auto");
  const card = box.items.find((item) => item.type === "card");
  const tick = box.items.find((item) => item.type === "wire" && Math.abs(item.x1 - item.x2) < 0.2);
  assert.ok(card && tick);
  assert.ok(Math.abs(card.x + card.w / 2 - box.anchorX) < 0.2);
  assert.ok(Math.abs(tick.x1 - box.anchorX) < 0.2);
  assert.equal(Math.min(tick.y1, tick.y2) < 0.2, true);
  assert.ok(Math.abs(Math.max(tick.y1, tick.y2) - card.y) < 0.2);
});

test("a vertical team hangs from a left spine with a tick into each card", () => {
  const org = wideOrg(1, 4);
  const lead = directReports(org, org.ceo).managers[0];
  const people = directReports(org, lead).all;
  const metrics = makeMetrics("Small", 10);
  const box = layoutLeaves(org, people, lead, metrics, "1");
  const cards = box.items.filter((item) => item.type === "card").sort((a, b) => a.y - b.y);
  assert.equal(cards.length, 4);
  assert.ok(cards.every((card) => Math.abs(card.x - metrics.gapX) < 0.1));
  const trunk = box.items.find(
    (item) => item.type === "wire" && Math.abs(item.x1) < 0.1 && Math.abs(item.x2) < 0.1 && item.y2 > item.y1 + metrics.h,
  );
  assert.ok(trunk, "left trunk");
  for (const card of cards) {
    const mid = card.y + card.h / 2;
    const tick = box.items.find(
      (item) =>
        item.type === "wire" &&
        Math.abs(item.y1 - mid) < 0.2 &&
        Math.abs(item.y2 - mid) < 0.2 &&
        item.x1 < card.x - 1 &&
        Math.abs(item.x2 - card.x) < 0.2,
    );
    assert.ok(tick, "side tick");
  }
  const elbow = box.items.find(
    (item) =>
      item.type === "wire" &&
      Math.abs(item.y1 - item.y2) < 0.1 &&
      Math.abs(item.x1 - box.anchorX) < 0.1 &&
      Math.abs(item.x2) < 0.1,
  );
  assert.ok(elbow, "elbow back to the parent");
  assert.ok(cards[0].y > metrics.gapY, "first card sits below the elbow");
});

test("rows share one trunk and a bar on every row", () => {
  const org = wideOrg(1, 8);
  const lead = directReports(org, org.ceo).managers[0];
  const people = directReports(org, lead).all;
  const metrics = makeMetrics("Small", 8);
  const box = layoutLeaves(org, people, lead, metrics, "2");
  const bars = box.items.filter((item) => item.type === "wire" && Math.abs(item.y1 - item.y2) < 0.1 && Math.abs(item.x2 - item.x1) > metrics.gapX);
  const levels = new Set(bars.map((item) => Math.round(item.y1)));
  assert.ok(levels.size >= 2, `row bars ${levels.size}`);
  const trunk = box.items.find(
    (item) => item.type === "wire" && Math.abs(item.x1 - item.x2) < 0.2 && Math.abs(item.y2 - item.y1) > metrics.h,
  );
  assert.ok(trunk, "trunk between rows");
  const cards = box.items.filter((item) => item.type === "card");
  assert.equal(cards.length, 8);
  for (const card of cards) {
    const tick = box.items.find(
      (item) =>
        item.type === "wire" &&
        Math.abs(item.x1 - (card.x + card.w / 2)) < 0.2 &&
        Math.abs(item.x2 - (card.x + card.w / 2)) < 0.2 &&
        Math.min(item.y1, item.y2) < card.y &&
        Math.max(item.y1, item.y2) <= card.y + 0.2,
    );
    assert.ok(tick, "top tick");
  }
});
