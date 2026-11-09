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
import { CARDS, GUIDE, LINK_STYLE, SLIDE, allVisible, buildSlides, chooseFit, departmentShadeAction, dottedDash, hiddenListShape, layoutLeaves, layoutSubordinates, makeMetrics } from "../js/layout.js";

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

test("headcount counts people still to be hired and intern staff", () => {
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
  assert.equal(grace.stats.subordinates, 3);
  assert.equal(grace.stats.intern, 1);
  assert.equal(grace.stats.cards, 3);
  assert.equal(displayName([...org.people.values()].find((person) => person.tbh)), "One");
  assert.equal(headcountLabel(grace), "3 (1)");
});

test("a person who is both important and intern is counted twice and drawn as a split card", () => {
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
  assert.equal(grace.stats.intern, 1);
  assert.equal(grace.stats.ordinary, 0);
  assert.equal(ada.stats.important, 1);
  assert.equal(ada.stats.intern, 1);
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

test("the pill above a card names the other manager: Ada on the solid copy under Grace, Grace on the dotted copy under Ada", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace"),
    row("Ken", "Thompson", "Eng", "Lab", "Grace", "Hopper", "Ada", "Lovelace"),
  ]);
  const ken = [...org.people.values()].find((person) => person.surname === "Ken").key;
  const metrics = makeMetrics("Medium");

  function cardsOf(managerSurname) {
    const manager = [...org.people.values()].find((person) => person.surname === managerSurname);
    const box = layoutSubordinates(org, manager.key, metrics, allVisible(org, manager.key));
    const cards = box.items.filter((shape) => shape.type === "card" && shape.key === ken);
    const bubbles = box.items.filter((shape) => shape.type === "bubble" && shape.key === ken);
    assert.ok(cards.length > 0, managerSurname);
    return cards.map((card) => {
      const bubble = bubbles.find(
        (shape) => Math.abs(shape.x + shape.w / 2 - (card.x + card.w / 2)) < 2 && shape.y + shape.h <= card.y + 0.5,
      );
      return { card, bubble };
    });
  }

  const underSolid = cardsOf("Grace");
  const underDotted = cardsOf("Ada").filter((pair) => pair.card.managerType === "Dotted");
  assert.equal(underSolid.length, 1);
  assert.equal(underDotted.length, 1);
  for (const pair of [...underSolid, ...underDotted]) {
    assert.ok(pair.bubble, `${pair.card.managerType} copy has a pill`);
    assert.equal(pair.bubble.text, pair.card.managerType === "Dotted" ? "Grace" : "Ada");
    assert.equal(pair.bubble.dotted, true);
  }
});

test("solid and dotted ticks on a two-manager card do not share a path", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace"),
    row("Ken", "Thompson", "Eng", "Lab", "Grace", "Hopper", "Ada", "Lovelace"),
  ]);
  const ken = [...org.people.values()].find((person) => person.surname === "Ken").key;
  const metrics = makeMetrics("Large");

  function ticksTouching(managerSurname, managerType) {
    const manager = [...org.people.values()].find((person) => person.surname === managerSurname);
    const box = layoutSubordinates(org, manager.key, metrics, allVisible(org, manager.key));
    const card = box.items.find((shape) => shape.type === "card" && shape.key === ken && shape.managerType === managerType);
    assert.ok(card, `${managerSurname} ${managerType}`);
    return box.items.filter((shape) => {
      if (shape.type !== "wire" || Math.abs(shape.x1 - shape.x2) > 0.8) return false;
      const top = Math.min(shape.y1, shape.y2);
      const bottom = Math.max(shape.y1, shape.y2);
      return bottom >= card.y - 0.6 && top <= card.y + 0.6 && Math.abs(shape.x1 - (card.x + card.w / 2)) < 20;
    });
  }

  for (const [manager, kind] of [["Grace", "Solid"], ["Ada", "Dotted"]]) {
    const wires = ticksTouching(manager, kind);
    const solid = wires.filter((shape) => !shape.dotted);
    const dotted = wires.filter((shape) => shape.dotted);
    assert.ok(solid.length > 0 && dotted.length > 0, `${kind} shows both line styles`);
    for (const left of solid) {
      for (const right of dotted) {
        const shared = Math.min(Math.max(left.y1, left.y2), Math.max(right.y1, right.y2))
          - Math.max(Math.min(left.y1, left.y2), Math.min(right.y1, right.y2));
        if (shared < 0.4) continue;
        assert.equal(Math.abs(left.x1 - right.x1), 5, `${kind} ticks are 5 points apart`);
      }
    }
  }
});

test("a thick dotted connector keeps a visible gap", () => {
  for (const weight of [2, 3, 4]) {
    const [dot, gap] = dottedDash(weight);
    assert.ok(gap - weight > dot, `weight ${weight} gap stays open past the round caps`);
  }
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
    "3 interns",
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

test("department shades toggle off on the next click, unless one shade is selected", () => {
  const shapes = [
    { id: "s1", type: "shade" },
    { id: "c1", type: "card" },
  ];
  assert.equal(departmentShadeAction([], new Set()), "add");
  assert.equal(departmentShadeAction(shapes, new Set()), "remove");
  assert.equal(departmentShadeAction(shapes, new Set(["s1"])), "lighten");
  assert.equal(departmentShadeAction(shapes, new Set(["s1", "c1"])), "remove");
});

test("slide links sit below the card and below an end date", () => {
  const check = (org) => {
    const slides = buildSlides(org, org.ceo, "infinity");
    let links = 0;
    for (const slide of slides) {
      for (const link of slide.shapes.filter((shape) => shape.type === "link")) {
        links += 1;
        const card = slide.shapes.find((shape) => shape.type === "card" && shape.key === link.key);
        assert.ok(card, link.text);
        assert.ok(link.y + 0.1 >= card.y + card.h, `${link.text} overlaps its card`);
        const note = slide.shapes.find((shape) => shape.type === "comment" && shape.key === link.key);
        if (note) assert.ok(link.y + 0.1 >= note.y + note.h, `${link.text} overlaps ${note.text}`);
      }
    }
    assert.ok(links > 0);
  };

  check(buildOrg(matrixFrom("Advanced")));

  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const rows = [header, row("Ada", "Lovelace", "CEO", "Office", "", "")];
  for (let team = 0; team < 10; team += 1) {
    const lead = `Lead${team}`;
    rows.push(row(lead, "Manager", "Director", `Dept ${team % 3}`, "Ada", "Lovelace", "", "", "USA, NY", "9 November 2018"));
    for (let person = 0; person < 8; person += 1) {
      rows.push(row(`P${team}_${person}`, "Member", "IC", `Dept ${team % 3}`, lead, "Manager"));
    }
  }
  check(buildOrg(rows));
});

test("slide links use the palette green bar and black type", () => {
  assert.equal(LINK_STYLE.fill.toLowerCase(), "#9bc225");
  assert.equal(LINK_STYLE.color.toLowerCase(), "#000000");
  const css = readFileSync(new URL("../css/app.css", import.meta.url), "utf8");
  const block = css.slice(css.indexOf(".link {"), css.indexOf(".link {") + 500);
  assert.match(block, /#9[Bb][Cc]225/);
  assert.match(block, /color:\s*#000/);
  const exp = readFileSync(new URL("../js/export.js", import.meta.url), "utf8");
  assert.match(exp, /LINK_STYLE/);
});

test("leaf cards run top to bottom: ordinary, important, interns, then to be hired", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("A", "Intern", "Eng", "Lab", "Ada", "Lovelace", "", "", "", "", "1"),
    row("B", "Key", "Eng", "Lab", "Ada", "Lovelace", "", "", "", "", "", "1"),
    row("C", "Plain", "Eng", "Lab", "Ada", "Lovelace"),
    row("D", "Both", "Eng", "Lab", "Ada", "Lovelace", "", "", "", "", "1", "1"),
    row("E", "Plain", "Eng", "Lab", "Ada", "Lovelace"),
    row("TBH_F", "Open", "Eng", "Lab", "Ada", "Lovelace", "", "", "", "", "", "1"),
  ]);
  assert.equal(org.fatal, false, org.issues.map((i) => i.message).join(" | "));
  const keys = org.reports.get(org.ceo);
  const box = layoutLeaves(org, keys, org.ceo, makeMetrics("Medium"), "1");
  const names = box.items
    .filter((item) => item.type === "card")
    .sort((a, b) => a.y - b.y)
    .map((card) => org.people.get(card.key).surname);
  assert.deepEqual(names, ["C", "E", "B", "A", "D", "TBH_F"]);
});

test("giving two employees one department groups them in one branch with one shade", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("A1", "x", "Eng", "Q", "Ada", "Lovelace"),
    row("A2", "x", "Eng", "X", "Ada", "Lovelace"),
    row("B1", "x", "Eng", "q", "Ada", "Lovelace"),
    row("B2", "x", "Eng", "Y", "Ada", "Lovelace"),
  ]);
  assert.equal(org.fatal, false, org.issues.map((i) => i.message).join(" | "));
  const [slide] = buildSlides(org, org.ceo, "clever");
  const at = (name) => slide.shapes.find((shape) => shape.type === "card" && org.people.get(shape.key)?.surname === name);
  assert.equal(at("A1").x, at("B1").x, "A1 and B1 share a column");
  assert.ok(Math.abs(at("A1").y - at("B1").y) > 0, "stacked, not side by side");
  const shades = slide.shapes.filter((shape) => shape.type === "shade");
  assert.equal(shades.length, 3);
  const q = shades.find((shade) => shade.text === "Q");
  for (const name of ["A1", "B1"]) {
    const card = at(name);
    assert.ok(q.x <= card.x && q.x + q.w >= card.x + card.w && q.y <= card.y && q.y + q.h >= card.y + card.h);
  }
  for (const name of ["A2", "B2"]) {
    const card = at(name);
    assert.ok(card.x >= q.x + q.w || card.x + card.w <= q.x, `${name} sits outside the Q shade`);
  }
});

test("a department shade around two people stays a column when everyone else has no department", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const rows = [
    header,
    row("Ada", "Lovelace", "CEO", "", "", ""),
    row("M1", "m", "mgr", "", "Ada", "Lovelace"),
    row("M2", "m", "mgr", "", "Ada", "Lovelace"),
    row("W1", "x", "eng", "GPS", "M1", "m"),
    row("W2", "x", "eng", "GPS", "M1", "m"),
  ];
  for (let i = 1; i <= 4; i += 1) rows.push(row(`V${i}`, "x", "eng", "", "M1", "m"));
  for (let i = 1; i <= 3; i += 1) rows.push(row(`N${i}`, "x", "eng", "", "M2", "m"));
  const org = buildOrg(rows);
  assert.equal(org.fatal, false, org.issues.map((i) => i.message).join(" | "));
  const [slide] = buildSlides(org, org.ceo, "clever");
  const shades = slide.shapes.filter((shape) => shape.type === "shade");
  assert.equal(shades.length, 1);
  const [shade] = shades;
  const card = (name) => slide.shapes.find((s) => s.type === "card" && org.people.get(s.key)?.surname === name);
  for (const name of ["W1", "W2"]) {
    assert.ok(card(name).x >= shade.x && card(name).x + card(name).w <= shade.x + shade.w, `${name} is inside`);
  }
  for (const name of ["V1", "N1", "M2"]) {
    const c = card(name);
    assert.ok(c.x >= shade.x + shade.w || c.x + c.w <= shade.x, `${name} is not covered by the shade`);
  }
  assert.ok(shade.w < 400, "a column, not the whole slide");
});

function hiddenRow(...cells) {
  const base = row(...cells);
  base.push("", "1");
  return base;
}

test("hidden employees are checked on upload but left off slides and counts, and tallied apart", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c", "h"];
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace"),
    row("Ken", "Thompson", "Eng", "Lab", "Grace", "Hopper"),
    hiddenRow("Hid", "Den", "Eng", "Lab", "Grace", "Hopper"),
    hiddenRow("Hid", "Intern", "Eng", "Lab", "Grace", "Hopper", "", "", "", "", "1"),
  ]);
  assert.equal(org.fatal, false, org.issues.map((i) => i.message).join(" | "));
  const grace = [...org.people.values()].find((person) => person.surname === "Grace");
  assert.equal(grace.stats.subordinates, 2, "Grace and Ken only");
  assert.equal(grace.stats.hidden, 2);
  assert.equal(grace.stats.hiddenOrdinary, 1);
  assert.equal(grace.stats.hiddenIntern, 1);
  assert.equal(org.companySize, 3);
  assert.equal(org.order.length, 3, "hidden people are not offered in the pickers");
  const lines = statisticsLines(grace, 1);
  assert.equal(lines.some((line) => /hidden/i.test(line)), false, "hidden people are not in the statistics box");
  assert.equal(lines.length, 5);
  assert.equal(
    hiddenListShape(org, org.ceo).text,
    "Hidden, not counted in the statistics – 2 people (1 ordinary, 0 non-ordinary, 1 interns): Hid Den, Hid Intern",
  );
  const [slide] = buildSlides(org, org.ceo, "clever");
  const surnames = slide.shapes.filter((shape) => shape.type === "card").map((card) => org.people.get(card.key).surname);
  assert.equal(surnames.includes("Hid"), false);
  assert.equal(surnames.length, 3);
});

test("a hidden employee still has to reach the CEO", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c", "h"];
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    hiddenRow("Hid", "Den", "Eng", "Lab", "Nobody", "Here"),
  ]);
  assert.equal(org.fatal, true);
});

test("a hidden manager must not have visible reports, and the CEO cannot be hidden", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c", "h"];
  const visibleUnderHidden = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    hiddenRow("Hid", "Den", "Mgr", "Lab", "Ada", "Lovelace"),
    row("Ken", "Thompson", "Eng", "Lab", "Hid", "Den"),
  ]);
  assert.equal(visibleUnderHidden.fatal, true);
  assert.match(visibleUnderHidden.issues[0].message, /hidden/);
  const allHidden = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    hiddenRow("Hid", "Den", "Mgr", "Lab", "Ada", "Lovelace"),
    hiddenRow("Ken", "Thompson", "Eng", "Lab", "Hid", "Den"),
  ]);
  assert.equal(allHidden.fatal, false, allHidden.issues.map((i) => i.message).join(" | "));
  assert.equal(allHidden.companySize, 1);
  const hiddenCeo = buildOrg([header, hiddenRow("Ada", "Lovelace", "CEO", "Office", "", "")]);
  assert.equal(hiddenCeo.fatal, true);
});

test("the head card tick is as thick as the bar it lands on and stops at that bar", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c"];
  const rows = [header, row("Ada", "Lovelace", "CEO", "Office", "", "")];
  for (let m = 1; m <= 12; m += 1) {
    rows.push(row(`M${m}`, "m", "mgr", `D${m}`, "Ada", "Lovelace"));
    for (let i = 1; i <= 3; i += 1) rows.push(row(`M${m}t${i}`, "x", "eng", `D${m}`, `M${m}`, "m"));
  }
  const org = buildOrg(rows);
  assert.equal(org.fatal, false, org.issues.map((i) => i.message).join(" | "));
  for (const mode of ["clever", "infinity"]) {
    for (const slide of buildSlides(org, org.ceo, mode)) {
      const wires = slide.shapes.filter((shape) => shape.type === "wire");
      const tick = wires.find((w) => w.x1 === SLIDE.width / 2 && w.x2 === SLIDE.width / 2 && w.y2 - w.y1 === 15);
      if (!tick) continue;
      const bar = wires.find((w) => w.y1 === w.y2 && Math.abs(w.y1 - tick.y2) < 0.01 && Math.min(w.x1, w.x2) <= tick.x1 && Math.max(w.x1, w.x2) >= tick.x1);
      if (!bar) continue;
      assert.equal(tick.weight, bar.weight, "same thickness, so the tick does not poke out below the bar");
    }
  }
});

test("hidden people are named in a small grey note along the bottom of the slide", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c", "h"];
  const none = buildOrg([header, row("Ada", "Lovelace", "CEO", "Office", "", ""), row("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace")]);
  assert.equal(hiddenListShape(none, none.ceo), null);
  assert.equal(buildSlides(none, none.ceo, "clever")[0].shapes.some((shape) => shape.type === "hiddenlist"), false);

  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace"),
    hiddenRow("Zed", "Last", "Eng", "Lab", "Grace", "Hopper"),
    hiddenRow("Alan", "Turing", "Eng", "Lab", "Grace", "Hopper"),
  ]);
  assert.equal(org.fatal, false, org.issues.map((i) => i.message).join(" | "));
  const [slide] = buildSlides(org, org.ceo, "clever");
  const note = slide.shapes.find((shape) => shape.type === "hiddenlist");
  assert.equal(note.text, "Hidden, not counted in the statistics – 2 people (2 ordinary, 0 non-ordinary, 0 interns): Alan Turing, Zed Last");
  assert.ok(note.y + note.h <= SLIDE.height, "inside the slide");
  assert.ok(note.y > 480, "along the bottom edge");
});

test("hidden people get no card in any slide of a multi-slide chart", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c", "h"];
  const rows = [header, row("Ada", "Lovelace", "CEO", "Office", "", "")];
  for (let m = 1; m <= 12; m += 1) {
    rows.push(row(`M${m}`, "m", "mgr", `D${m}`, "Ada", "Lovelace"));
    for (let i = 1; i <= 3; i += 1) rows.push(row(`M${m}t${i}`, "x", "eng", `D${m}`, `M${m}`, "m"));
  }
  rows.push(hiddenRow("Hid", "Den", "eng", "D1", "M1", "m"));
  rows.push(hiddenRow("Hid", "Top", "eng", "D1", "Ada", "Lovelace"));
  const org = buildOrg(rows);
  assert.equal(org.fatal, false, org.issues.map((i) => i.message).join(" | "));
  for (const mode of ["clever", "infinity"]) {
    for (const slide of buildSlides(org, org.ceo, mode)) {
      const surnames = slide.shapes.filter((shape) => shape.type === "card").map((card) => org.people.get(card.key).surname);
      assert.equal(surnames.includes("Hid"), false, `${mode}: no card for a hidden person`);
    }
  }
});

test("the newcomer flag is read from column P and does not change any count", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c", "h", "x"];
  const withFlag = (...cells) => {
    const r = row(...cells);
    r.push("", "", "1");
    return r;
  };
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    withFlag("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace"),
    row("Ken", "Thompson", "Eng", "Lab", "Ada", "Lovelace"),
  ]);
  assert.equal(org.fatal, false, org.issues.map((i) => i.message).join(" | "));
  const byName = (surname) => [...org.people.values()].find((person) => person.surname === surname);
  assert.equal(byName("Grace").newcomer, true);
  assert.equal(byName("Ken").newcomer, false);
  assert.equal(byName("Ada").newcomer, false);
  assert.equal(byName("Ada").stats.subordinates, 3);
});

test("a 16-column sheet puts HR ID in N, hide in O and newcomer in P", () => {
  const header = ["s", "n", "j", "d", "ms", "mn", "ds", "dn", "l", "t", "o", "i", "c", "hr", "h", "x"];
  const org = buildOrg([
    header,
    [...row("Ada", "Lovelace", "CEO", "Office", "", ""), "42", "", ""],
    [...row("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace"), "7", "", "1"],
    [...row("Ken", "Thompson", "Eng", "Lab", "Ada", "Lovelace"), "8", "1", ""],
  ]);
  assert.equal(org.fatal, false, org.issues.map((i) => i.message).join(" | "));
  const byName = (surname) => [...org.people.values()].find((person) => person.surname === surname);
  assert.equal(byName("Grace").hrId, "7");
  assert.equal(byName("Grace").newcomer, true);
  assert.equal(byName("Ken").hidden, true);
  assert.equal(byName("Ada").stats.subordinates, 2);
});

test("the newcomer date in column Q is kept in compact day.month form", async () => {
  const { compactDate } = await import("../js/org.js");
  assert.equal(compactDate("16.11.2026"), "16.11");
  assert.equal(compactDate("5/3/2026"), "05.03");
  assert.equal(compactDate("2026-11-16"), "16.11");
  assert.equal(compactDate(new Date(Date.UTC(2026, 10, 16))), "16.11");
  assert.equal(compactDate(46342), "16.11");
  assert.equal(compactDate(""), "");
  const header = Array(17).fill("x");
  const org = buildOrg([
    header,
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    [...row("Grace", "Hopper", "Eng", "Lab", "Ada", "Lovelace"), "", "", "1", "16.11.2026"],
  ]);
  assert.equal(org.fatal, false, org.issues.map((i) => i.message).join(" | "));
  assert.equal(org.people.get("gracehopper").newcomerDate, "16.11");
});

test("vacancies get their own picker category and a readable label", async () => {
  const { categories, pickerLabel, VACANCIES } = await import("../js/org.js");
  const org = buildOrg([
    ["s", "n", "j", "d", "ms", "mn"],
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("TBH_Asyat", "TBH_Asyat", "Eng", "Lab", "Ada", "Lovelace"),
  ]);
  assert.equal(org.fatal, false);
  assert.ok(categories(org).includes(VACANCIES));
  assert.equal(categories(org).includes("T"), false);
  const [key] = peopleInCategory(org, VACANCIES);
  assert.equal(pickerLabel(org.people.get(key)), "Asyat (vacancy)");
  assert.equal(peopleInCategory(org, "T").length, 0);
});

test("Cyrillic surnames get their own letter category, so clicking their card finds them", async () => {
  const { categories } = await import("../js/org.js");
  const org = buildOrg([
    ["s", "n", "j", "d", "ms", "mn"],
    row("Толстой", "Никита", "CEO", "Office", "", ""),
    row("Апонасенко", "Юлия", "Analyst", "Lab", "Толстой", "Никита"),
    row("Ada", "Lovelace", "Eng", "Lab", "Толстой", "Никита"),
  ]);
  assert.equal(org.fatal, false);
  const list = categories(org);
  assert.ok(list.includes("А") && list.includes("Т") && list.includes("A"), list.join(","));
  assert.deepEqual(peopleInCategory(org, "А"), ["апонасенкоюлия"]);
});

test("cards keep their own copy of the person, so a later import cannot change a slide", async () => {
  const { shapePerson } = await import("../js/org.js");
  const rows = (job) => [
    ["s", "n", "j", "d", "ms", "mn"],
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Grace", "Hopper", job, "Lab", "Ada", "Lovelace"),
  ];
  const first = buildOrg(rows("Engineer"));
  const [slide] = buildSlides(first, first.ceo, "clever");
  const other = buildOrg(rows("Janitor"));
  const card = slide.shapes.find((shape) => shape.type === "card" && shape.key === "gracehopper");
  assert.equal(shapePerson(other, card).job, "Engineer", "the later sheet does not leak in");
  assert.equal(shapePerson(buildOrg([["s"]]), card).surname, "Grace", "and a sheet without this person does not blank it");
  assert.equal(slide.rootLabel, "Ada Lovelace");
});

test("the pill over a two-manager card names the other manager, whichever line the card hangs under", () => {
  const org = buildOrg([
    ["s", "n", "j", "d", "ms", "mn", "ds", "dn"],
    row("Ada", "Lovelace", "CEO", "Office", "", ""),
    row("Mike", "Boss", "Mgr", "Lab", "Ada", "Lovelace"),
    row("Sem", "Dotted", "Mgr", "Lab", "Ada", "Lovelace"),
    row("Ust", "Person", "Eng", "Lab", "Mike", "Boss", "Sem", "Dotted"),
  ]);
  assert.equal(org.fatal, false, org.issues.map((i) => i.message).join(" | "));
  const ust = org.people.get("ustperson");
  const metrics = makeMetrics("Medium", 10);
  const pill = (parent) =>
    layoutLeaves(org, ["ustperson"], parent, metrics, "1").items.find((item) => item.type === "bubble").text;
  assert.equal(pill("mikeboss"), "Sem", "under the solid-line manager: the dotted-line one");
  assert.equal(pill("semdotted"), "Mike", "under the dotted-line manager: the solid-line one");
  assert.equal(ust.solidKey, "mikeboss");
});
