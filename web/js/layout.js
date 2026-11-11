/**
 * Org-chart geometry.
 *
 * The PowerPoint tool measures a tree, hides the largest teams until the
 * drawing fits a 920×350 pt area, then grows card size and spacing. The same
 * decisions are made here. Coordinates are PowerPoint points on a 960×540 slide.
 */

import {
  displayName,
  freezeCards,
  departmentsOf,
  directReports,
  hasTwoManagers,
  isManager,
  managerType,
  moderated,
  personFullName,
  statisticsLines,
} from "./org.js";

export const SLIDE = { width: 960, height: 540 };
export const GUIDE = { width: 920, height: 360 };
export const MAX_PEOPLE_FOR_2D = 6;
export const MAX_PEOPLE_FOR_4D = 14;
export const MAX_BRANCHES = 7;
export const TREE_TOP = 132;

export const CARD_SIZES = ["ExtraSmall", "Small", "Medium", "Large"];

/** name / title = type size in pt for the person's name and job line. A job line under 5 pt is unreadable on a slide, so Small and up keep it at 5.5 pt or more. */
export const CARDS = {
  ExtraSmall: { w: 70.9, h: 26.2, name: 7, title: 4.5, bubbleW: 52.2, bubbleH: 7.5, commentH: 6.2, line: 2 },
  Small: { w: 70.9, h: 35.5, name: 7, title: 5.5, bubbleW: 52.2, bubbleH: 7.5, commentH: 6.2, line: 2 },
  Medium: { w: 87.9, h: 42.3, name: 9, title: 6.5, bubbleW: 65.2, bubbleH: 9.4, commentH: 7.7, line: 3 },
  Large: { w: 108.5, h: 54.5, name: 11, title: 7.5, bubbleW: 81.1, bubbleH: 11.6, commentH: 8.8, line: 4 },
};

export const SHADE_COLORS = ["#DEEBF7C4", "#BDD7EEC4", "#E7EEF6C4"];

/** Palette SlideN: bold black type on a green bar. */
export const LINK_STYLE = { fill: "#9BC225", color: "#000000" };

/**
 * Departments button, matching createDepartments.
 * One selected shade lightens. Otherwise existing shades are removed, and a
 * later click builds them again.
 */
export function departmentShadeAction(shapes, selectedIds = new Set()) {
  const shades = shapes.filter((shape) => shape.type === "shade");
  const selectedShades = shades.filter((shape) => selectedIds.has(shape.id));
  if (selectedIds.size === 1 && selectedShades.length === 1) return "lighten";
  if (shades.length) return "remove";
  return "add";
}

/** Solid and dotted connector colours from applyLineStyle. */
export const LINE = { solid: "#5B9BD5", dotted: "#002060" };

/**
 * Round dots whose gap stays open after the stroke caps.
 * A fixed 2.8 gap closes up once the line is thicker than that gap.
 */
export function dottedDash(weight = 2) {
  const size = Number(weight) || 2;
  return [0.1, size * 2];
}

export function makeMetrics(size = "Medium", gapX = 10, gapExt = gapX, gapY = gapX) {
  const card = CARDS[size] || CARDS.Medium;
  return { ...card, size, gapX, gapY, gapExt };
}

const COMMENT_GAP = 3;

/** Two lines, so "Works until 9 November 2018" is not cut off mid-word. */
function commentHeight(metrics) {
  if (metrics.size === "Large") return 20;
  if (metrics.size === "Medium") return 18;
  return 16;
}

function commentSpace(metrics) {
  return COMMENT_GAP + commentHeight(metrics);
}

/** Clearance that keeps a surname pill off the card and off the connector. */
function bubbleGap(metrics) {
  return (metrics.line || 2) + 4;
}

function bubbleStack(metrics) {
  return metrics.bubbleH + bubbleGap(metrics);
}

function parentDrop(org, parentKey, metrics) {
  const parent = org.people.get(parentKey);
  return metrics.gapY + (parent?.termination ? commentSpace(metrics) : 0);
}

/** Top of a “See slide” bar: under the card, and under its end-date note. */
export function slideLinkTop(card, shapes) {
  let y = card.y + card.h;
  for (const shape of shapes) {
    if (shape.type === "comment" && shape.key === card.key) y = Math.max(y, shape.y + shape.h);
  }
  return y;
}

function commentItem(key, x, cardBottom, metrics, text) {
  return {
    type: "comment",
    key,
    x,
    y: cardBottom + COMMENT_GAP,
    w: metrics.w,
    h: commentHeight(metrics),
    z: 4,
    text: `Works until ${text}`,
  };
}

/**
 * The pill over a two-manager card names the manager the card is NOT hanging under:
 * under the solid-line manager it shows the dotted-line one, and under the dotted-line manager the solid-line one.
 */
export function otherManagerSurname(org, person, shownUnder) {
  const other = shownUnder === "Dotted" ? person.solidKey : person.dottedKey;
  return org.people.get(other)?.surname || "";
}

function addBubble(items, org, person, key, centerX, cardY, metrics, shownUnder, offsetTick) {
  const gap = bubbleGap(metrics);
  // Upward ticks are 5 points apart. The second tick uses the other line style,
  // so a dotted parent line does not cover the solid one.
  const nudge = offsetTick ? 5 : 0;
  const bottom = cardY - gap;
  items.push(wire(centerX + nudge, bottom, centerX + nudge, cardY, shownUnder !== "Dotted", metrics));
  items.push({
    type: "bubble",
    key,
    x: centerX - metrics.bubbleW / 2,
    y: bottom - metrics.bubbleH,
    w: metrics.bubbleW,
    h: metrics.bubbleH,
    z: 4,
    dotted: true,
    text: otherManagerSurname(org, person, shownUnder),
  });
}

function translate(items, dx, dy) {
  return items.map((item) => {
    if (item.type === "wire") {
      return { ...item, x1: item.x1 + dx, y1: item.y1 + dy, x2: item.x2 + dx, y2: item.y2 + dy };
    }
    return { ...item, x: item.x + dx, y: item.y + dy };
  });
}

function shiftMap(count) {
  const shifts = [];
  let cursor = -1 * Math.floor(count / 2);
  for (let i = 0; i < count; i += 1) {
    const sign = cursor < 0 ? -1 : cursor > 0 ? 1 : 0;
    shifts.push(cursor - sign * 0.5);
    cursor += 1;
    if (count % 2 === 0 && cursor === 0) cursor += 1;
  }
  return shifts;
}

function wire(x1, y1, x2, y2, dotted, metrics) {
  return {
    type: "wire",
    x1,
    y1,
    x2,
    y2,
    dotted: Boolean(dotted),
    weight: metrics.line,
  };
}

function cardItem(key, x, y, metrics, shownUnder) {
  return {
    type: "card",
    key,
    x,
    y,
    w: metrics.w,
    h: metrics.h,
    size: metrics.size,
    managerType: shownUnder,
    collapsed: false,
    head: false,
  };
}

/**
 * One card with the tick pointing up. The tick apex is the local origin,
 * which is where this card meets its parent connector.
 */
function leafLayout(org, key, metrics, shownUnder) {
  const person = org.people.get(key);
  const two = hasTwoManagers(person);
  const tick = metrics.gapY + (two ? bubbleStack(metrics) : 0);
  const anchorX = metrics.w / 2;
  const items = [wire(anchorX, 0, anchorX, tick, shownUnder === "Dotted", metrics)];
  if (two) addBubble(items, org, person, key, anchorX, tick, metrics, shownUnder, true);
  items.push(cardItem(key, 0, tick, metrics, shownUnder));
  let height = tick + metrics.h;
  if (person.termination) {
    items.push(commentItem(key, 0, tick + metrics.h, metrics, person.termination));
    height += commentSpace(metrics);
  }
  return { w: metrics.w, h: height, anchorX, items };
}

/**
 * Vertical team. Matches make1DSubordinates: a trunk on the left, a tick into
 * the middle of each card's left edge, and an elbow back to the parent.
 */
function columnLayout(org, people, parentKey, metrics) {
  const gapX = metrics.gapX;
  const gapY = metrics.gapY;
  const anchorX = gapX + metrics.w / 2;
  const drop = parentDrop(org, parentKey, metrics);
  const items = [];
  let y = drop;
  let lowestSolid = null;
  let lowestDotted = null;
  let bottom = drop;

  for (const key of people) {
    const person = org.people.get(key);
    const shownUnder = managerType(person, parentKey) || "Solid";
    const two = hasTwoManagers(person);
    y += gapY + metrics.h / 2;
    if (two) y += bubbleStack(metrics);
    const tickY = y;
    const cardY = tickY - metrics.h / 2;
    const cardX = gapX;
    items.push(wire(0, tickY, cardX, tickY, shownUnder === "Dotted", metrics));
    if (two) addBubble(items, org, person, key, cardX + metrics.w / 2, cardY, metrics, shownUnder);
    items.push(cardItem(key, cardX, cardY, metrics, shownUnder));
    if (person.termination) {
      items.push(commentItem(key, cardX, cardY + metrics.h, metrics, person.termination));
    }
    if (shownUnder === "Dotted") lowestDotted = tickY;
    else lowestSolid = tickY;
    y += metrics.h / 2;
    if (person.termination) y += commentSpace(metrics);
    bottom = Math.max(bottom, y);
  }

  const spine = (dotted, lowest) => {
    if (lowest == null) return;
    items.push(wire(anchorX, 0, anchorX, drop, dotted, metrics));
    items.push(wire(anchorX, drop, 0, drop, dotted, metrics));
    if (lowest - drop > 0.4) items.push(wire(0, drop, 0, lowest, dotted, metrics));
  };
  spine(false, lowestSolid);
  spine(true, lowestDotted);
  return { w: gapX + metrics.w, h: bottom, anchorX, items };
}

/**
 * Several cards on each row. Matches makeNDSubordinates: every row has a
 * horizontal bar, each card hangs from that bar by a top tick, and one trunk
 * joins the rows under the parent.
 */
function rowsLayout(org, people, parentKey, metrics, perRow) {
  const pitch = metrics.gapX + metrics.w;
  const shifts = shiftMap(perRow);
  const gapY = metrics.gapY;
  const drop = parentDrop(org, parentKey, metrics);
  const items = [];
  let y = drop;
  let lowestSolid = null;
  let lowestDotted = null;
  let bottom = drop;

  for (let start = 0; start < people.length; start += perRow) {
    const row = people.slice(start, start + perRow);
    const bubbleRow = row.some((key) => hasTwoManagers(org.people.get(key)));
    for (let i = 0; i < row.length; i += 1) {
      const person = org.people.get(row[i]);
      const shownUnder = managerType(person, parentKey) || "Solid";
      const two = hasTwoManagers(person);
      const tick = gapY + (bubbleRow ? bubbleStack(metrics) : 0);
      const center = shifts[i] * pitch;
      const cardY = y + tick;
      const cardX = center - metrics.w / 2;
      items.push(wire(center, y, center, cardY, shownUnder === "Dotted", metrics));
      if (Math.abs(center) > 0.4) items.push(wire(0, y, center, y, shownUnder === "Dotted", metrics));
      if (two) addBubble(items, org, person, row[i], center, cardY, metrics, shownUnder, true);
      items.push(cardItem(row[i], cardX, cardY, metrics, shownUnder));
      if (person.termination) {
        items.push(commentItem(row[i], cardX, cardY + metrics.h, metrics, person.termination));
      }
      if (shownUnder === "Dotted") lowestDotted = y;
      else lowestSolid = y;
      bottom = Math.max(bottom, cardY + metrics.h + (person.termination ? commentSpace(metrics) : 0));
    }
    const commentRow = row.some((key) => org.people.get(key)?.termination);
    const aligned = metrics.h + gapY + (bubbleRow ? bubbleStack(metrics) : 0) + (commentRow ? commentSpace(metrics) : 0);
    y += gapY + aligned;
  }

  const spine = (dotted, lowest) => {
    if (lowest == null) return;
    items.push(wire(0, 0, 0, drop, dotted, metrics));
    if (lowest - drop > 0.4) items.push(wire(0, drop, 0, lowest, dotted, metrics));
  };
  spine(false, lowestSolid);
  spine(true, lowestDotted);

  let minLeft = Infinity;
  let maxRight = -Infinity;
  for (const item of items) {
    if (item.type === "wire") {
      minLeft = Math.min(minLeft, item.x1, item.x2);
      maxRight = Math.max(maxRight, item.x1, item.x2);
    } else {
      minLeft = Math.min(minLeft, item.x);
      maxRight = Math.max(maxRight, item.x + (item.w || 0));
    }
  }
  const shift = Number.isFinite(minLeft) ? -minLeft : 0;
  return {
    w: Number.isFinite(minLeft) ? maxRight - minLeft : 0,
    h: bottom,
    anchorX: shift,
    items: translate(items, shift, 0),
  };
}

/** Ordinary staff, then important employees, then interns, then vacancies. Someone who is both important and an intern sorts as an intern; a vacancy always goes last. */
function leafRank(person) {
  if (person?.tbh) return 3;
  if (person?.intern) return 2;
  if (person?.important) return 1;
  return 0;
}

export function layoutLeaves(org, unsortedPeople, parentKey, metrics, mode = "auto") {
  // Array.prototype.sort is stable, so the existing order holds inside each group.
  const people = [...unsortedPeople].sort(
    (a, b) => leafRank(org.people.get(a)) - leafRank(org.people.get(b)),
  );
  const count = people.length;
  if (count === 0) return { w: 0, h: 0, anchorX: 0, items: [] };
  // branch_auto draws one person as a card with the tick pointing up.
  if (mode === "auto" && count === 1) {
    const shownUnder = managerType(org.people.get(people[0]), parentKey) || "Solid";
    return leafLayout(org, people[0], metrics, shownUnder);
  }
  if (mode === "1" || (mode === "auto" && count <= MAX_PEOPLE_FOR_2D)) {
    return columnLayout(org, people, parentKey, metrics);
  }
  if (mode === "2" || (mode === "auto" && count <= MAX_PEOPLE_FOR_4D)) {
    return rowsLayout(org, people, parentKey, metrics, 2);
  }
  if (mode === "4" || mode === "auto") {
    return rowsLayout(org, people, parentKey, metrics, 4);
  }
  const shownUnder = managerType(org.people.get(people[0]), parentKey) || "Solid";
  return leafLayout(org, people[0], metrics, shownUnder);
}

function markCollapsed(box) {
  box.items = box.items.map((item) => (item.type === "card" ? { ...item, collapsed: true } : item));
  return box;
}

function hangPoints(box) {
  const verticals = box.items.filter(
    (item) => item.type === "wire" && Math.abs(item.x1 - item.x2) < 0.8 && Math.min(item.y1, item.y2) <= 0.5,
  );
  if (!verticals.length) return [{ x: box.anchorX, dotted: box.line === "dotted" }];
  const longest = Math.max(...verticals.map((item) => Math.abs(item.y2 - item.y1)));
  return verticals
    .filter((item) => Math.abs(item.y2 - item.y1) >= longest - 1)
    .map((item) => ({ x: (item.x1 + item.x2) / 2, dotted: Boolean(item.dotted) }));
}

function placeBranches(branches, metrics) {
  if (branches.length === 0) return { w: 0, h: 0, anchorX: 0, items: [] };
  let x = 0;
  const placed = [];
  for (const branch of branches) {
    placed.push({ branch, x });
    x += branch.box.w + metrics.gapExt;
  }
  const totalW = Math.max(0, x - metrics.gapExt);
  const stem = branches.length >= 2 ? metrics.gapY : 0;
  const items = [];
  const hangs = [];
  let maxH = 0;
  for (const entry of placed) {
    for (const hang of hangPoints(entry.branch.box)) {
      hangs.push({ x: entry.x + hang.x, dotted: hang.dotted });
    }
    items.push(...translate(entry.branch.box.items, entry.x, stem));
    maxH = Math.max(maxH, stem + entry.branch.box.h);
  }
  const xs = hangs.map((hang) => hang.x);
  const anchorX = xs.length ? (Math.min(...xs) + Math.max(...xs)) / 2 : totalW / 2;
  const drawBar = (dotted) => {
    const line = hangs.filter((hang) => hang.dotted === dotted).map((hang) => hang.x);
    if (!line.length || (stem <= 0 && line.length < 2)) return;
    const left = Math.min(...line);
    const right = Math.max(...line);
    if (right - left < 0.5) return;
    items.push({
      type: "wire",
      x1: left,
      y1: stem,
      x2: right,
      y2: stem,
      dotted,
      weight: metrics.line,
    });
  };
  drawBar(false);
  drawBar(true);
  if (stem > 0) {
    items.push({
      type: "wire",
      x1: anchorX,
      y1: 0,
      x2: anchorX,
      y2: stem,
      dotted: hangs.every((hang) => hang.dotted),
      weight: metrics.line,
    });
  }
  return { w: totalW, h: maxH, anchorX, items };
}

function clusterBranches(branches) {
  const ordered = [];
  const used = new Set();
  branches.forEach((branch, index) => {
    if (used.has(index)) return;
    ordered.push(branch);
    used.add(index);
    branches.forEach((other, otherIndex) => {
      if (used.has(otherIndex)) return;
      if (other.deptKey && other.deptKey === branch.deptKey) {
        ordered.push(other);
        used.add(otherIndex);
      }
    });
  });
  return ordered;
}

function bubbleRowShift(org, key) {
  const reports = directReports(org, key);
  for (const department of departmentsOf(org, reports.workers)) {
    if (department.people.length === 1 && hasTwoManagers(org.people.get(department.people[0]))) return true;
  }
  return reports.managers.some((id) => hasTwoManagers(org.people.get(id)));
}

/** Move a branch down and stretch the tick that started at the top, so its apex stays put. */
function dropBranch(box, dy) {
  if (!(dy > 0)) return box;
  const items = translate(box.items, 0, dy).map((item) => {
    if (item.type !== "wire" || Math.abs(item.x1 - item.x2) > 0.8) return item;
    const top = Math.min(item.y1, item.y2);
    if (Math.abs(top - dy) > 0.6) return item;
    if (item.y1 <= item.y2) return { ...item, y1: 0 };
    return { ...item, y2: 0 };
  });
  return { ...box, h: box.h + dy, items };
}

function peelOuterTick(box) {
  const starters = box.items.filter(
    (item) => item.type === "wire" && Math.abs(item.x1 - item.x2) < 0.8 && Math.min(item.y1, item.y2) <= 0.5,
  );
  if (!starters.length) return { box, drop: 0, solid: false, dotted: false, weight: null };
  const drop = Math.max(...starters.map((item) => Math.abs(item.y2 - item.y1)));
  const remove = new Set(starters);
  return {
    box: {
      ...box,
      h: Math.max(0, box.h - drop),
      items: translate(box.items.filter((item) => !remove.has(item)), 0, -drop),
    },
    drop,
    // The head card's tick has to be as thick as the bar it lands on, or it pokes out below it.
    weight: Math.max(...starters.map((item) => item.weight || 0)) || null,
    solid: starters.some((item) => !item.dotted),
    dotted: starters.some((item) => item.dotted),
  };
}

function branchLine(org, parentKey, keys) {
  let solid = false;
  let dotted = false;
  for (const key of keys) {
    const kind = managerType(org.people.get(key), parentKey);
    if (kind === "Dotted") dotted = true;
    else solid = true;
  }
  if (solid && dotted) return "both";
  if (dotted) return "dotted";
  return "solid";
}

/**
 * Subordinates of `key`, not including `key`'s own card.
 * `visibility` maps a direct report to 1 (draw their team) or 0 (card only).
 */
export function layoutSubordinates(org, key, metrics, visibility, options = {}) {
  const asWorkers = Boolean(options.asWorkers);
  const shrunk = Boolean(options.shrunk);
  const stack = options.stack || new Set();
  if (stack.has(key)) return { w: 0, h: 0, anchorX: 0, items: [] };
  const reports = directReports(org, key);
  if (reports.all.length === 0) return { w: 0, h: 0, anchorX: 0, items: [] };
  if (shrunk) return layoutLeaves(org, reports.all, key, metrics, "4");

  const shiftDown = bubbleRowShift(org, key);
  const branches = [];
  if (asWorkers) {
    for (const department of departmentsOf(org, reports.all)) {
      let box = layoutLeaves(org, department.people, key, metrics, "auto");
      const alone = department.people.length === 1 && !hasTwoManagers(org.people.get(department.people[0]));
      if (shiftDown && alone) box = dropBranch(box, bubbleStack(metrics));
      branches.push({
        deptKey: department.key,
        line: branchLine(org, key, department.people),
        box,
      });
    }
  } else {
    for (const department of departmentsOf(org, reports.workers)) {
      let box = layoutLeaves(org, department.people, key, metrics, "auto");
      const alone = department.people.length === 1 && !hasTwoManagers(org.people.get(department.people[0]));
      if (shiftDown && alone) box = dropBranch(box, bubbleStack(metrics));
      branches.push({
        deptKey: department.key,
        line: branchLine(org, key, department.people),
        box,
      });
    }
    for (const managerKey of reports.managers) {
      const person = org.people.get(managerKey);
      const shownUnder = managerType(person, key) || "Solid";
      const expanded = shownUnder === "Solid" && visibility?.get(managerKey) === 1;
      let box;
      if (expanded) {
        const next = new Set(stack);
        next.add(key);
        box = layoutExpandedManager(org, managerKey, metrics, allVisible(org, managerKey), {
          stack: next,
          shownUnder,
        });
      } else {
        box = leafLayout(org, managerKey, metrics, shownUnder);
        if (isManager(org, managerKey)) markCollapsed(box);
      }
      if (shiftDown && !hasTwoManagers(person)) box = dropBranch(box, bubbleStack(metrics));
      branches.push({ deptKey: person.departmentKey, line: shownUnder === "Dotted" ? "dotted" : "solid", box });
    }
  }
  return placeBranches(clusterBranches(branches), metrics);
}

function layoutExpandedManager(org, key, metrics, visibility, options) {
  const shownUnder = options.shownUnder || "Solid";
  const subs = layoutSubordinates(org, key, metrics, visibility, { stack: options.stack });
  const person = org.people.get(key);
  const bubble = hasTwoManagers(person) ? bubbleStack(metrics) : 0;
  const tick = metrics.gapY + bubble;
  if (!subs.items.length) {
    const leaf = leafLayout(org, key, metrics, shownUnder);
    return leaf;
  }
  const cardX = subs.anchorX - metrics.w / 2;
  const left = Math.min(0, cardX);
  const right = Math.max(subs.w, cardX + metrics.w);
  const shift = -left;
  const cardY = tick;
  const subsY = cardY + metrics.h + (person.termination ? commentSpace(metrics) : 0);
  const items = [
    {
      type: "wire",
      x1: subs.anchorX + shift,
      y1: 0,
      x2: subs.anchorX + shift,
      y2: cardY,
      dotted: shownUnder === "Dotted",
      weight: metrics.line,
    },
    {
      type: "card",
      key,
      x: cardX + shift,
      y: cardY,
      w: metrics.w,
      h: metrics.h,
      size: metrics.size,
      managerType: shownUnder,
      collapsed: false,
      head: false,
    },
  ];
  if (hasTwoManagers(person)) {
    addBubble(items, org, person, key, subs.anchorX + shift, cardY, metrics, shownUnder, true);
  }
  if (person.termination) {
    items.push(commentItem(key, cardX + shift, cardY + metrics.h, metrics, person.termination));
  }
  items.push(...translate(subs.items, shift, subsY));
  return {
    w: right - left,
    h: subsY + subs.h,
    anchorX: subs.anchorX + shift,
    items,
  };
}

export function allVisible(org, key) {
  const flags = new Map();
  for (const id of directReports(org, key).all) flags.set(id, 1);
  return flags;
}

export function allHidden(org, key) {
  const flags = new Map();
  for (const id of directReports(org, key).all) flags.set(id, 0);
  return flags;
}

function horizontalBranchCount(org, key) {
  const reports = directReports(org, key);
  return reports.managers.length + departmentsOf(org, reports.workers).length;
}

function recursiveBranchCount(org, key, visibility) {
  const reports = directReports(org, key);
  let count = departmentsOf(org, reports.workers).length;
  for (const id of reports.managers) {
    const solid = managerType(org.people.get(id), key) === "Solid";
    if (solid && visibility?.get(id) === 1) count += recursiveBranchCount(org, id, allVisible(org, id));
    else count += 1;
  }
  return count;
}

function visibleManagers(org, visibility) {
  let count = 0;
  for (const [id, flag] of visibility) {
    if (flag === 1 && isManager(org, id)) count += 1;
  }
  return count;
}

function fits(box) {
  return box.w <= GUIDE.width + 0.5 && box.h <= GUIDE.height + 0.5;
}

function measure(org, key, metrics, visibility, shrunk) {
  return layoutSubordinates(org, key, metrics, visibility, { shrunk });
}

/**
 * Pick card size, spacing, and which manager teams stay expanded
 * so the tree fits one slide. Mirrors makeFullTreeAuto.
 */
export function chooseFit(org, rootKey) {
  const reports = directReports(org, rootKey);
  const person = org.people.get(rootKey);
  const heavy = (person?.stats.subordinates || 0) >= 100;
  const managers = reports.managers.filter((id) => managerType(org.people.get(id), rootKey) === "Solid");
  let visibility = heavy ? allHidden(org, rootKey) : allVisible(org, rootKey);
  let size = "Small";
  let gapX = 5;
  let gapY = 5;
  let gapExt = 5;

  const layoutAt = (nextSize, nextX, nextY, nextExt, flags, shrunk) =>
    measure(org, rootKey, makeMetrics(nextSize, nextX, nextExt, nextY), flags, shrunk);

  let box = layoutAt(size, gapX, gapY, gapExt, visibility, false);

  if (!heavy) {
    let guard = 0;
    while (!fits(box) && visibleManagers(org, visibility) > 0 && guard < 100) {
      guard += 1;
      let largest = null;
      let largestSize = -1;
      for (const id of managers) {
        if (visibility.get(id) !== 1) continue;
        const team = org.people.get(id).stats.subordinates;
        if (team > largestSize) {
          largestSize = team;
          largest = id;
        }
      }
      if (!largest) break;
      visibility.set(largest, 0);
      box = layoutAt(size, gapX, gapY, gapExt, visibility, false);
    }
  } else {
    let guard = 0;
    while (guard < managers.length) {
      guard += 1;
      let smallest = null;
      let smallestSize = Infinity;
      for (const id of managers) {
        if (visibility.get(id) !== 0) continue;
        const team = org.people.get(id).stats.subordinates;
        if (team < smallestSize) {
          smallestSize = team;
          smallest = id;
        }
      }
      if (!smallest) break;
      visibility.set(smallest, 1);
      const next = layoutAt(size, gapX, gapY, gapExt, visibility, false);
      if (!fits(next)) {
        visibility.set(smallest, 0);
        break;
      }
      box = next;
    }
  }

  const shrunk =
    visibleManagers(org, visibility) === 0 &&
    reports.managers.length > 0 &&
    horizontalBranchCount(org, rootKey) > MAX_BRANCHES;
  if (shrunk) box = layoutAt(size, gapX, gapY, gapExt, visibility, true);

  for (const candidate of ["Medium", "Large"]) {
    const next = layoutAt(candidate, gapX, gapY, gapExt, visibility, shrunk);
    if (!fits(next)) break;
    size = candidate;
    box = next;
  }

  const card = CARDS[size];
  const probed = layoutAt(size, gapX + 1 / 3, gapY + 1, gapExt, visibility, shrunk);
  const nDx = probed.w === box.w ? 0 : Math.abs(box.w - GUIDE.width) / (probed.w - box.w);
  const nDy = probed.h === box.h ? nDx : Math.abs(box.h - GUIDE.height) / (probed.h - box.h);
  const steps = Number.isFinite(nDx) && Number.isFinite(nDy) ? Math.round(Math.min(nDx, nDy)) - 1 : 0;
  gapX = Math.max(0, Math.min(card.w / 2, 5 + steps / 3));
  gapY = Math.max(0, Math.min(card.h / 3, 5 + steps));
  box = layoutAt(size, gapX, gapY, gapExt, visibility, shrunk);

  const branchCount = recursiveBranchCount(org, rootKey, visibility);
  if (branchCount > 1) {
    gapExt = Math.min(card.w, gapExt + Math.abs(box.w - GUIDE.width) / (branchCount - 1));
    box = layoutAt(size, gapX, gapY, gapExt, visibility, shrunk);
  }

  return { size, gap: gapY, gapX, gapY, gapExt, visibility, shrunk, box };
}

let shapeSerial = 1;
function sid(prefix) {
  shapeSerial += 1;
  return `${prefix}-${shapeSerial}`;
}

/** Shift the drawing so cards and lines stay inside the slide. Title, legend, and navigation stay put. */
export function containOnSlide(shapes) {
  const fixed = new Set(["title", "legend", "nav", "stats", "hiddenlist"]);
  let minX = Infinity;
  let maxX = -Infinity;
  for (const shape of shapes) {
    if (fixed.has(shape.type)) continue;
    if (shape.type === "wire") {
      minX = Math.min(minX, shape.x1, shape.x2);
      maxX = Math.max(maxX, shape.x1, shape.x2);
    } else if (Number.isFinite(shape.x)) {
      minX = Math.min(minX, shape.x);
      maxX = Math.max(maxX, shape.x + (shape.w || 0));
    }
  }
  if (!Number.isFinite(minX)) return;
  const margin = 8;
  const width = maxX - minX;
  let dx = 0;
  if (width <= SLIDE.width - margin * 2) {
    if (minX < margin) dx = margin - minX;
    if (maxX + dx > SLIDE.width - margin) dx -= maxX + dx - (SLIDE.width - margin);
  } else {
    dx = (SLIDE.width - width) / 2 - minX;
  }
  if (Math.abs(dx) < 0.4) return;
  for (const shape of shapes) {
    if (fixed.has(shape.type)) continue;
    if (shape.type === "wire") {
      shape.x1 += dx;
      shape.x2 += dx;
    } else if (Number.isFinite(shape.x)) shape.x += dx;
  }
}

function paint(items, dx, dy, extra = {}) {
  return translate(items, dx, dy).map((item) => ({ id: sid(item.type), z: item.type === "wire" ? 1 : 2, ...item, ...extra }));
}

function titleShape(org, rootKey) {
  const person = org.people.get(rootKey);
  return {
    id: sid("title"),
    type: "title",
    x: 18,
    y: 8,
    w: 360,
    h: 58,
    z: 5,
    key: rootKey,
    name: `${person.surname} ${person.name}`.trim(),
    department: person.department || "Organisation",
  };
}

function legendShape() {
  return { id: sid("legend"), type: "legend", x: 18, y: 66, w: 210, h: 58, z: 5 };
}

/** Full names of hidden people in the team under `rootKey` (solid lines), A to Z. */
export function hiddenNames(org, rootKey) {
  const names = [];
  const seen = new Set();
  const walk = (key) => {
    if (seen.has(key)) return;
    seen.add(key);
    for (const childKey of org.reports.get(key) || []) {
      const child = org.people.get(childKey);
      if (!child || child.solidKey !== key) continue;
      if (child.hidden) names.push(personFullName(child));
      walk(childKey);
    }
  };
  walk(rootKey);
  return names.sort((a, b) => a.localeCompare(b));
}

const HIDDEN_LIST = { x: 18, w: 924, bottom: SLIDE.height - 6, lineH: 10, charsPerLine: 200, maxLines: 4 };

/** Small grey line along the bottom of the slide naming everyone hidden from this chart. Null when nobody is. */
export function hiddenListShape(org, rootKey) {
  const names = hiddenNames(org, rootKey);
  if (!names.length) return null;
  const limit = HIDDEN_LIST.charsPerLine * HIDDEN_LIST.maxLines;
  const stats = org.people.get(rootKey)?.stats || {};
  const count = names.length;
  const mix = `${stats.hiddenOrdinary || 0} ordinary, ${stats.hiddenImportant || 0} non-ordinary, ${stats.hiddenIntern || 0} interns`;
  const head = `Hidden, not counted in the statistics – ${count} ${count === 1 ? "person" : "people"} (${mix}): `;
  let text = `${head}${names.join(", ")}`;
  if (text.length > limit) {
    let shown = names.length;
    while (shown > 1 && `${head}${names.slice(0, shown).join(", ")} and ${names.length - shown} more`.length > limit) shown -= 1;
    text = `${head}${names.slice(0, shown).join(", ")} and ${names.length - shown} more`;
  }
  const lines = Math.ceil(text.length / HIDDEN_LIST.charsPerLine);
  const h = lines * HIDDEN_LIST.lineH + 2;
  return {
    id: sid("hiddenlist"),
    type: "hiddenlist",
    x: HIDDEN_LIST.x,
    y: HIDDEN_LIST.bottom - h,
    w: HIDDEN_LIST.w,
    h,
    z: 5,
    key: rootKey,
    text,
  };
}

/** The box grows by one text line (12 pt) for every line past the usual five. */
export function statsBoxHeight(lines) {
  return 96 + 12 * Math.max(0, (lines?.length || 0) - 5);
}

function statsShape(org, rootKey) {
  const person = org.people.get(rootKey);
  const direct = directReports(org, rootKey).all.length;
  const lines = statisticsLines(person, direct);
  return {
    id: sid("stats"),
    type: "stats",
    x: 742,
    y: 10,
    w: 204,
    h: statsBoxHeight(lines),
    z: 5,
    key: rootKey,
    heading: "Department statistics",
    lines,
  };
}

const SHADE_GUIDE = { top: 140, right: 945, bottom: 505, left: 15 };

function boxOf(card) {
  return { top: card.y, left: card.x, right: card.x + card.w, bottom: card.y + card.h };
}

function pointInside(rect, x, y) {
  return x >= rect.left && x <= rect.right && y >= rect.top && y <= rect.bottom;
}

function cornersOf(rect) {
  return [
    [rect.left, rect.top],
    [rect.right, rect.top],
    [rect.right, rect.bottom],
    [rect.left, rect.bottom],
  ];
}

function rectsMeet(a, b) {
  return cornersOf(a).some(([x, y]) => pointInside(b, x, y)) || cornersOf(b).some(([x, y]) => pointInside(a, x, y));
}

function sideOf(shape, reference) {
  if (rectsMeet(shape, reference)) return "intersecting";
  let side = "safe";
  const bands = {
    top: { top: -1000, right: reference.right, bottom: reference.top, left: reference.left },
    right: { top: reference.top, right: 1000, bottom: reference.bottom, left: reference.right },
    bottom: { top: reference.bottom, right: reference.right, bottom: 1000, left: reference.left },
    left: { top: reference.top, right: reference.left, bottom: reference.bottom, left: -1000 },
  };
  for (const name of ["top", "right", "bottom", "left"]) {
    if (rectsMeet(shape, bands[name])) side = name;
  }
  return side;
}

/**
 * A sibling bar: long enough to leave the column, and not the short elbow
 * that only reaches the trunk. Those bars are the horizontal lines a shade
 * must not paint through.
 */
function connectorBars(wires, box, span) {
  const bars = [];
  for (const wire of wires || []) {
    if (Math.abs(wire.y1 - wire.y2) > 0.8) continue;
    const left = Math.min(wire.x1, wire.x2);
    const right = Math.max(wire.x1, wire.x2);
    const length = right - left;
    if (length < 48) continue;
    if (Math.min(right, span.right) - Math.max(left, span.left) < 4) continue;
    const overlap = Math.min(right, box.right) - Math.max(left, box.left);
    const overhang = overlap > 4 ? Math.max(box.left - left, right - box.right) : length;
    if (overhang < 20) continue;
    bars.push((wire.y1 + wire.y2) / 2);
  }
  return bars;
}

/** Pull the top or bottom back under a connector the cards do not already straddle. */
function clearOfConnectors(top, bottom, box, wires, span) {
  let ceiling = null;
  let floor = null;
  for (const y of connectorBars(wires, box, span)) {
    if (y < box.top - 0.4 && (ceiling === null || y > ceiling)) ceiling = y;
    if (y > box.bottom + 0.4 && (floor === null || y < floor)) floor = y;
  }
  // A shade that already sits under the bar stays put. One that runs through
  // the bar, or up to the slide guide, stops just below the line.
  if (ceiling != null && top < ceiling) {
    const padded = (box.top + ceiling) / 2 - 1;
    top = Math.min(box.top, Math.max(padded, ceiling + 2));
  }
  if (floor != null && bottom > floor) {
    const padded = (box.bottom + floor) / 2 + 1;
    bottom = Math.max(box.bottom, Math.min(padded, floor - 2));
  }
  return { top, bottom };
}

/** Department rectangles from makeDepartments: grow from the cards, then pull each side in from the slide guide. */
export function layoutDepartmentShades(org, cards, wires = []) {
  const members = [];
  for (const card of cards) {
    if (card.head || card.managerType === "Dotted" || card.y <= SHADE_GUIDE.top) continue;
    const person = org.people.get(card.key);
    if (!person) continue;
    members.push({ card, person, box: boxOf(card) });
  }
  const groups = new Map();
  for (const member of members) {
    // Cards with no department get no shade, but they still stop a neighbour's shade from running over them.
    if (!member.person.department) continue;
    const bucket = groups.get(member.person.departmentKey) || {
      key: member.person.departmentKey,
      name: member.person.department,
      boxes: [],
    };
    bucket.boxes.push(member.box);
    groups.set(member.person.departmentKey, bucket);
  }
  const shades = [];
  for (const bucket of groups.values()) {
    const shade = {
      top: Math.min(...bucket.boxes.map((box) => box.top)),
      left: Math.min(...bucket.boxes.map((box) => box.left)),
      right: Math.max(...bucket.boxes.map((box) => box.right)),
      bottom: Math.max(...bucket.boxes.map((box) => box.bottom)),
    };
    const stretched = { ...SHADE_GUIDE };
    for (const other of members) {
      if (other.person.departmentKey === bucket.key) continue; // same department: part of the shade, not a wall
      const side = sideOf(other.box, shade);
      if (side === "top" && other.box.bottom > stretched.top) stretched.top = other.box.bottom;
      if (side === "bottom" && other.box.top < stretched.bottom) stretched.bottom = other.box.top;
      if (side === "right" && other.box.left < stretched.right) stretched.right = other.box.left;
      if (side === "left" && other.box.right > stretched.left) stretched.left = other.box.right;
    }
    const right = stretched.right !== SHADE_GUIDE.right ? (shade.right + stretched.right) / 2 + 1 : stretched.right + 1;
    const left = stretched.left !== SHADE_GUIDE.left ? (shade.left + stretched.left) / 2 - 1 : stretched.left - 1;
    const cleared = clearOfConnectors(
      stretched.top !== SHADE_GUIDE.top ? (shade.top + stretched.top) / 2 - 1 : stretched.top - 1,
      stretched.bottom !== SHADE_GUIDE.bottom ? (shade.bottom + stretched.bottom) / 2 + 1 : stretched.bottom + 1,
      shade,
      wires,
      { left, right },
    );
    shades.push({ name: bucket.name, top: cleared.top, left, right, bottom: cleared.bottom });
  }
  shades.sort((a, b) => (b.right - b.left) * (b.bottom - b.top) - (a.right - a.left) * (a.bottom - a.top));
  return shades.map((shade, index) => ({
    id: sid("shade"),
    type: "shade",
    x: shade.left,
    y: shade.top,
    w: Math.max(8, shade.right - shade.left),
    h: Math.max(24, shade.bottom - shade.top),
    z: 0,
    text: shade.name,
    fill: SHADE_COLORS[index % SHADE_COLORS.length],
  }));
}

function shadeShapes(org, cards, wires) {
  return layoutDepartmentShades(org, cards, wires);
}

function headCard(org, rootKey, peeled) {
  const metrics = CARDS.Large;
  const x = SLIDE.width / 2 - metrics.w / 2;
  const y = TREE_TOP - metrics.h;
  const items = [
    {
      type: "card",
      key: rootKey,
      x,
      y,
      w: metrics.w,
      h: metrics.h,
      managerType: "Solid",
      collapsed: false,
      head: true,
      size: "Large",
    },
  ];
  const person = org.people.get(rootKey);
  const comment = person?.termination ? commentSpace({ size: "Large" }) : 0;
  if (comment) {
    items.push({
      type: "comment",
      key: rootKey,
      x,
      y: y + metrics.h + COMMENT_GAP,
      w: metrics.w,
      h: commentHeight({ size: "Large" }),
      z: 4,
      text: `Works until ${person.termination}`,
    });
  }
  const tickTop = y + metrics.h + comment;
  const weight = peeled?.weight || metrics.line;
  if (peeled?.solid) {
    items.push({ type: "wire", x1: SLIDE.width / 2, y1: tickTop, x2: SLIDE.width / 2, y2: tickTop + 15, dotted: false, weight });
  }
  if (peeled?.dotted) {
    items.push({ type: "wire", x1: SLIDE.width / 2, y1: tickTop, x2: SLIDE.width / 2, y2: tickTop + 15, dotted: true, weight });
  }
  return paint(items, 0, 0);
}

/**
 * Build one or many slides.
 * mode: "clever" (single slide) or "infinity" (a slide for every folded team).
 */
export function buildSlides(org, rootKey, mode = "clever") {
  shapeSerial = 1;
  const slides = [];
  const slideOf = new Map();

  function build(key, parentKey) {
    const fit = chooseFit(org, key);
    const metrics = makeMetrics(fit.size, fit.gapX ?? fit.gap, fit.gapExt, fit.gapY ?? fit.gap);
    const grown = layoutSubordinates(org, key, metrics, fit.visibility, { shrunk: fit.shrunk });
    const peeled = peelOuterTick(grown);
    const tree = peeled.box;
    const dx = SLIDE.width / 2 - tree.anchorX;
    const headComment = org.people.get(key)?.termination ? commentSpace({ size: "Large" }) : 0;
    const body = paint(tree.items, dx, TREE_TOP + headComment + 15);
    const shapes = [
      titleShape(org, key),
      legendShape(),
      statsShape(org, key),
      ...[hiddenListShape(org, key)].filter(Boolean),
      ...shadeShapes(
        org,
        body.filter((item) => item.type === "card"),
        body.filter((item) => item.type === "wire"),
      ),
      ...headCard(org, key, peeled),
      ...body,
    ];
    const slide = {
      id: sid("slide"),
      rootKey: key,
      parentKey,
      fit: {
        size: fit.size,
        gap: fit.gap,
        gapExt: fit.gapExt,
        shrunk: fit.shrunk,
        visibility: [...fit.visibility.entries()],
      },
      shapes,
    };
    slides.push(slide);
    slideOf.set(key, slides.length);
    if (mode === "infinity") {
      for (const [child, flag] of fit.visibility) {
        if (flag !== 0) continue;
        if (!isManager(org, child)) continue;
        if (managerType(org.people.get(child), key) !== "Solid") continue;
        build(child, key);
      }
    }
    return slide;
  }

  build(rootKey, null);

  if (mode === "infinity") {
    // Hidden people never get a card, not even a loose one.
    const required = new Set([...org.people.keys()].filter((key) => !org.people.get(key).hidden));
    const drawnKeys = () => {
      const placed = new Set();
      for (const slide of slides) {
        for (const shape of slide.shapes) if (shape.type === "card" && shape.key) placed.add(shape.key);
      }
      return placed;
    };
    let guard = 0;
    while (guard < required.size + 1) {
      guard += 1;
      const missing = [...required].filter((key) => !drawnKeys().has(key));
      if (!missing.length) break;
      const person = org.people.get(missing[0]);
      const manager = person?.solidKey;
      if (manager && !slides.some((slide) => slide.rootKey === manager)) {
        build(manager, org.people.get(manager)?.solidKey || rootKey);
        continue;
      }
      const slide = slides.find((item) => item.rootKey === manager) || slides[0];
      if (slide.shapes.some((shape) => shape.type === "card" && shape.key === missing[0])) break;
      const metrics = makeMetrics(slide.fit?.size || "Small", slide.fit?.gap || 8);
      const cards = slide.shapes.filter((shape) => shape.type === "card");
      const y = cards.length ? Math.max(...cards.map((card) => card.y + card.h)) + 8 : TREE_TOP;
      const x = cards.length ? Math.min(...cards.map((card) => card.x)) : 48;
      slide.shapes.push(...layoutSingleCard(org, missing[0], metrics, { x, y }));
    }
  }

  for (const slide of slides) {
    containOnSlide(slide.shapes);
    freezeCards(org, slide.shapes);
    slide.rootLabel = displayName(org.people.get(slide.rootKey));
  }

  for (const slide of slides) {
    for (const shape of [...slide.shapes]) {
      if (shape.type !== "card" || !shape.collapsed) continue;
      const target = slideOf.get(shape.key);
      if (!target) continue;
      shape.linkSlide = target;
      slide.shapes.push({
        id: sid("link"),
        type: "link",
        x: shape.x,
        y: slideLinkTop(shape, slide.shapes),
        w: shape.w,
        h: 11,
        z: 4,
        text: `See slide ${target}`,
        target,
        key: shape.key,
      });
    }
    if (slide.parentKey) {
      const back = slideOf.get(slide.parentKey);
      const parent = org.people.get(slide.parentKey);
      slide.shapes.push({
        id: sid("nav"),
        type: "nav",
        x: 460,
        y: 14,
        w: 280,
        h: 18,
        z: 4,
        text: `Up to ${parent.surname} ${moderated(parent.name)} (slide ${back})`,
        target: back,
      });
    }
  }

  return slides;
}

/** A single card, used by the manual “Employee” and “Manager” commands. */
export function layoutSingleCard(org, key, metrics, { head = false, x = 80, y = 120 } = {}) {
  const person = org.people.get(key);
  const items = [
    {
      id: sid("card"),
      type: "card",
      key,
      x,
      y,
      w: metrics.w,
      h: metrics.h,
      z: 2,
      managerType: "Solid",
      collapsed: false,
      head,
      size: metrics.size,
    },
  ];
  if (hasTwoManagers(person)) {
    const gap = bubbleGap(metrics);
    items.push({
      id: sid("bubble"),
      type: "bubble",
      key,
      x: x + metrics.w / 2 - metrics.bubbleW / 2,
      y: y - gap - metrics.bubbleH,
      w: metrics.bubbleW,
      h: metrics.bubbleH,
      z: 4,
      dotted: true,
      text: org.people.get(person.dottedKey)?.surname || "",
    });
  }
  return items;
}

/**
 * Manual branch under an existing card.
 * mode: "auto" | "1" | "2" | "4" | "tree"
 */
export function layoutManualBranch(org, key, metrics, mode) {
  if (mode === "tree") {
    return layoutSubordinates(org, key, metrics, allVisible(org, key), {});
  }
  if (mode === "auto") {
    return layoutSubordinates(org, key, metrics, allVisible(org, key), { asWorkers: true });
  }
  const people = directReports(org, key).all;
  return layoutLeaves(org, people, key, metrics, mode);
}
