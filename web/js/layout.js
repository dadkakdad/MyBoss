/**
 * Org-chart geometry.
 *
 * The PowerPoint tool measures a tree, hides the largest teams until the
 * drawing fits a 920×350 pt area, then grows card size and spacing. The same
 * decisions are made here. Coordinates are PowerPoint points on a 960×540 slide.
 */

import {
  comparePeople,
  departmentsOf,
  directReports,
  hasTwoManagers,
  isManager,
  managerType,
  moderated,
  statisticsLines,
} from "./org.js";

export const SLIDE = { width: 960, height: 540 };
export const GUIDE = { width: 920, height: 400 };
export const MAX_PEOPLE_FOR_2D = 6;
export const MAX_PEOPLE_FOR_4D = 14;
export const MAX_BRANCHES = 7;
export const TREE_TOP = 132;

export const CARD_SIZES = ["ExtraSmall", "Small", "Medium", "Large"];

export const CARDS = {
  ExtraSmall: { w: 70.9, h: 26.2, name: 7, title: 4, bubbleW: 52.2, bubbleH: 7.5, commentH: 6.2, line: 2 },
  Small: { w: 70.9, h: 35.5, name: 7, title: 4, bubbleW: 52.2, bubbleH: 7.5, commentH: 6.2, line: 2 },
  Medium: { w: 87.9, h: 42.3, name: 9, title: 5, bubbleW: 65.2, bubbleH: 9.4, commentH: 7.7, line: 3 },
  Large: { w: 108.5, h: 54.5, name: 11, title: 6, bubbleW: 81.1, bubbleH: 11.6, commentH: 8.8, line: 4 },
};

export const SHADE_COLORS = ["#DEEBF7", "#BDD7EE", "#D6DCE8"];

export function makeMetrics(size = "Medium", gap = 10, gapExt = gap) {
  const card = CARDS[size] || CARDS.Medium;
  return { ...card, size, gapX: gap, gapY: gap, gapExt };
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

function otherManagerSurname(org, person, shownUnder) {
  const key = shownUnder === "Solid" ? person.dottedKey : person.solidKey;
  const other = org.people.get(key);
  return other ? other.surname : "";
}

function leafLayout(org, key, metrics, shownUnder) {
  const person = org.people.get(key);
  const bubble = hasTwoManagers(person) ? metrics.bubbleH : 0;
  const tick = metrics.gapY + bubble;
  const anchorX = metrics.w / 2;
  const items = [
    {
      type: "wire",
      x1: anchorX,
      y1: 0,
      x2: anchorX,
      y2: tick,
      dotted: shownUnder === "Dotted",
      weight: metrics.line,
    },
  ];
  if (hasTwoManagers(person)) {
    const other = shownUnder === "Dotted" ? "Solid" : "Dotted";
    items.push({
      type: "wire",
      x1: anchorX + 5,
      y1: 0,
      x2: anchorX + 5,
      y2: metrics.gapY / 2 + metrics.bubbleH / 2,
      dotted: other === "Dotted",
      weight: metrics.line,
    });
    items.push({
      type: "bubble",
      key,
      x: anchorX - metrics.bubbleW / 2,
      y: tick - metrics.bubbleH,
      w: metrics.bubbleW,
      h: metrics.bubbleH,
      dotted: other === "Dotted",
      text: otherManagerSurname(org, person, shownUnder),
    });
  }
  items.push({
    type: "card",
    key,
    x: 0,
    y: tick,
    w: metrics.w,
    h: metrics.h,
    size: metrics.size,
    managerType: shownUnder,
    collapsed: false,
    head: false,
  });
  return { w: metrics.w, h: tick + metrics.h, anchorX, items };
}

/** A vertical stack, one card wide, centered on the connector. Matches the columns in the demo. */
function columnLayout(org, people, parentKey, metrics) {
  const stackGap = Math.max(10, metrics.gapY);
  const anchorX = metrics.w / 2;
  const items = [];
  let y = 0;
  let previousBottom = 0;
  for (const key of people) {
    const person = org.people.get(key);
    const shownUnder = managerType(person, parentKey) || "Solid";
    const bubble = hasTwoManagers(person) ? metrics.bubbleH + 1 : 0;
    if (y > 0) y += stackGap;
    y += bubble;
    const cardY = y;
    if (bubble) {
      items.push({
        type: "bubble",
        key,
        x: anchorX - metrics.bubbleW / 2,
        y: cardY - metrics.bubbleH - 1,
        w: metrics.bubbleW,
        h: metrics.bubbleH,
        dotted: shownUnder !== "Dotted",
        text: otherManagerSurname(org, person, shownUnder),
      });
    }
    items.push({
      type: "wire",
      x1: anchorX,
      y1: previousBottom,
      x2: anchorX,
      y2: cardY,
      dotted: shownUnder === "Dotted",
      weight: metrics.line,
    });
    items.push({
      type: "card",
      key,
      x: 0,
      y: cardY,
      w: metrics.w,
      h: metrics.h,
      size: metrics.size,
      managerType: shownUnder,
      collapsed: false,
      head: false,
    });
    y += metrics.h;
    previousBottom = y;
  }
  return { w: metrics.w, h: y, anchorX, items };
}

function rowsLayout(org, people, parentKey, metrics, perRow) {
  const pitch = metrics.gapX + metrics.w;
  const shifts = shiftMap(perRow);
  const items = [];
  let y = 0;
  let minLeft = Infinity;
  let maxRight = -Infinity;
  for (let start = 0; start < people.length; start += perRow) {
    const row = people.slice(start, start + perRow);
    let rowH = 0;
    for (let i = 0; i < row.length; i += 1) {
      const shownUnder = managerType(org.people.get(row[i]), parentKey) || "Solid";
      const leaf = leafLayout(org, row[i], metrics, shownUnder);
      const center = shifts[i] * pitch;
      const ox = center - metrics.w / 2;
      items.push(...translate(leaf.items, ox, y));
      minLeft = Math.min(minLeft, ox);
      maxRight = Math.max(maxRight, ox + leaf.w);
      rowH = Math.max(rowH, leaf.h);
    }
    y += rowH;
  }
  const shift = Number.isFinite(minLeft) ? -minLeft : 0;
  return {
    w: Number.isFinite(minLeft) ? maxRight - minLeft : 0,
    h: y,
    anchorX: shift,
    items: translate(items, shift, 0),
  };
}

export function layoutLeaves(org, people, parentKey, metrics, mode = "auto") {
  const count = people.length;
  if (count === 0) return { w: 0, h: 0, anchorX: 0, items: [] };
  if (mode === "1" || (mode === "auto" && count <= MAX_PEOPLE_FOR_2D && count !== 1)) {
    return columnLayout(org, people, parentKey, metrics);
  }
  if (mode === "2" || (mode === "auto" && count <= MAX_PEOPLE_FOR_4D)) {
    return rowsLayout(org, people, parentKey, metrics, 2);
  }
  if (mode === "4" || mode === "auto") {
    if (count === 1 && mode === "auto") {
      const shownUnder = managerType(org.people.get(people[0]), parentKey) || "Solid";
      return leafLayout(org, people[0], metrics, shownUnder);
    }
    return rowsLayout(org, people, parentKey, metrics, mode === "2" ? 2 : 4);
  }
  const shownUnder = managerType(org.people.get(people[0]), parentKey) || "Solid";
  return leafLayout(org, people[0], metrics, shownUnder);
}

function markCollapsed(box) {
  box.items = box.items.map((item) => (item.type === "card" ? { ...item, collapsed: true } : item));
  return box;
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
  const solidAnchors = [];
  const dottedAnchors = [];
  let maxH = 0;
  for (const entry of placed) {
    const anchor = entry.x + entry.branch.box.anchorX;
    if (entry.branch.line !== "dotted") solidAnchors.push(anchor);
    if (entry.branch.line !== "solid") dottedAnchors.push(anchor);
    items.push(...translate(entry.branch.box.items, entry.x, stem));
    maxH = Math.max(maxH, stem + entry.branch.box.h);
  }
  const anchors = solidAnchors.length ? solidAnchors : dottedAnchors;
  const anchorX = anchors.length ? (Math.min(...anchors) + Math.max(...anchors)) / 2 : totalW / 2;
  const bar = (lineAnchors, dotted) => {
    if (!lineAnchors.length || stem <= 0 && lineAnchors.length < 2) return;
    const xs = lineAnchors.length === 1 ? [lineAnchors[0], anchorX] : lineAnchors;
    const left = Math.min(...xs);
    const right = Math.max(...xs);
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
  bar(solidAnchors, false);
  bar(dottedAnchors, true);
  if (stem > 0) {
    items.push({
      type: "wire",
      x1: anchorX,
      y1: 0,
      x2: anchorX,
      y2: stem,
      dotted: false,
      weight: metrics.line,
    });
    if (dottedAnchors.length && solidAnchors.length) {
      items.push({
        type: "wire",
        x1: anchorX + 3,
        y1: 0,
        x2: anchorX + 3,
        y2: stem,
        dotted: true,
        weight: metrics.line,
      });
    }
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

  const branches = [];
  if (asWorkers) {
    for (const department of departmentsOf(org, reports.all)) {
      branches.push({
        deptKey: department.key,
        line: branchLine(org, key, department.people),
        box: layoutLeaves(org, department.people, key, metrics, "auto"),
      });
    }
  } else {
    for (const department of departmentsOf(org, reports.workers)) {
      branches.push({
        deptKey: department.key,
        line: branchLine(org, key, department.people),
        box: layoutLeaves(org, department.people, key, metrics, "auto"),
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
      branches.push({ deptKey: person.departmentKey, line: shownUnder === "Dotted" ? "dotted" : "solid", box });
    }
  }
  return placeBranches(clusterBranches(branches), metrics);
}

function layoutExpandedManager(org, key, metrics, visibility, options) {
  const shownUnder = options.shownUnder || "Solid";
  const subs = layoutSubordinates(org, key, metrics, visibility, { stack: options.stack });
  const person = org.people.get(key);
  const bubble = hasTwoManagers(person) ? metrics.bubbleH : 0;
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
  const subsY = cardY + metrics.h;
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
    items.push({
      type: "bubble",
      key,
      x: subs.anchorX + shift - metrics.bubbleW / 2,
      y: cardY - metrics.bubbleH,
      w: metrics.bubbleW,
      h: metrics.bubbleH,
      dotted: shownUnder !== "Dotted",
      text: otherManagerSurname(org, person, shownUnder),
    });
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
  let gap = 5;
  let gapExt = 5;

  const layoutAt = (nextSize, nextGap, nextExt, flags, shrunk) =>
    measure(org, rootKey, makeMetrics(nextSize, nextGap, nextExt), flags, shrunk);

  let box = layoutAt(size, gap, gapExt, visibility, false);

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
      box = layoutAt(size, gap, gapExt, visibility, false);
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
      const next = layoutAt(size, gap, gapExt, visibility, false);
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
  if (shrunk) box = layoutAt(size, gap, gapExt, visibility, true);

  for (const candidate of ["Medium", "Large"]) {
    const next = layoutAt(candidate, gap, gapExt, visibility, shrunk);
    if (!fits(next)) break;
    size = candidate;
    box = next;
  }
  if (!fits(box)) {
    const tiny = layoutAt("ExtraSmall", gap, gapExt, visibility, shrunk);
    if (tiny.w < box.w || tiny.h < box.h) {
      size = "ExtraSmall";
      box = tiny;
    }
  }

  const card = CARDS[size];
  const gapLimit = Math.max(5, Math.min(card.w / 2, card.h / 3));
  for (let nextGap = gap + 1; nextGap <= gapLimit; nextGap += 1) {
    const next = layoutAt(size, nextGap, gapExt, visibility, shrunk);
    if (!fits(next)) break;
    gap = nextGap;
    box = next;
  }

  const extLimit = card.w;
  const branchCount = Math.max(1, horizontalBranchCount(org, rootKey));
  if (branchCount > 1 && box.w < GUIDE.width) {
    const room = (GUIDE.width - box.w) / (branchCount - 1);
    const grown = Math.min(extLimit, gapExt + room);
    const next = layoutAt(size, gap, grown, visibility, shrunk);
    if (fits(next)) {
      gapExt = grown;
      box = next;
    }
  }

  return { size, gap, gapExt, visibility, shrunk, box };
}

let shapeSerial = 1;
function sid(prefix) {
  shapeSerial += 1;
  return `${prefix}-${shapeSerial}`;
}

/** Shift the drawing so cards and lines stay inside the slide. Title, legend, and navigation stay put. */
export function containOnSlide(shapes) {
  const fixed = new Set(["title", "legend", "nav", "stats"]);
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

function statsShape(org, rootKey) {
  const person = org.people.get(rootKey);
  const direct = directReports(org, rootKey).all.length;
  return {
    id: sid("stats"),
    type: "stats",
    x: 742,
    y: 10,
    w: 204,
    h: 96,
    z: 5,
    key: rootKey,
    heading: "Department statistics",
    lines: statisticsLines(person, direct),
  };
}

function shadeShapes(org, cards) {
  const groups = new Map();
  for (const card of cards) {
    if (card.head) continue;
    const person = org.people.get(card.key);
    if (!person?.department) continue;
    const bucket = groups.get(person.departmentKey) || {
      name: person.department,
      x: Infinity,
      y: Infinity,
      r: -Infinity,
      b: -Infinity,
    };
    bucket.x = Math.min(bucket.x, card.x);
    bucket.y = Math.min(bucket.y, card.y);
    bucket.r = Math.max(bucket.r, card.x + card.w);
    bucket.b = Math.max(bucket.b, card.y + card.h);
    groups.set(person.departmentKey, bucket);
  }
  const shades = [];
  let index = 0;
  const lowest = Math.max(...[...groups.values()].map((bucket) => bucket.b));
  for (const bucket of groups.values()) {
    if (bucket.b - bucket.y < 8) continue;
    const pad = 8;
    const top = bucket.y;
    const bottom = Math.min(SLIDE.height - 6, Math.max(bucket.b, lowest) + 22);
    shades.push({
      id: sid("shade"),
      type: "shade",
      x: bucket.x - pad,
      y: top,
      w: bucket.r - bucket.x + pad * 2,
      h: Math.max(24, bottom - top),
      z: 0,
      text: bucket.name,
      fill: SHADE_COLORS[index % SHADE_COLORS.length],
    });
    index += 1;
  }
  return shades;
}

function headCard(org, rootKey, tree, bodySize) {
  const metrics = CARDS[bodySize] || CARDS.Large;
  const x = SLIDE.width / 2 - metrics.w / 2;
  const y = TREE_TOP - metrics.h - 16;
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
      size: bodySize || "Large",
    },
  ];
  if (tree.items.length) {
    items.push({
      type: "wire",
      x1: SLIDE.width / 2,
      y1: y + metrics.h,
      x2: SLIDE.width / 2,
      y2: TREE_TOP,
      dotted: false,
      weight: 4,
    });
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
    const metrics = makeMetrics(fit.size, fit.gap, fit.gapExt);
    const tree = layoutSubordinates(org, key, metrics, fit.visibility, { shrunk: fit.shrunk });
    const dx = SLIDE.width / 2 - tree.anchorX;
    const body = paint(tree.items, dx, TREE_TOP);
    const shapes = [
      titleShape(org, key),
      legendShape(),
      statsShape(org, key),
      ...shadeShapes(org, body.filter((item) => item.type === "card")),
      ...headCard(org, key, tree, fit.size),
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

  for (const slide of slides) containOnSlide(slide.shapes);

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
        y: shape.y + shape.h - 11,
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
    items.push({
      id: sid("bubble"),
      type: "bubble",
      key,
      x: x + metrics.w / 2 - metrics.bubbleW / 2,
      y: y - metrics.bubbleH - 2,
      w: metrics.bubbleW,
      h: metrics.bubbleH,
      z: 3,
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
