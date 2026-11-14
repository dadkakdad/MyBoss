/**
 * Slide list, selection geometry, and cost-center stamps.
 * These stay free of the DOM so the ribbon and the filmstrip can share them.
 */

import { CARDS, SLIDE } from "./layout.js";
import { moderated, shapePerson } from "./org.js";

export function shapeBox(shape) {
  if (!shape) return null;
  if (shape.type === "wire") {
    const x = Math.min(shape.x1, shape.x2);
    const y = Math.min(shape.y1, shape.y2);
    return {
      x,
      y,
      w: Math.max(Math.abs(shape.x2 - shape.x1), 1),
      h: Math.max(Math.abs(shape.y2 - shape.y1), 1),
    };
  }
  if (!Number.isFinite(shape.x) || !Number.isFinite(shape.w)) return null;
  return { x: shape.x, y: shape.y, w: shape.w, h: shape.h || 0 };
}

export function boxesIntersect(a, b) {
  return a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
}

/** Union of the slide page and every shape, with a little air past the page. */
export function contentFrame(shapes) {
  let minX = 0;
  let minY = 0;
  let maxX = SLIDE.width;
  let maxY = SLIDE.height;
  for (const shape of shapes || []) {
    const box = shapeBox(shape);
    if (!box) continue;
    minX = Math.min(minX, box.x);
    minY = Math.min(minY, box.y);
    maxX = Math.max(maxX, box.x + box.w);
    maxY = Math.max(maxY, box.y + box.h);
  }
  const pad = 16;
  if (minX < 0) minX -= pad;
  if (minY < 0) minY -= pad;
  if (maxX > SLIDE.width) maxX += pad;
  if (maxY > SLIDE.height) maxY += pad;
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
}

export function selectionBounds(shapes) {
  const boxes = (shapes || []).map((shape) => shapeBox(shape)).filter(Boolean);
  if (!boxes.length) return null;
  const x = Math.min(...boxes.map((box) => box.x));
  const y = Math.min(...boxes.map((box) => box.y));
  const right = Math.max(...boxes.map((box) => box.x + box.w));
  const bottom = Math.max(...boxes.map((box) => box.y + box.h));
  return { x, y, w: Math.max(right - x, 1), h: Math.max(bottom - y, 1) };
}

export function scaleShapes(shapes, factor, origin) {
  const next = Math.max(0.2, Math.min(4, factor));
  for (const shape of shapes) {
    if (shape.type === "wire") {
      shape.x1 = origin.x + (shape.x1 - origin.x) * next;
      shape.y1 = origin.y + (shape.y1 - origin.y) * next;
      shape.x2 = origin.x + (shape.x2 - origin.x) * next;
      shape.y2 = origin.y + (shape.y2 - origin.y) * next;
      if (Number.isFinite(shape.weight)) shape.weight *= next;
    } else if (Number.isFinite(shape.x)) {
      shape.x = origin.x + (shape.x - origin.x) * next;
      shape.y = origin.y + (shape.y - origin.y) * next;
      shape.w *= next;
      shape.h *= next;
    }
  }
  return next;
}

/** Pull one side of a box. `edge` is a combination of n, s, e, w. The opposite side stays put. */
export function resizeShapes(shapes, edge, dx, dy, minimum = { w: 16, h: 24 }) {
  const minW = minimum.w ?? 16;
  const minH = minimum.h ?? 24;
  const side = String(edge || "");
  for (const shape of shapes) {
    if (!Number.isFinite(shape.x) || !Number.isFinite(shape.w) || !Number.isFinite(shape.h)) continue;
    let { x, y, w, h } = shape;
    if (side.includes("w")) {
      const right = x + w;
      x = Math.min(x + dx, right - minW);
      w = right - x;
    }
    if (side.includes("e")) w = Math.max(minW, w + dx);
    if (side.includes("n")) {
      const bottom = y + h;
      y = Math.min(y + dy, bottom - minH);
      h = bottom - y;
    }
    if (side.includes("s")) h = Math.max(minH, h + dy);
    shape.x = x;
    shape.y = y;
    shape.w = w;
    shape.h = h;
  }
}

export function rewriteSlideNumber(text, number) {
  if (typeof text !== "string") return text;
  return text.replace(/slide \d+/gi, `slide ${number}`);
}

function bumpField(shape, field, delta, above) {
  if (!Number.isFinite(shape[field]) || shape[field] <= above) return;
  shape[field] += delta;
  if (field === "target") shape.text = rewriteSlideNumber(shape.text, shape[field]);
}

/** Add `delta` to slide numbers strictly greater than `above` (1-based). */
export function offsetSlideTargets(slides, delta, above = 0) {
  if (!delta) return;
  for (const slide of slides) {
    for (const shape of slide.shapes) {
      bumpField(shape, "target", delta, above);
      bumpField(shape, "linkSlide", delta, above);
    }
  }
}

/**
 * New charts go in front of whatever is already there.
 * A lone empty starter slide is the untouched canvas, so it does not linger.
 */
export function prependCharts(existing, fresh) {
  const placeholder = existing.length === 1 && !existing[0].title && existing[0].shapes.length === 0;
  const kept = placeholder ? [] : existing;
  offsetSlideTargets(kept, fresh.length, 0);
  return [...fresh, ...kept];
}

function dropRemovedPointer(shape, field, removed) {
  if (!Number.isFinite(shape[field])) return;
  if (shape[field] === removed) {
    delete shape[field];
    return;
  }
  if (shape[field] > removed) {
    shape[field] -= 1;
    if (field === "target") shape.text = rewriteSlideNumber(shape.text, shape[field]);
  }
}

/** Drop one slide and keep "See slide N" pointed at the slides that remain. */
export function removeSlide(slides, index) {
  const removed = index + 1;
  const next = slides.filter((_, i) => i !== index);
  for (const slide of next) {
    slide.shapes = slide.shapes.filter((shape) => shape.target !== removed);
    for (const shape of slide.shapes) {
      dropRemovedPointer(shape, "target", removed);
      dropRemovedPointer(shape, "linkSlide", removed);
    }
  }
  return next;
}

/** VBA shows the cost center, or "?" when it is blank or zero. */
export function costCenterText(value) {
  const text = String(value ?? "").trim();
  if (!text || text === "0") return "?";
  return text;
}

/**
 * Small green stamp flush in the card's bottom-right corner.
 */
export const CC_STAMP = { w: 26, h: 11 };

/**
 * The stamp grows leftwards with its text, up to the card's width. Past that the type shrinks
 * (never below 4.5 pt) so a long center like "СТАЖЕР" is not cut off.
 */
export function stampSize(text, cardWidth) {
  const base = CC_STAMP.h * 0.64;
  const wide = 0.72; // average bold capital, in em
  const need = text.length * base * wide + 6;
  const w = Math.max(CC_STAMP.w, Math.min(need, cardWidth));
  const fs = need <= cardWidth ? base : Math.max(4.5, (base * (cardWidth - 6)) / (text.length * base * wide));
  return { w, h: CC_STAMP.h, fs };
}

export function showCostCenterStamps(org, slides, idFor) {
  let count = 0;
  for (const slide of slides) {
    const cards = slide.shapes.filter((shape) => shape.type === "card");
    slide.shapes = slide.shapes.filter((shape) => shape.type !== "cc");
    for (const card of cards) {
      const person = shapePerson(org, card);
      const text = costCenterText(person?.costCenter);
      const size = stampSize(text, card.w);
      slide.shapes.push({
        id: idFor("cc"),
        type: "cc",
        x: card.x + card.w - size.w,
        y: card.y + card.h - size.h,
        w: size.w,
        h: size.h,
        fs: size.fs,
        z: 6,
        text,
        key: card.key || "",
      });
      count += 1;
    }
  }
  return count;
}

export function hideCostCenterStamps(slides) {
  for (const slide of slides) slide.shapes = slide.shapes.filter((shape) => shape.type !== "cc");
}

/** Plain text of one shape, or "" when it has none (connectors, frames). */
export function shapeText(org, shape) {
  if (!shape) return "";
  if (shape.type === "card") {
    const person = shapePerson(org, shape);
    if (!person) return (shape.lines || []).filter(Boolean).join(" ");
    const name = person.tbh ? "To be hired" : `${moderated(person.surname)} ${moderated(person.name)}`.trim();
    return [name, moderated(person.job), person.location].filter(Boolean).join(" — ");
  }
  if (shape.type === "title") return [shape.name, shape.department].filter(Boolean).join("\n");
  if (shape.type === "stats") return [shape.heading, ...(shape.lines || [])].filter(Boolean).join("\n");
  if (shape.type === "legend") return "Indicators\nOrdinary employee\nImportant employee\nIntern";
  return typeof shape.text === "string" ? shape.text : "";
}

/** Text of the given shapes in reading order: top to bottom, then left to right. */
export function shapesText(org, shapes) {
  return [...shapes]
    .filter((shape) => shapeBox(shape))
    .sort((a, b) => shapeBox(a).y - shapeBox(b).y || shapeBox(a).x - shapeBox(b).x)
    .map((shape) => shapeText(org, shape))
    .filter(Boolean)
    .join("\n");
}

const mdCell = (value) => String(value ?? "").replace(/\|/g, "\\|").replace(/\s+/g, " ").trim() || "—";

function cardNotes(person, card) {
  const notes = [];
  if (card.managerType === "Dotted") notes.push("dotted line");
  if (person.tbh) notes.push("vacancy");
  if (person.important) notes.push("key employee");
  if (person.intern) notes.push("intern");
  if (person.newcomer) notes.push(person.newcomerDate ? `newcomer, joins ${person.newcomerDate}` : "newcomer");
  if (person.termination) notes.push(`works until ${person.termination}`);
  if (card.collapsed && card.linkSlide) notes.push(`team on slide ${card.linkSlide}`);
  return notes;
}

/** Sticky-note colors. The first one is the orange used for a new note. */
export const NOTE_COLORS = ["#E8891A", "#F6C445", "#F29CA3", "#8FCB9B", "#E7B15A"];

/** The note uses the same type size as the name on its card, including after that card is scaled. */
export function noteTypeSize(card) {
  const spec = CARDS[card?.size] || CARDS.Medium;
  const scale = spec.h && Number(card?.h) ? Number(card.h) / spec.h : 1;
  return spec.name * scale;
}

/**
 * A note sits on the card's lower edge and hangs to the right, like a sticky.
 * Another note on the same card stacks under the previous one.
 * The rectangle is one line of the card's own name size.
 */
export function noteBox(card, index = 0) {
  const cardW = Number(card?.w) || 0;
  const cardH = Number(card?.h) || 0;
  const fs = noteTypeSize(card);
  const w = Math.max(cardW, Math.round(fs * 10));
  const h = Math.ceil(fs + 5);
  const dx = Math.round((cardW || w) * 0.18);
  const dy = cardH - Math.round(h * 0.45) + index * (h + 3);
  return { w, h, dx, dy, fs, x: (Number(card?.x) || 0) + dx, y: (Number(card?.y) || 0) + dy };
}

/** Height of a note for its text, so a second line still fits inside the rectangle. */
export function noteHeightFor(text, width, fontSize = noteTypeSize(null)) {
  const fs = Math.max(4.5, Number(fontSize) || noteTypeSize(null));
  const inner = Math.max(12, (Number(width) || 80) - 8);
  const perLine = Math.max(4, Math.floor(inner / (fs * 0.62)));
  const lines = String(text ?? "").split("\n").reduce((count, line) => count + Math.max(1, Math.ceil(line.length / perLine) || 1), 0);
  const lineH = fs * 1.2;
  return Math.min(fs * 8, Math.max(Math.ceil(fs + 5), (lines || 1) * lineH + 2));
}

function cardsByKey(shapes) {
  const map = new Map();
  for (const shape of shapes || []) {
    if (shape.type !== "card" || !shape.key) continue;
    if (!map.has(shape.key)) map.set(shape.key, []);
    map.get(shape.key).push(shape);
  }
  return map;
}

function cardForNote(cards, note) {
  const list = cards.get(note.key) || [];
  if (!list.length) return null;
  if (note.cardId) {
    const exact = list.find((card) => card.id === note.cardId);
    if (exact) return exact;
  }
  if (list.length === 1) return list[0];
  const ax = note.x - (Number.isFinite(note.dx) ? note.dx : 0);
  const ay = note.y - (Number.isFinite(note.dy) ? note.dy : 0);
  let best = list[0];
  let bestDist = Infinity;
  for (const card of list) {
    const dist = Math.hypot(card.x - ax, card.y - ay);
    if (dist < bestDist) {
      best = card;
      bestDist = dist;
    }
  }
  return best;
}

/** Put each note back on its card using the offset stored when it was added or dragged. */
export function stickNotes(shapes) {
  const cards = cardsByKey(shapes);
  for (const shape of shapes || []) {
    if (shape.type !== "note" || !shape.key) continue;
    const card = cardForNote(cards, shape);
    if (!card || !Number.isFinite(shape.dx) || !Number.isFinite(shape.dy)) continue;
    shape.x = card.x + shape.dx;
    shape.y = card.y + shape.dy;
    shape.cardId = card.id;
    shape.fs = noteTypeSize(card);
  }
}

/**
 * While a card is dragged or scaled, its notes follow.
 * A note that is itself being dragged keeps the new offset from its card.
 */
export function anchorNotes(shapes, movingIds) {
  const cards = cardsByKey(shapes);
  for (const shape of shapes || []) {
    if (shape.type !== "note" || !shape.key) continue;
    const card = cardForNote(cards, shape);
    if (!card) continue;
    if (movingIds.has(shape.id)) {
      shape.dx = shape.x - card.x;
      shape.dy = shape.y - card.y;
      shape.cardId = card.id;
    } else if (movingIds.has(card.id) && Number.isFinite(shape.dx) && Number.isFinite(shape.dy)) {
      shape.x = card.x + shape.dx;
      shape.y = card.y + shape.dy;
      shape.cardId = card.id;
    }
  }
}

/**
 * Notes already on the deck, one record per id.
 * The preferred slide wins, so an edit on the slide you are looking at is what the next chart uses.
 */
export function cardNoteRecords(slides, preferIndex = 0) {
  const byId = new Map();
  const take = (slide) => {
    for (const shape of slide?.shapes || []) {
      if (shape?.type !== "note" || !shape.key || !shape.id) continue;
      byId.set(shape.id, {
        id: shape.id,
        key: shape.key,
        cardId: shape.cardId || "",
        text: shape.text || "",
        fill: shape.fill || NOTE_COLORS[0],
        dx: Number.isFinite(shape.dx) ? shape.dx : 0,
        dy: Number.isFinite(shape.dy) ? shape.dy : 0,
        w: shape.w,
        h: shape.h,
      });
    }
  };
  (slides || []).forEach((slide, index) => {
    if (index !== preferIndex) take(slide);
  });
  take((slides || [])[preferIndex]);
  const byKey = new Map();
  for (const note of byId.values()) {
    if (!byKey.has(note.key)) byKey.set(note.key, []);
    byKey.get(note.key).push(note);
  }
  return byKey;
}

/** Stamp saved notes onto the cards of a newly generated slide. The same note id is not added twice. */
export function placeCardNotes(shapes, records) {
  const present = new Set();
  for (const shape of shapes || []) {
    if (shape.type === "note" && shape.id) present.add(shape.id);
  }
  const cards = cardsByKey(shapes);
  for (const [key, list] of cards) {
    const card = list[0];
    for (const note of records?.get(key) || []) {
      if (present.has(note.id)) continue;
      present.add(note.id);
      shapes.push({
        id: note.id,
        type: "note",
        key,
        cardId: card.id,
        text: note.text || "",
        fill: note.fill || NOTE_COLORS[0],
        dx: note.dx,
        dy: note.dy,
        w: note.w,
        h: note.h,
        fs: noteTypeSize(card),
        x: card.x + note.dx,
        y: card.y + note.dy,
        z: 8,
      });
    }
  }
}

function userNoteTexts(shapes, key) {
  const seen = new Set();
  const texts = [];
  for (const shape of shapes) {
    if (shape.type !== "note" || shape.key !== key) continue;
    const text = String(shape.text || "").trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    texts.push(text);
  }
  return texts;
}

/**
 * The slide as Markdown: title, statistics, the reporting tree as a nested list, and a table with the
 * data behind every card. Built from the cards' own copies of the people, so it matches the slide.
 * @param {{ number?: number, total?: number }} [meta]
 */
export function slideMarkdown(org, shapes, meta = {}) {
  const out = [];
  const title = shapes.find((shape) => shape.type === "title");
  if (title) {
    out.push(`# ${title.name || "Org chart"}`);
    if (title.department) out.push(`_${title.department}_`);
  } else out.push("# Org chart");
  if (meta.number) out.push(`\nSlide ${meta.number}${meta.total ? ` of ${meta.total}` : ""}`);

  const stats = shapes.find((shape) => shape.type === "stats");
  if (stats) out.push(`\n**${stats.heading || "Department statistics"}**\n\n${(stats.lines || []).map((line) => `- ${line}`).join("\n")}`);

  const cards = shapes.filter((shape) => shape.type === "card").map((card) => ({ card, person: shapePerson(org, card) }));
  const people = cards.filter((item) => item.person);
  if (people.length) {
    const inSlide = new Set(people.map((item) => item.person.key));
    const parentOf = ({ card, person }) => {
      const key = card.managerType === "Dotted" ? person.dottedKey : person.solidKey;
      return inSlide.has(key) && key !== person.key ? key : "";
    };
    const byParent = new Map();
    for (const item of people) {
      const parent = parentOf(item);
      if (!byParent.has(parent)) byParent.set(parent, []);
      byParent.get(parent).push(item);
    }
    const order = (list) => list.sort((a, b) => a.card.y - b.card.y || a.card.x - b.card.x);
    const line = ({ card, person }, depth) => {
      const name = person.tbh ? "To be hired" : `${moderated(person.surname)} ${moderated(person.name)}`.trim();
      const job = moderated(person.job);
      const facts = [];
      if (person.location) facts.push(person.location);
      if ((person.stats?.subordinates || 0) > 1) facts.push(`team of ${person.stats.subordinates}`);
      facts.push(...cardNotes(person, card), ...userNoteTexts(shapes, person.key));
      const tail = facts.length ? ` _(${facts.join("; ")})_` : "";
      return `${"  ".repeat(depth)}- **${name}**${job ? ` — ${job}` : ""}${tail}`;
    };
    const seen = new Set();
    const rows = [];
    const walk = (parent, depth) => {
      for (const item of order(byParent.get(parent) || [])) {
        const id = `${item.person.key}|${item.card.managerType}`;
        if (seen.has(id)) continue;
        seen.add(id);
        rows.push(item);
        out.push(line(item, depth));
        walk(item.person.key, depth + 1);
      }
    };
    out.push("\n## Structure\n");
    walk("", 0);
    // Cards whose manager is not on this slide (a loose selection) still get listed.
    for (const item of order(people.filter((entry) => !seen.has(`${entry.person.key}|${entry.card.managerType}`)))) {
      rows.push(item);
      out.push(line(item, 0));
    }

    out.push("\n## Data\n");
    out.push("| Name | Job | Department | Location | Reports to | Team | Notes |");
    out.push("| --- | --- | --- | --- | --- | --- | --- |");
    for (const { card, person } of rows) {
      const name = person.tbh ? "To be hired" : `${moderated(person.surname)} ${moderated(person.name)}`;
      const reports = card.managerType === "Dotted" ? person.dottedName : person.solidName;
      out.push(
        `| ${[name, moderated(person.job), person.department, person.location, reports, person.stats?.subordinates || 1, [...cardNotes(person, card), ...userNoteTexts(shapes, person.key)].join("; ")].map(mdCell).join(" | ")} |`,
      );
    }
  }

  const hidden = shapes.find((shape) => shape.type === "hiddenlist");
  if (hidden) out.push(`\n> ${hidden.text}`);

  // Nothing structured selected (only notes, links, shade labels): fall back to the plain text lines.
  if (!people.length && !title && !stats && !hidden) return shapesText(org, shapes);
  return `${out.join("\n")}\n`;
}
