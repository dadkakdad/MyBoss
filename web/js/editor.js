/**
 * Slide list, selection geometry, and cost-center stamps.
 * These stay free of the DOM so the ribbon and the filmstrip can share them.
 */

import { SLIDE } from "./layout.js";

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
 * Green 50×30 stamp at the card's bottom-right, matching showCostCenters:
 * Left + Width - 50, Top + Height - 20.
 */
export function showCostCenterStamps(org, slides, idFor) {
  let count = 0;
  for (const slide of slides) {
    const cards = slide.shapes.filter((shape) => shape.type === "card");
    slide.shapes = slide.shapes.filter((shape) => shape.type !== "cc");
    for (const card of cards) {
      const person = card.key && org?.people?.get(card.key);
      slide.shapes.push({
        id: idFor("cc"),
        type: "cc",
        x: card.x + card.w - 50,
        y: card.y + card.h - 20,
        w: 50,
        h: 30,
        z: 6,
        text: costCenterText(person?.costCenter),
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
