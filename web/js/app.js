import {
  buildOrg,
  cardStyle,
  categories,
  displayName,
  headcountLabel,
  isManager,
  moderated,
  peopleInCategory,
  personFullName,
  statisticsLines,
} from "./org.js";
import {
  CARDS,
  LINE,
  LINK_STYLE,
  dottedDash,
  SHADE_COLORS,
  departmentShadeAction,
  SLIDE,
  buildSlides,
  layoutDepartmentShades,
  layoutManualBranch,
  layoutSingleCard,
  makeMetrics,
  slideLinkTop,
} from "./layout.js";
import { deckToPptx, slideToCanvas } from "./export.js";
import {
  boxesIntersect,
  contentFrame,
  hideCostCenterStamps,
  offsetSlideTargets,
  prependCharts,
  removeSlide,
  scaleShapes,
  selectionBounds,
  shapeBox,
  showCostCenterStamps,
} from "./editor.js";

const $ = (id) => document.getElementById(id);

let serial = 1;
function nid(prefix) {
  serial += 1;
  return `u-${prefix}-${serial}`;
}

const state = {
  fileName: "",
  sheetName: "",
  matrix: null,
  org: null,
  category: "CEO",
  personKey: "",
  slides: [blankSlide()],
  active: 0,
  selection: new Set(),
  zoom: 1,
  undo: [],
  redo: [],
  shade: 0,
  tool: "move",
  costCenters: false,
  renaming: null,
  frame: { x: 0, y: 0, w: SLIDE.width, h: SLIDE.height },
};

function blankSlide() {
  return { id: nid("slide"), rootKey: null, parentKey: null, fit: null, shapes: [] };
}

function current() {
  return state.slides[state.active];
}

function toast(message) {
  const node = $("toast");
  node.textContent = message;
  node.classList.add("show");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => node.classList.remove("show"), 2400);
}

function remember() {
  state.undo.push(JSON.stringify({ slides: state.slides, active: state.active }));
  if (state.undo.length > 40) state.undo.shift();
  state.redo.length = 0;
}

function restore(stack, other) {
  const raw = stack.pop();
  if (!raw) return;
  other.push(JSON.stringify({ slides: state.slides, active: state.active }));
  const data = JSON.parse(raw);
  state.slides = data.slides;
  state.active = data.active;
  state.selection.clear();
  render();
}

function requireOrg() {
  if (state.org && !state.org.fatal) return true;
  toast(state.org ? "Fix the sheet before drawing." : "Import a sheet first.");
  return false;
}

function selectedPersonKey() {
  const slide = current();
  for (const id of state.selection) {
    const shape = slide.shapes.find((item) => item.id === id);
    if (shape?.key && state.org?.people.has(shape.key)) return shape.key;
  }
  return state.personKey;
}

function readWorkbook(buffer, fileName) {
  if (!/\.(xlsx|xls|xlsm)$/i.test(fileName)) {
    toast("File extension should be .xls, .xlsx, or .xlsm.");
    return;
  }
  const workbook = window.XLSX.read(buffer, { type: "array", cellDates: true });
  state.fileName = fileName;
  state.imported = false;
  state.org = null;
  state.workbookSheets = workbook.SheetNames.map((name) => ({
    name,
    matrix: window.XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "", raw: true }),
  }));
  const preferred = state.workbookSheets.findIndex((sheet) => /advanced/i.test(sheet.name));
  fillSheets(preferred >= 0 ? preferred : 0);
}

function fillSheets(index) {
  const select = $("sheet");
  select.innerHTML = "";
  for (const sheet of state.workbookSheets) {
    const option = document.createElement("option");
    option.value = sheet.name;
    option.textContent = sheet.name;
    select.appendChild(option);
  }
  select.selectedIndex = index;
  state.sheetName = select.value;
  const name = $("file-name");
  if (name) name.value = state.fileName || "";
  document.body.dataset.stage = "file";
}

function importSheet() {
  const sheet = state.workbookSheets?.find((item) => item.name === $("sheet").value);
  if (!sheet) {
    toast("Choose an Excel file first.");
    return;
  }
  state.sheetName = sheet.name;
  state.matrix = sheet.matrix;
  state.org = buildOrg(sheet.matrix, { hasTitles: $("titles").checked });
  state.category = "CEO";
  state.personKey = state.org.ceo || state.org.order[0] || "";
  fillPeople();
  if (state.org.fatal) {
    state.imported = false;
    document.body.dataset.stage = "file";
    render();
    toast("Import found problems. The chart stays blocked until they are fixed.");
    return;
  }
  state.imported = true;
  document.body.dataset.stage = "ready";
  render();
  const size = state.org.companySize === 1 ? "1 person" : `${state.org.companySize} people`;
  toast(`Imported ${state.sheetName}. CEO ${personFullName(state.org.people.get(state.org.ceo))}, ${size}.`);
}

function fillPeople() {
  const org = state.org;
  const category = $("category");
  const employee = $("employee");
  category.innerHTML = "";
  employee.innerHTML = "";
  if (!org) return;
  for (const name of categories(org)) {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name === "AZ" ? "All, A–Z" : name === "Managers" ? "Managers" : name;
    category.appendChild(option);
  }
  if (![...category.options].some((option) => option.value === state.category)) state.category = "CEO";
  category.value = state.category;
  const keys = peopleInCategory(org, state.category);
  for (const key of keys) {
    const option = document.createElement("option");
    option.value = key;
    option.textContent = displayName(org.people.get(key));
    employee.appendChild(option);
  }
  if (!keys.includes(state.personKey)) state.personKey = keys[0] || "";
  if (state.personKey) employee.value = state.personKey;
}

function metrics() {
  return makeMetrics($("size").value, Number($("gap").value), Number($("gap").value));
}

function createDeck(mode) {
  if (!requireOrg()) return;
  const key = $("employee").value || state.org.ceo;
  remember();
  const fresh = buildSlides(state.org, key, mode);
  state.slides = prependCharts(state.slides, fresh);
  state.active = 0;
  state.selection.clear();
  state.renaming = null;
  if (state.costCenters) showCostCenterStamps(state.org, state.slides, nid);
  const hidden = hiddenTeams(state.slides[0]);
  render();
  const placed = fresh.length === 1 ? "Added 1 slide in front." : `Added ${fresh.length} slides in front.`;
  const fit = mode === "infinity" ? "" : hidden ? ` ${hidden} team${hidden === 1 ? "" : "s"} folded to fit.` : " Chart fits on one slide.";
  toast(`${placed}${fit}`);
}

function hiddenTeams(slide) {
  if (!slide?.fit || !state.org) return 0;
  let count = 0;
  const entries = slide.fit.visibility instanceof Map ? [...slide.fit.visibility.entries()] : slide.fit.visibility;
  for (const [key, flag] of entries) {
    if (flag === 0 && isManager(state.org, key)) count += 1;
  }
  return count;
}

function placeCard(head) {
  if (!requireOrg()) return;
  const key = selectedPersonKey();
  if (!key) return;
  remember();
  const size = head ? "Large" : $("size").value;
  const cardMetrics = makeMetrics(size, Number($("gap").value));
  const slide = current();
  slide.shapes = slide.shapes.filter((shape) => !(shape.key === key && ["card", "bubble", "comment"].includes(shape.type)));
  const x = head ? SLIDE.width / 2 - cardMetrics.w / 2 : 48 + (slide.shapes.length % 5) * 20;
  const y = head ? 96 : 120 + (slide.shapes.length % 4) * 16;
  slide.shapes.push(...layoutSingleCard(state.org, key, cardMetrics, { head, x, y }));
  if (head && !slide.shapes.some((shape) => shape.type === "title")) {
    slide.shapes.unshift(titleFor(key));
    slide.shapes.push(
      { id: nid("legend"), type: "legend", x: 18, y: 66, w: 210, h: 58, z: 5 },
      {
        id: nid("stats"),
        type: "stats",
        x: 742,
        y: 10,
        w: 204,
        h: 96,
        z: 5,
        heading: "Department statistics",
        lines: statisticsLines(state.org.people.get(key), (state.org.reports.get(key) || []).length),
      },
    );
  }
  slide.rootKey = slide.rootKey || key;
  render();
}

function titleFor(key) {
  const person = state.org.people.get(key);
  const direct = (state.org.reports.get(key) || []).length;
  return {
    id: nid("title"),
    type: "title",
    x: 18,
    y: 12,
    w: 430,
    h: 74,
    z: 3,
    key,
    name: personFullName(person),
    department: person.department,
    direct,
    stats: person.stats,
  };
}

function placeBranch(mode) {
  if (!requireOrg()) return;
  const key = selectedPersonKey();
  const person = state.org.people.get(key);
  if (!person) return;
  if ((state.org.reports.get(key) || []).length === 0) {
    toast(`${personFullName(person)} has no direct reports.`);
    return;
  }
  remember();
  const slide = current();
  slide.shapes = slide.shapes.filter((shape) => {
    if (shape.type === "title" || shape.type === "legend" || shape.type === "nav" || shape.type === "stats") return true;
    if (shape.key === key && (shape.type === "card" || shape.type === "bubble" || shape.type === "comment")) return true;
    return false;
  });
  const box = layoutManualBranch(state.org, key, metrics(), mode);
  const anchor = cardAnchor(slide, key);
  const dx = anchor.x - box.anchorX;
  const dy = anchor.y;
  for (const item of box.items) {
    const shape = { ...item, id: nid(item.type), z: item.type === "wire" ? 1 : 2, parentKey: key };
    if (shape.type === "wire") {
      shape.x1 += dx;
      shape.y1 += dy;
      shape.x2 += dx;
      shape.y2 += dy;
    } else {
      shape.x += dx;
      shape.y += dy;
    }
    slide.shapes.push(shape);
  }
  render();
}

function cardAnchor(slide, key) {
  const card = slide.shapes.find((shape) => shape.type === "card" && shape.key === key);
  if (card) return { x: card.x + card.w / 2, y: card.y + card.h + 4 };
  return { x: SLIDE.width / 2, y: 168 };
}

function lightenShade(fill) {
  const hex = String(fill || "#DEEBF7").replace("#", "");
  const channel = (start) => parseInt(hex.slice(start, start + 2), 16);
  const mix = (value) => Math.min(255, Math.round(value + (255 - value) / 2));
  const next = [mix(channel(0)), mix(channel(2)), mix(channel(4))];
  if (next.some((value) => value >= 250)) return "#8fa4c4";
  return `#${next.map((value) => value.toString(16).padStart(2, "0")).join("")}`;
}

function shadeDepartments() {
  if (!requireOrg()) return;
  const slide = current();
  const action = departmentShadeAction(slide.shapes, state.selection);
  if (action === "lighten") {
    const selectedShade = selectedShapes().find((shape) => shape.type === "shade");
    remember();
    selectedShade.fill = lightenShade(selectedShade.fill);
    render();
    return;
  }
  if (action === "remove") {
    remember();
    slide.shapes = slide.shapes.filter((shape) => shape.type !== "shade");
    for (const id of [...state.selection]) {
      if (!slide.shapes.some((shape) => shape.id === id)) state.selection.delete(id);
    }
    render();
    toast("Department shades removed.");
    return;
  }
  remember();
  slide.shapes = slide.shapes.filter((shape) => shape.type !== "shade");
  const shades = layoutDepartmentShades(
    state.org,
    slide.shapes.filter((shape) => shape.type === "card"),
  ).map((shade) => ({ ...shade, id: nid("shade") }));
  slide.shapes.push(...shades);
  render();
  toast(shades.length ? `${shades.length} department shades.` : "No departments on this slide.");
}

function linkSlides() {
  if (!requireOrg()) return;
  const already = state.slides.some((slide) => slide.shapes.some((shape) => shape.type === "link" || shape.type === "nav"));
  if (already) {
    remember();
    for (const slide of state.slides) {
      slide.shapes = slide.shapes.filter((shape) => shape.type !== "link" && shape.type !== "nav");
    }
    render();
    toast("Links removed.");
    return;
  }
  remember();
  const found = new Map();
  state.slides.forEach((slide, index) => {
    for (const shape of slide.shapes) {
      if (shape.type !== "card" || !shape.key) continue;
      if (!isManager(state.org, shape.key)) continue;
      const list = found.get(shape.key) || [];
      list.push({ slide: index + 1, shape });
      found.set(shape.key, list);
    }
  });
  let links = 0;
  for (const slide of state.slides) {
    slide.shapes = slide.shapes.filter((shape) => shape.type !== "link" && shape.type !== "nav");
  }
  for (const [key, hits] of found) {
    const uniqueSlides = [...new Set(hits.map((hit) => hit.slide))];
    if (uniqueSlides.length !== 2) continue;
    const [earlier, later] = uniqueSlides.sort((a, b) => a - b);
    const source = hits.find((hit) => hit.slide === earlier).shape;
    const slide = state.slides[earlier - 1];
    slide.shapes.push({
      id: nid("link"),
      type: "link",
      x: source.x,
      y: slideLinkTop(source, slide.shapes),
      w: source.w,
      h: 11,
      z: 3,
      text: `See slide ${later}`,
      target: later,
      key,
    });
    const parent = state.org.people.get(key);
    state.slides[later - 1].shapes.push({
      id: nid("nav"),
      type: "nav",
      x: 460,
      y: 14,
      w: 280,
      h: 18,
      z: 4,
      text: `Up to ${parent.surname} ${parent.name} (slide ${earlier})`,
      target: earlier,
    });
    links += 1;
  }
  render();
  toast(links ? `${links} cross-links.` : "Links need a person on exactly two slides.");
}

function addDecoration(kind) {
  if (kind === "title" && !requireOrg()) return;
  remember();
  const slide = current();
  if (kind === "slide") {
    const at = state.active + 1;
    state.slides.splice(at, 0, blankSlide());
    offsetSlideTargets(state.slides, 1, at);
    state.active = at;
    render();
    return;
  }
  if (kind === "title") {
    slide.shapes.push(titleFor(selectedPersonKey() || state.org.ceo));
  } else if (kind === "shade") {
    state.shade = (state.shade + 1) % SHADE_COLORS.length;
    slide.shapes.push({
      id: nid("shade"),
      type: "shade",
      x: 80,
      y: 180,
      w: 160,
      h: 220,
      z: 0,
      text: "Department",
      fill: SHADE_COLORS[state.shade],
    });
  } else if (kind === "frame") {
    slide.shapes.push({ id: nid("frame"), type: "frame", x: 70, y: 160, w: 180, h: 240, z: 4 });
  } else if (kind === "card") {
    const card = CARDS[$("size").value];
    ["ordinary", "important", "outsource", "tbh"].forEach((placeholder, index) => {
      slide.shapes.push({
        id: nid("card"),
        type: "card",
        x: 40 + index * (card.w + 8),
        y: 200,
        w: card.w,
        h: card.h,
        z: 2,
        placeholder,
        lines: ["Surname", "Name", placeholder === "tbh" ? "To be hired" : "Job title"],
        size: card === CARDS.Large ? "Large" : $("size").value,
      });
    });
  } else if (kind === "line") {
    slide.shapes.push({
      id: nid("wire"),
      type: "wire",
      x1: 120,
      y1: 220,
      x2: 280,
      y2: 220,
      z: 1,
      weight: 2,
      dotted: false,
      manual: true,
    });
  }
  render();
}

function selectedShapes() {
  const slide = current();
  return slide.shapes.filter((shape) => state.selection.has(shape.id));
}

function boxShapes(list = selectedShapes()) {
  return list.filter((shape) => Number.isFinite(shape.x) && Number.isFinite(shape.w));
}

function arrange(kind) {
  const boxes = boxShapes();
  if (kind === "group") {
    if (boxes.length < 2) return toast("Select at least two shapes.");
    remember();
    const groupId = nid("group");
    for (const shape of boxes) shape.groupId = groupId;
    render();
    return;
  }
  if (kind === "ungroup") {
    remember();
    for (const shape of selectedShapes()) delete shape.groupId;
    render();
    return;
  }
  if (kind === "delete") {
    if (state.selection.size === 0) return;
    remember();
    const slide = current();
    slide.shapes = slide.shapes.filter((shape) => !state.selection.has(shape.id));
    state.selection.clear();
    render();
    return;
  }
  if (boxes.length < 2) return;
  if ((kind === "dist-h" || kind === "dist-v") && boxes.length < 3) return;
  remember();
  if (kind === "align-left") {
    const left = Math.min(...boxes.map((shape) => shape.x));
    for (const shape of boxes) shape.x = left;
  } else if (kind === "align-top") {
    const top = Math.min(...boxes.map((shape) => shape.y));
    for (const shape of boxes) shape.y = top;
  } else if (kind === "dist-h") {
    const ordered = [...boxes].sort((a, b) => a.x - b.x);
    const first = ordered[0].x;
    const last = ordered[ordered.length - 1].x;
    const step = (last - first) / (ordered.length - 1);
    ordered.forEach((shape, index) => {
      shape.x = first + step * index;
    });
  } else if (kind === "dist-v") {
    const ordered = [...boxes].sort((a, b) => a.y - b.y);
    const first = ordered[0].y;
    const last = ordered[ordered.length - 1].y;
    const step = (last - first) / (ordered.length - 1);
    ordered.forEach((shape, index) => {
      shape.y = first + step * index;
    });
  }
  render();
}

function nudge(dx, dy) {
  const shapes = selectedShapes();
  if (!shapes.length) return;
  remember();
  for (const shape of shapes) moveShape(shape, dx, dy);
  render();
}

function moveShape(shape, dx, dy) {
  if (shape.type === "wire") {
    shape.x1 += dx;
    shape.y1 += dy;
    shape.x2 += dx;
    shape.y2 += dy;
  } else {
    shape.x += dx;
    shape.y += dy;
  }
}

function selectGroup(shape, additive) {
  if (!additive) state.selection.clear();
  if (!shape) return;
  const slide = current();
  const members = shape.groupId ? slide.shapes.filter((item) => item.groupId === shape.groupId) : [shape];
  for (const member of members) state.selection.add(member.id);
}

function applySticky(shape) {
  if (!$("sticky").checked || shape?.type !== "card") return;
  const height = shape.h;
  let match = "Medium";
  let best = Infinity;
  for (const [name, card] of Object.entries(CARDS)) {
    const delta = Math.abs(card.h - height);
    if (delta < best) {
      best = delta;
      match = name;
    }
  }
  if (best < 5) $("size").value = match;
  const slide = current();
  const center = shape.x + shape.w / 2;
  const mid = shape.y + shape.h / 2;
  let length = null;
  for (const wire of slide.shapes) {
    if (wire.type !== "wire") continue;
    if (Math.abs(wire.x1 - wire.x2) < 1 && Math.abs(wire.x1 - center) < 4) {
      const top = Math.min(wire.y1, wire.y2);
      const bottom = Math.max(wire.y1, wire.y2);
      if (Math.abs(bottom - shape.y) < 2 || Math.abs(top - (shape.y + shape.h)) < 2) length = Math.abs(wire.y2 - wire.y1);
    }
    if (Math.abs(wire.y1 - wire.y2) < 1 && Math.abs(wire.y1 - mid) < 4 && Math.abs(Math.max(wire.x1, wire.x2) - shape.x) < 2) {
      length = Math.abs(wire.x2 - wire.x1);
    }
  }
  const ticks = ["5", "7", "10", "12", "15"];
  const snapped = ticks.find((tick) => length != null && Math.abs(Number(tick) - length) < 1);
  if (snapped) $("gap").value = snapped;
}

function render() {
  state.frame = contentFrame(current()?.shapes || []);
  renderSlide();
  renderFilmstrip();
  renderInspector();
  renderStatus();
  applyScale();
  syncChrome();
}

function viewX(x) {
  return x - state.frame.x;
}

function viewY(y) {
  return y - state.frame.y;
}

function syncChrome() {
  const imported = Boolean(state.imported && state.org && !state.org.fatal);
  if (!document.body.dataset.stage) document.body.dataset.stage = "empty";
  if (imported) document.body.dataset.stage = "ready";
  const refresh = $("refresh");
  if (refresh && imported) {
    const ceo = state.org.people.get(state.org.ceo);
    const size = state.org.companySize === 1 ? "1 person" : `${state.org.companySize} people`;
    refresh.title = `Input File: ${state.fileName}\nSheet: ${state.sheetName}\nCEO Name: ${ceo ? personFullName(ceo) : ""}\nCompany Size: ${size}`;
  }
  const grouped = selectedShapes().some((shape) => shape.groupId);
  if ($("group")) $("group").disabled = !imported || boxShapes().length < 2;
  if ($("ungroup")) $("ungroup").disabled = !imported || !grouped;
  const card = current()?.shapes.find((shape) => state.selection.has(shape.id) && shape.type === "card" && shape.key);
  if ($("add-link")) $("add-link").disabled = !card;
  document.body.dataset.tool = state.tool;
  for (const [id, tool] of [["tool-select", "select"], ["tool-scale", "scale"]]) {
    const button = $(id);
    if (!button) continue;
    const on = state.tool === tool;
    button.classList.toggle("pressed", on);
    button.setAttribute("aria-pressed", String(on));
  }
  const costs = $("cost-centers");
  if (costs) {
    costs.classList.toggle("pressed", state.costCenters);
    costs.setAttribute("aria-pressed", String(state.costCenters));
  }
}

function renderSlide() {
  const slide = current();
  const root = $("slide");
  root.innerHTML = "";
  const frame = state.frame;
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.classList.add("wires");
  svg.setAttribute("viewBox", `${frame.x} ${frame.y} ${frame.w} ${frame.h}`);
  svg.style.left = "0";
  svg.style.top = "0";
  svg.style.width = `${frame.w}px`;
  svg.style.height = `${frame.h}px`;
  root.appendChild(svg);

  if (!slide.shapes.length) {
    const empty = document.createElement("div");
    empty.className = "empty-slide";
    empty.innerHTML = state.imported
      ? "<div><strong>Empty slide</strong>Choose a person, then Create org chart.</div>"
      : "<div><strong>Empty slide</strong>Select an Excel file, import the sheet, then create the chart.</div>";
    root.appendChild(empty);
  }

  const shapes = [...slide.shapes].sort((a, b) => (a.z || 0) - (b.z || 0));
  for (const shape of shapes) {
    if (shape.type === "wire") {
      const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
      line.setAttribute("x1", shape.x1);
      line.setAttribute("y1", shape.y1);
      line.setAttribute("x2", shape.x2);
      line.setAttribute("y2", shape.y2);
      line.setAttribute("stroke", shape.dotted ? LINE.dotted : LINE.solid);
      line.setAttribute("stroke-width", shape.weight || 2);
      line.setAttribute("stroke-linecap", shape.dotted ? "round" : "square");
      if (shape.dotted) line.setAttribute("stroke-dasharray", dottedDash(shape.weight || 2).join(" "));
      line.dataset.id = shape.id;
      line.style.cursor = "pointer";
      if (state.selection.has(shape.id)) line.setAttribute("stroke", "#c45500");
      line.addEventListener("pointerdown", (event) => startPointer(event, shape));
      svg.appendChild(line);
      continue;
    }
    root.appendChild(renderShape(shape));
  }
  const bounds = selectionBounds(selectedShapes());
  if (bounds) root.appendChild(selectionFrame(bounds));
}

function selectionFrame(bounds) {
  const frame = document.createElement("div");
  frame.className = "selection-frame";
  frame.style.left = `${viewX(bounds.x)}px`;
  frame.style.top = `${viewY(bounds.y)}px`;
  frame.style.width = `${bounds.w}px`;
  frame.style.height = `${bounds.h}px`;
  const handle = document.createElement("button");
  handle.type = "button";
  handle.className = "scale-handle";
  handle.title = "Drag to scale the selection";
  handle.setAttribute("aria-label", "Scale selection");
  const visual = 12 / Math.max(0.4, baseScale() * state.zoom);
  handle.style.width = `${visual}px`;
  handle.style.height = `${visual}px`;
  handle.style.right = `${-visual / 2}px`;
  handle.style.bottom = `${-visual / 2}px`;
  handle.addEventListener("pointerdown", (event) => beginScale(event, "corner"));
  frame.appendChild(handle);
  return frame;
}

function renderShape(shape) {
  const node = document.createElement("button");
  node.type = "button";
  node.className = `shape ${shape.type}`;
  node.dataset.id = shape.id;
  node.style.left = `${viewX(shape.x)}px`;
  node.style.top = `${viewY(shape.y)}px`;
  node.style.width = `${shape.w}px`;
  node.style.height = `${shape.h}px`;
  node.style.zIndex = String(shape.z || 2);
  if (state.selection.has(shape.id)) node.classList.add("selected");

  if (shape.type === "card") {
    const person = shape.key ? state.org?.people.get(shape.key) : null;
    const kind = person ? cardStyle(person) : shape.placeholder || "ordinary";
    node.classList.add(kind);
    const lines = person
      ? person.tbh
        ? ["To be hired", moderated(person.job)]
        : [moderated(person.surname), moderated(person.name), moderated(person.job)]
      : shape.lines || [];
    const spec = CARDS[shape.size] || CARDS.Medium;
    const fontScale = spec.h ? shape.h / spec.h : 1;
    node.style.fontSize = `${spec.name * fontScale}px`;
    const badge = person ? headcountLabel(person) : "";
    const footer = person?.location || "";
    const badgeHtml = badge ? `<span class="hc">${escapeHtml(badge)}</span>` : "";
    node.innerHTML = `<span class="who">${escapeHtml(lines[0] || "")}</span>
      <span class="who">${escapeHtml(lines[1] || "")}</span>
      ${shape.h >= 32 && lines[2] ? `<span class="job" style="font-size:${spec.title * fontScale}px">${escapeHtml(lines[2])}</span>` : ""}
      ${footer ? `<span class="loc"><span class="loc-text">${escapeHtml(footer)}</span>${badgeHtml}</span>` : badgeHtml}`;
    node.setAttribute("aria-label", person ? personFullName(person) : "Empty card");
    node.addEventListener("dblclick", (event) => {
      event.stopPropagation();
      focusPerson(shape.key);
    });
  } else if (shape.type === "title") {
    node.innerHTML = `<div class="name">${escapeHtml(shape.name || "")}</div>
      <div class="dept">${escapeHtml(shape.department || "Organisation")}</div>`;
  } else if (shape.type === "legend") {
    node.innerHTML = `<div class="legend"><div class="legend-title">Indicators</div><div><i class="ordinary"></i>Ordinary employee</div><div><i class="important"></i>Important employee</div><div><i class="outsource"></i>Outsourcer</div></div>`;
  } else if (shape.type === "stats") {
    node.innerHTML = `<div class="stats-title">${escapeHtml(shape.heading || "Department statistics")}</div><ul>${(shape.lines || []).map((line) => `<li>${escapeHtml(line)}</li>`).join("")}</ul>`;
  } else if (shape.type === "shade") {
    node.textContent = shape.text || "";
    node.style.background = shape.fill || "#DEEBF7";
  } else if (shape.type === "bubble") {
    node.textContent = shape.text || "";
    if (shape.dotted) node.classList.add("dotted");
  } else if (shape.type === "cc") {
    node.textContent = shape.text || "?";
    node.style.fontSize = `${Math.max(6, (shape.h || 30) * (10 / 30))}px`;
  } else if (shape.type === "comment" || shape.type === "link" || shape.type === "nav") {
    node.textContent = shape.text || "";
    if (shape.type === "link") {
      node.style.background = LINK_STYLE.fill;
      node.style.color = LINK_STYLE.color;
    }
  }

  node.addEventListener("pointerdown", (event) => startPointer(event, shape));
  if (shape.target) {
    node.addEventListener("click", (event) => {
      if (pointer.moved || state.tool !== "move") return;
      event.stopPropagation();
      state.active = shape.target - 1;
      state.selection.clear();
      render();
    });
  }
  return node;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
}

function focusPerson(key) {
  if (!key || !state.org?.people.has(key)) return;
  state.personKey = key;
  state.category = key.slice(0, 1).toUpperCase();
  fillPeople();
  $("employee").value = key;
}

const pointer = { id: null, x: 0, y: 0, moved: false, shape: null, marquee: null, scale: null };

function localPoint(event) {
  const frame = state.frame;
  const rect = $("slide").getBoundingClientRect();
  return {
    x: frame.x + ((event.clientX - rect.left) / rect.width) * frame.w,
    y: frame.y + ((event.clientY - rect.top) / rect.height) * frame.h,
  };
}

function startPointer(event, shape) {
  event.stopPropagation();
  if (state.tool === "select") {
    beginMarquee(event, shape);
    return;
  }
  if (state.tool === "scale") {
    if (shape && !state.selection.has(shape.id)) selectGroup(shape, event.shiftKey);
    if (shape?.type === "card" && shape.key) focusPerson(shape.key);
    beginScale(event, "center");
    return;
  }
  const point = localPoint(event);
  pointer.id = event.pointerId;
  pointer.x = point.x;
  pointer.y = point.y;
  pointer.moved = false;
  pointer.shape = shape;
  pointer.scale = null;
  if (shape && !state.selection.has(shape.id)) selectGroup(shape, event.shiftKey);
  if (shape?.type === "card" && shape.key) focusPerson(shape.key);
  else if (shape && event.shiftKey) {
    state.selection.delete(shape.id);
    pointer.shape = null;
  }
  applySticky(shape);
  event.currentTarget.setPointerCapture?.(event.pointerId);
  render();
}

function beginMarquee(event, shape) {
  event.stopPropagation();
  const point = localPoint(event);
  pointer.id = event.pointerId;
  pointer.x = point.x;
  pointer.y = point.y;
  pointer.moved = false;
  pointer.shape = null;
  pointer.scale = null;
  pointer.marquee?.remove();
  pointer.marqueeShape = shape || null;
  pointer.marqueeAdditive = event.shiftKey;
  if (!event.shiftKey && !shape) state.selection.clear();
  pointer.marquee = document.createElement("div");
  pointer.marquee.className = "marquee";
  $("slide").appendChild(pointer.marquee);
  $("slide").setPointerCapture?.(event.pointerId);
  if (!shape) renderInspector();
}

function beginScale(event, mode) {
  event.stopPropagation();
  event.preventDefault();
  const live = selectedShapes();
  const bounds = selectionBounds(live);
  if (!bounds) {
    toast("Select elements to scale.");
    return;
  }
  const point = localPoint(event);
  const origin = mode === "center"
    ? { x: bounds.x + bounds.w / 2, y: bounds.y + bounds.h / 2 }
    : { x: bounds.x, y: bounds.y };
  pointer.marquee?.remove();
  pointer.id = event.pointerId;
  pointer.moved = false;
  pointer.shape = null;
  pointer.marquee = null;
  pointer.marqueeShape = null;
  pointer.scale = {
    mode,
    origin,
    anchor: point,
    start: Math.max(12, Math.hypot(point.x - origin.x, point.y - origin.y)),
    snapshot: live.map((shape) => ({ ...shape })),
    live,
    remembered: false,
  };
  event.currentTarget.setPointerCapture?.(event.pointerId);
}

function onHolderPointerDown(event) {
  if (event.target.closest?.(".shape, .scale-handle, .wires line")) return;
  if (state.tool === "scale") return;
  beginMarquee(event, null);
}

function paintShape(shape) {
  const node = $("slide").querySelector(`[data-id="${shape.id}"]`);
  if (!node) return;
  if (shape.type === "wire") {
    node.setAttribute("x1", shape.x1);
    node.setAttribute("y1", shape.y1);
    node.setAttribute("x2", shape.x2);
    node.setAttribute("y2", shape.y2);
  } else {
    node.style.left = `${viewX(shape.x)}px`;
    node.style.top = `${viewY(shape.y)}px`;
    node.style.width = `${shape.w}px`;
    node.style.height = `${shape.h}px`;
  }
}

function applyScaleGesture(point) {
  const gesture = pointer.scale;
  const snapshot = gesture.snapshot.map((shape) => ({ ...shape }));
  let factor = 1;
  if (gesture.mode === "corner") {
    factor = Math.hypot(point.x - gesture.origin.x, point.y - gesture.origin.y) / gesture.start;
  } else {
    factor = 1 + (point.y - gesture.anchor.y) / 180;
  }
  scaleShapes(snapshot, factor, gesture.origin);
  gesture.live.forEach((shape, index) => {
    const next = snapshot[index];
    if (shape.type === "wire") {
      shape.x1 = next.x1;
      shape.y1 = next.y1;
      shape.x2 = next.x2;
      shape.y2 = next.y2;
      shape.weight = next.weight;
    } else {
      shape.x = next.x;
      shape.y = next.y;
      shape.w = next.w;
      shape.h = next.h;
    }
    paintShape(shape);
  });
  const frame = $("slide").querySelector(".selection-frame");
  const bounds = selectionBounds(gesture.live);
  if (frame && bounds) {
    frame.style.left = `${viewX(bounds.x)}px`;
    frame.style.top = `${viewY(bounds.y)}px`;
    frame.style.width = `${bounds.w}px`;
    frame.style.height = `${bounds.h}px`;
  }
}

function onPointerMove(event) {
  if (pointer.id !== event.pointerId) return;
  const point = localPoint(event);
  const dx = point.x - pointer.x;
  const dy = point.y - pointer.y;
  if (!pointer.moved && Math.hypot(point.x - (pointer.scale?.anchor.x ?? pointer.x), point.y - (pointer.scale?.anchor.y ?? pointer.y)) < 2) return;
  if (pointer.scale) {
    if (!pointer.scale.remembered) {
      remember();
      pointer.scale.remembered = true;
    }
    pointer.moved = true;
    applyScaleGesture(point);
    return;
  }
  if (!pointer.moved && pointer.shape) remember();
  pointer.moved = true;
  if (pointer.marquee) {
    const x = Math.min(pointer.x, point.x);
    const y = Math.min(pointer.y, point.y);
    pointer.marquee.style.left = `${viewX(x)}px`;
    pointer.marquee.style.top = `${viewY(y)}px`;
    pointer.marquee.style.width = `${Math.abs(dx)}px`;
    pointer.marquee.style.height = `${Math.abs(dy)}px`;
    return;
  }
  if (!pointer.shape) return;
  const slide = current();
  const moving = slide.shapes.filter((shape) => state.selection.has(shape.id));
  for (const shape of moving) moveShape(shape, dx, dy);
  pointer.x = point.x;
  pointer.y = point.y;
  for (const shape of moving) paintShape(shape);
  const frame = $("slide").querySelector(".selection-frame");
  const bounds = selectionBounds(moving);
  if (frame && bounds) {
    frame.style.left = `${viewX(bounds.x)}px`;
    frame.style.top = `${viewY(bounds.y)}px`;
    frame.style.width = `${bounds.w}px`;
    frame.style.height = `${bounds.h}px`;
  }
}

function marqueeRect() {
  const box = pointer.marquee.getBoundingClientRect();
  const slideRect = $("slide").getBoundingClientRect();
  const frame = state.frame;
  const scaleX = frame.w / slideRect.width;
  const scaleY = frame.h / slideRect.height;
  return {
    x: frame.x + (box.left - slideRect.left) * scaleX,
    y: frame.y + (box.top - slideRect.top) * scaleY,
    w: box.width * scaleX,
    h: box.height * scaleY,
  };
}

function onPointerUp(event) {
  if (pointer.id !== event.pointerId) return;
  if (pointer.marquee) {
    if (pointer.moved) {
      const rect = marqueeRect();
      if (!pointer.marqueeAdditive) state.selection.clear();
      for (const shape of current().shapes) {
        const box = shapeBox(shape);
        if (!box) continue;
        const hit = shape.type === "wire"
          ? { x: box.x - 4, y: box.y - 4, w: box.w + 8, h: box.h + 8 }
          : box;
        if (!boxesIntersect(hit, rect)) continue;
        const members = shape.groupId ? current().shapes.filter((item) => item.groupId === shape.groupId) : [shape];
        for (const member of members) state.selection.add(member.id);
      }
    } else if (pointer.marqueeShape) {
      selectGroup(pointer.marqueeShape, pointer.marqueeAdditive);
    } else if (!pointer.marqueeAdditive) {
      state.selection.clear();
    }
    pointer.marquee.remove();
    pointer.marquee = null;
    pointer.marqueeShape = null;
    render();
  } else if (pointer.scale) {
    pointer.scale = null;
    render();
  }
  pointer.id = null;
  pointer.shape = null;
}

function slideCaption(slide, index) {
  const person = slide.rootKey && state.org ? state.org.people.get(slide.rootKey) : null;
  const cards = slide.shapes.filter((shape) => shape.type === "card").length;
  return {
    title: slide.title || `Slide ${index + 1}`,
    detail: person ? displayName(person) : cards ? `${cards} cards` : "Empty",
  };
}

function commitRename(index, value) {
  if (state.renaming !== index) return;
  state.renaming = null;
  const slide = state.slides[index];
  if (!slide) {
    renderFilmstrip();
    return;
  }
  const trimmed = value.trim();
  const next = !trimmed || trimmed === `Slide ${index + 1}` ? "" : trimmed;
  if ((slide.title || "") === next) {
    renderFilmstrip();
    return;
  }
  remember();
  slide.title = next;
  render();
}

function deleteSlide(index) {
  remember();
  state.renaming = null;
  state.slides = removeSlide(state.slides, index);
  if (!state.slides.length) state.slides = [blankSlide()];
  if (state.active > index) state.active -= 1;
  if (state.active >= state.slides.length) state.active = state.slides.length - 1;
  state.selection.clear();
  render();
}

function renderFilmstrip() {
  const strip = $("filmstrip");
  strip.innerHTML = "";
  state.slides.forEach((slide, index) => {
    const tab = document.createElement("div");
    tab.className = `slide-tab${index === state.active ? " active" : ""}`;
    const open = document.createElement("div");
    open.className = "open";
    open.tabIndex = 0;
    open.setAttribute("role", "button");
    const caption = slideCaption(slide, index);
    if (state.renaming === index) {
      const input = document.createElement("input");
      input.className = "slide-name";
      input.value = slide.title || `Slide ${index + 1}`;
      input.setAttribute("aria-label", "Slide name");
      input.addEventListener("click", (event) => event.stopPropagation());
      input.addEventListener("keydown", (event) => {
        event.stopPropagation();
        if (event.key === "Enter") {
          event.preventDefault();
          commitRename(index, input.value);
        } else if (event.key === "Escape") {
          state.renaming = null;
          renderFilmstrip();
        }
      });
      input.addEventListener("blur", () => commitRename(index, input.value));
      const detail = document.createElement("small");
      detail.textContent = caption.detail;
      open.append(input, detail);
      queueMicrotask(() => {
        input.focus();
        input.select();
      });
    } else {
      open.innerHTML = `<b>${escapeHtml(caption.title)}</b><small>${escapeHtml(caption.detail)}</small>`;
      open.title = "Double-click to rename";
      open.addEventListener("click", () => {
        state.active = index;
        state.selection.clear();
        render();
      });
      open.addEventListener("dblclick", (event) => {
        event.preventDefault();
        state.active = index;
        state.renaming = index;
        render();
      });
    }
    const actions = document.createElement("span");
    actions.className = "slide-actions";
    const rename = document.createElement("button");
    rename.type = "button";
    rename.className = "slide-mini";
    rename.title = "Rename slide";
    rename.setAttribute("aria-label", `Rename slide ${index + 1}`);
    rename.textContent = "✎";
    rename.addEventListener("click", (event) => {
      event.stopPropagation();
      state.active = index;
      state.renaming = index;
      render();
    });
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "slide-mini danger";
    remove.title = "Delete slide";
    remove.setAttribute("aria-label", `Delete slide ${index + 1}`);
    remove.textContent = "×";
    remove.addEventListener("click", (event) => {
      event.stopPropagation();
      deleteSlide(index);
    });
    actions.append(rename, remove);
    tab.append(open, actions);
    strip.appendChild(tab);
  });
}

function renderInspector() {
  const summary = $("sheet-summary");
  const issues = $("issues");
  const personNode = $("person");
  personNode.innerHTML = "";
  if (!state.org) {
    summary.innerHTML = "<p>One row per person. The CEO is the only row with an empty solid-line manager. Everyone else names the person they report to.</p>";
    issues.innerHTML = "<p class='ok'>Nothing imported yet.</p>";
    return;
  }
  const ceo = state.org.ceo ? state.org.people.get(state.org.ceo) : null;
  summary.innerHTML = `<p><b>${escapeHtml(state.sheetName)}</b> · ${state.org.companySize} people${ceo ? ` · CEO ${escapeHtml(personFullName(ceo))}` : ""}.</p>`;
  const drawn = new Set();
  for (const slide of state.slides) {
    for (const shape of slide.shapes) if (shape.type === "card" && shape.key) drawn.add(shape.key);
  }
  const missing = [...state.org.people.keys()].filter((key) => !drawn.has(key));
  const coverage = !drawn.size
    ? ""
    : !missing.length
      ? `<p class="ok">Every person is on a slide.</p>`
      : state.slides.length > 1
        ? `<p class="issue">The chart left out ${missing.length} ${missing.length === 1 ? "person" : "people"}.</p>`
        : "";
  if (state.org.issues.length === 0) issues.innerHTML = `<p class='ok'>Reporting lines reach the CEO.</p>${coverage}`;
  else {
    issues.innerHTML = `<ul>${state.org.issues
      .map((issue) => `<li class="${issue.level === "error" ? "issue" : ""}">${escapeHtml(issue.message)}</li>`)
      .join("")}</ul>`;
  }
  const key = [...state.selection]
    .map((id) => current().shapes.find((shape) => shape.id === id)?.key)
    .find((id) => id && state.org.people.has(id));
  const person = key ? state.org.people.get(key) : null;
  if (!person) {
    personNode.innerHTML = "<p>Select a card.</p>";
    return;
  }
  const solid = state.org.people.get(person.solidKey);
  const dotted = state.org.people.get(person.dottedKey);
  const rows = [
    ["Name", personFullName(person)],
    ["Job", person.job || "—"],
    ["Department", person.department || "—"],
    ["Location", person.location || "—"],
    ["Reports to", solid ? personFullName(solid) : "—"],
    ["Dotted line", dotted ? personFullName(dotted) : "—"],
    ["Team", headcountLabel(person) || "1"],
    ["Excel row", String(person.row)],
    ["Flags", [person.important ? "Key" : "", person.outsource ? "External" : "", person.tbh ? "To be hired" : ""].filter(Boolean).join(", ") || "Staff"],
    ["Cost center", person.costCenter || "—"],
  ];
  personNode.innerHTML = rows.map(([label, value]) => `<dt>${label}</dt><dd>${escapeHtml(value)}</dd>`).join("");
}

function renderStatus() {
  const fit = current()?.fit;
  const fitText = fit ? `${fit.size} · tick ${Math.round(fit.gap)} pt` : "";
  $("status-left").textContent = state.org
    ? `${state.fileName || "Workbook"} · ${state.sheetName} · ${state.org.companySize} people${fitText ? ` · ${fitText}` : ""}`
    : "No workbook yet.";
  $("zoom-label").textContent = `${Math.round(baseScale() * state.zoom * 100)}%`;
}

function baseScale() {
  const stage = $("stage");
  const availW = Math.max(200, stage.clientWidth - 28);
  const availH = Math.max(160, stage.clientHeight - 28);
  return Math.min(availW / SLIDE.width, availH / SLIDE.height, 1.4);
}

function applyScale() {
  const frame = state.frame || contentFrame(current()?.shapes || []);
  state.frame = frame;
  const scale = baseScale() * state.zoom;
  const holder = $("holder");
  holder.style.width = `${frame.w * scale}px`;
  holder.style.height = `${frame.h * scale}px`;
  const scaler = $("scaler");
  scaler.style.width = `${frame.w}px`;
  scaler.style.height = `${frame.h}px`;
  scaler.style.transform = `scale(${scale})`;
  const page = $("page");
  page.style.left = `${-frame.x}px`;
  page.style.top = `${-frame.y}px`;
  const slide = $("slide");
  slide.style.width = `${frame.w}px`;
  slide.style.height = `${frame.h}px`;
  $("zoom-label").textContent = `${Math.round(scale * 100)}%`;
}

function setTool(tool) {
  state.tool = state.tool === tool ? "move" : tool;
  syncChrome();
}

function toggleCostCenters() {
  if (!requireOrg()) return;
  remember();
  state.costCenters = !state.costCenters;
  if (state.costCenters) {
    const count = showCostCenterStamps(state.org, state.slides, nid);
    toast(count ? `Cost centers on ${count} ${count === 1 ? "card" : "cards"}.` : "No cards to label.");
  } else {
    hideCostCenterStamps(state.slides);
    toast("Cost centers hidden.");
  }
  render();
}

function downloadBlob(blob, name) {
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = name;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

function exportPng() {
  const canvas = slideToCanvas(state.org || { people: new Map() }, current());
  canvas.toBlob((blob) => {
    if (!blob) return;
    downloadBlob(blob, `myboss-slide-${state.active + 1}.png`);
  });
}

async function exportPptx() {
  try {
    const base = (state.fileName || "org-chart").replace(/\.(xlsx|xls|xlsm)$/i, "");
    await deckToPptx(state.org || { people: new Map() }, state.slides, `MyBoss-${base}.pptx`);
  } catch (error) {
    toast(error.message || "PowerPoint export failed.");
  }
}

function exportJson() {
  const payload = {
    fileName: state.fileName,
    sheetName: state.sheetName,
    slides: state.slides,
  };
  downloadBlob(new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" }), "myboss-project.json");
}

function onKey(event) {
  const typing = ["INPUT", "SELECT", "TEXTAREA"].includes(event.target.tagName);
  if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "z") {
    event.preventDefault();
    if (event.shiftKey) restore(state.redo, state.undo);
    else restore(state.undo, state.redo);
    return;
  }
  if (typing) return;
  if (event.key === "Delete" || event.key === "Backspace") {
    event.preventDefault();
    arrange("delete");
  } else if (event.key === "Escape") {
    state.selection.clear();
    render();
  } else if (event.key === "ArrowLeft") nudge(event.shiftKey ? -10 : -1, 0);
  else if (event.key === "ArrowRight") nudge(event.shiftKey ? 10 : 1, 0);
  else if (event.key === "ArrowUp") nudge(0, event.shiftKey ? -10 : -1);
  else if (event.key === "ArrowDown") nudge(0, event.shiftKey ? 10 : 1);
  else if (event.key === "+" || event.key === "=") {
    state.zoom = Math.min(2.5, state.zoom + 0.1);
    applyScale();
  } else if (event.key === "-" || event.key === "_") {
    state.zoom = Math.max(0.4, state.zoom - 0.1);
    applyScale();
  }
}

function toggleMenu(menuId, event) {
  event.stopPropagation();
  const menu = $(menuId);
  const willOpen = !menu.classList.contains("open");
  document.querySelectorAll(".menu.open").forEach((node) => node.classList.remove("open"));
  if (willOpen) menu.classList.add("open");
}

function bind() {
  $("open-file").addEventListener("click", () => $("file").click());
  $("pick-again").addEventListener("click", () => $("file").click());
  $("file").addEventListener("change", async () => {
    const file = $("file").files[0];
    if (!file) return;
    readWorkbook(await file.arrayBuffer(), file.name);
    toast(`${file.name}: choose a sheet and import.`);
  });
  $("import").addEventListener("click", importSheet);
  $("refresh").addEventListener("click", importSheet);
  $("pick-sheet").addEventListener("click", () => {
    state.imported = false;
    document.body.dataset.stage = "file";
    toast("Choose a sheet, then Import data.");
  });
  $("refresh-more").addEventListener("click", (event) => toggleMenu("menu-refresh", event));
  $("create-main").addEventListener("click", (event) => toggleMenu("menu-create", event));
  $("manager-more").addEventListener("click", (event) => toggleMenu("menu-manager", event));
  $("reports-more").addEventListener("click", (event) => toggleMenu("menu-reports", event));
  $("elements-main").addEventListener("click", (event) => toggleMenu("menu-elements", event));
  $("arrange-main").addEventListener("click", (event) => toggleMenu("menu-arrange", event));
  document.addEventListener("click", () => document.querySelectorAll(".menu.open").forEach((node) => node.classList.remove("open")));
  $("sheet").addEventListener("change", () => {
    state.sheetName = $("sheet").value;
  });
  $("category").addEventListener("change", () => {
    state.category = $("category").value;
    state.personKey = "";
    fillPeople();
  });
  $("employee").addEventListener("change", () => {
    state.personKey = $("employee").value;
  });
  $("single").addEventListener("click", () => createDeck("clever"));
  $("infinity").addEventListener("click", () => createDeck("infinity"));
  $("place-head").addEventListener("click", () => placeCard(true));
  $("place-one").addEventListener("click", () => placeCard(false));
  $("place-auto").addEventListener("click", () => placeBranch("auto"));
  $("place-1").addEventListener("click", () => placeBranch("1"));
  $("place-2").addEventListener("click", () => placeBranch("2"));
  $("place-4").addEventListener("click", () => placeBranch("4"));
  $("place-tree").addEventListener("click", () => placeBranch("tree"));
  $("departments").addEventListener("click", shadeDepartments);
  $("links").addEventListener("click", linkSlides);
  $("cost-centers").addEventListener("click", toggleCostCenters);
  $("tool-select").addEventListener("click", () => setTool("select"));
  $("tool-scale").addEventListener("click", () => setTool("scale"));
  $("add-slide").addEventListener("click", () => addDecoration("slide"));
  $("add-title").addEventListener("click", () => addDecoration("title"));
  $("add-shade").addEventListener("click", () => addDecoration("shade"));
  $("add-frame").addEventListener("click", () => addDecoration("frame"));
  $("add-card").addEventListener("click", () => addDecoration("card"));
  $("add-line").addEventListener("click", () => addDecoration("line"));
  $("add-link").addEventListener("click", () => {
    const card = current().shapes.find((shape) => state.selection.has(shape.id) && shape.type === "card");
    if (!card) return;
    remember();
    current().shapes.push({
      id: nid("link"),
      type: "link",
      x: card.x,
      y: slideLinkTop(card, current().shapes),
      w: card.w,
      h: 11,
      z: 4,
      text: "See slide",
      key: card.key,
    });
    render();
  });
  $("group").addEventListener("click", () => arrange("group"));
  $("ungroup").addEventListener("click", () => arrange("ungroup"));
  $("align-left").addEventListener("click", () => arrange("align-left"));
  $("align-top").addEventListener("click", () => arrange("align-top"));
  $("dist-h").addEventListener("click", () => arrange("dist-h"));
  $("dist-v").addEventListener("click", () => arrange("dist-v"));
  $("undo").addEventListener("click", () => restore(state.undo, state.redo));
  $("remove").addEventListener("click", () => arrange("delete"));
  $("png").addEventListener("click", exportPng);
  $("pptx").addEventListener("click", () => exportPptx());
  $("json").addEventListener("click", exportJson);
  $("zoom-in").addEventListener("click", () => {
    state.zoom = Math.min(2.5, state.zoom + 0.1);
    applyScale();
  });
  $("zoom-out").addEventListener("click", () => {
    state.zoom = Math.max(0.4, state.zoom - 0.1);
    applyScale();
  });
  $("zoom-fit").addEventListener("click", () => {
    state.zoom = 1;
    applyScale();
  });
  $("holder").addEventListener("pointerdown", onHolderPointerDown);
  window.addEventListener("pointermove", onPointerMove);
  window.addEventListener("pointerup", onPointerUp);
  window.addEventListener("keydown", onKey);
  window.addEventListener("resize", applyScale);
}

bind();
render();
