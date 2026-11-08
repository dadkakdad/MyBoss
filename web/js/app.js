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
  SHADE_COLORS,
  SLIDE,
  buildSlides,
  layoutDepartmentShades,
  layoutManualBranch,
  layoutSingleCard,
  makeMetrics,
} from "./layout.js";
import { deckToPptx, slideToCanvas } from "./export.js";

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
  state.slides = buildSlides(state.org, key, mode);
  state.active = 0;
  state.selection.clear();
  const hidden = hiddenTeams(state.slides[0]);
  render();
  toast(mode === "infinity" ? `${state.slides.length} slides.` : hidden ? `${hidden} team${hidden === 1 ? "" : "s"} folded to fit.` : "Chart fits on one slide.");
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
  const selectedShade = selectedShapes().find((shape) => shape.type === "shade");
  if (selectedShade) {
    remember();
    selectedShade.fill = lightenShade(selectedShade.fill);
    render();
    return;
  }
  remember();
  const slide = current();
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
      y: source.y + source.h,
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
    state.slides.splice(state.active + 1, 0, blankSlide());
    state.active += 1;
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
  renderSlide();
  renderFilmstrip();
  renderInspector();
  renderStatus();
  applyScale();
  syncChrome();
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
}

function renderSlide() {
  const slide = current();
  const root = $("slide");
  root.innerHTML = "";
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.classList.add("wires");
  svg.setAttribute("viewBox", `0 0 ${SLIDE.width} ${SLIDE.height}`);
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
      if (shape.dotted) line.setAttribute("stroke-dasharray", "0.8 2.8");
      line.dataset.id = shape.id;
      line.style.cursor = "pointer";
      if (state.selection.has(shape.id)) line.setAttribute("stroke", "#c45500");
      line.addEventListener("pointerdown", (event) => startPointer(event, shape));
      svg.appendChild(line);
      continue;
    }
    root.appendChild(renderShape(shape));
  }
}

function renderShape(shape) {
  const node = document.createElement("button");
  node.type = "button";
  node.className = `shape ${shape.type}`;
  node.dataset.id = shape.id;
  node.style.left = `${shape.x}px`;
  node.style.top = `${shape.y}px`;
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
    node.style.fontSize = `${spec.name}px`;
    const badge = person ? headcountLabel(person) : "";
    const footer = person?.location || "";
    const badgeHtml = badge ? `<span class="hc">${escapeHtml(badge)}</span>` : "";
    node.innerHTML = `<span class="who">${escapeHtml(lines[0] || "")}</span>
      <span class="who">${escapeHtml(lines[1] || "")}</span>
      ${shape.h >= 32 && lines[2] ? `<span class="job" style="font-size:${spec.title}px">${escapeHtml(lines[2])}</span>` : ""}
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
  } else if (shape.type === "comment" || shape.type === "link" || shape.type === "nav") {
    node.textContent = shape.text || "";
  }

  node.addEventListener("pointerdown", (event) => startPointer(event, shape));
  if (shape.target) {
    node.addEventListener("click", (event) => {
      if (pointer.moved) return;
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

const pointer = { id: null, x: 0, y: 0, moved: false, shape: null, marquee: null };

function localPoint(event) {
  const rect = $("slide").getBoundingClientRect();
  return {
    x: ((event.clientX - rect.left) / rect.width) * SLIDE.width,
    y: ((event.clientY - rect.top) / rect.height) * SLIDE.height,
  };
}

function startPointer(event, shape) {
  event.stopPropagation();
  const point = localPoint(event);
  pointer.id = event.pointerId;
  pointer.x = point.x;
  pointer.y = point.y;
  pointer.moved = false;
  pointer.shape = shape;
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

function onSlidePointerDown(event) {
  if (event.target !== $("slide")) return;
  const point = localPoint(event);
  pointer.id = event.pointerId;
  pointer.x = point.x;
  pointer.y = point.y;
  pointer.moved = false;
  pointer.shape = null;
  if (!event.shiftKey) state.selection.clear();
  pointer.marquee = document.createElement("div");
  pointer.marquee.className = "marquee";
  $("slide").appendChild(pointer.marquee);
  $("slide").setPointerCapture(event.pointerId);
  renderInspector();
}

function onPointerMove(event) {
  if (pointer.id !== event.pointerId) return;
  const point = localPoint(event);
  const dx = point.x - pointer.x;
  const dy = point.y - pointer.y;
  if (!pointer.moved && Math.hypot(dx, dy) < 2) return;
  if (!pointer.moved && pointer.shape) remember();
  pointer.moved = true;
  if (pointer.marquee) {
    const x = Math.min(pointer.x, point.x);
    const y = Math.min(pointer.y, point.y);
    pointer.marquee.style.left = `${x}px`;
    pointer.marquee.style.top = `${y}px`;
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
  for (const shape of moving) {
    const node = $("slide").querySelector(`[data-id="${shape.id}"]`);
    if (!node) continue;
    if (shape.type === "wire") {
      node.setAttribute("x1", shape.x1);
      node.setAttribute("y1", shape.y1);
      node.setAttribute("x2", shape.x2);
      node.setAttribute("y2", shape.y2);
    } else {
      node.style.left = `${shape.x}px`;
      node.style.top = `${shape.y}px`;
    }
  }
}

function onPointerUp(event) {
  if (pointer.id !== event.pointerId) return;
  if (pointer.marquee) {
    const box = pointer.marquee.getBoundingClientRect();
    const slideRect = $("slide").getBoundingClientRect();
    const scaleX = SLIDE.width / slideRect.width;
    const scaleY = SLIDE.height / slideRect.height;
    const rect = {
      x: (box.left - slideRect.left) * scaleX,
      y: (box.top - slideRect.top) * scaleY,
      w: box.width * scaleX,
      h: box.height * scaleY,
    };
    for (const shape of current().shapes) {
      if (!Number.isFinite(shape.x)) continue;
      const hits = shape.x < rect.x + rect.w && shape.x + shape.w > rect.x && shape.y < rect.y + rect.h && shape.y + shape.h > rect.y;
      if (hits) state.selection.add(shape.id);
    }
    pointer.marquee.remove();
    pointer.marquee = null;
    render();
  }
  pointer.id = null;
  pointer.shape = null;
}

function renderFilmstrip() {
  const strip = $("filmstrip");
  strip.innerHTML = "";
  state.slides.forEach((slide, index) => {
    const button = document.createElement("button");
    button.type = "button";
    if (index === state.active) button.className = "active";
    const person = slide.rootKey && state.org ? state.org.people.get(slide.rootKey) : null;
    const cards = slide.shapes.filter((shape) => shape.type === "card").length;
    button.innerHTML = `<b>Slide ${index + 1}</b><small>${escapeHtml(person ? displayName(person) : cards ? `${cards} cards` : "Empty")}</small>`;
    button.addEventListener("click", () => {
      state.active = index;
      state.selection.clear();
      render();
    });
    strip.appendChild(button);
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
  const scale = baseScale() * state.zoom;
  const holder = $("holder");
  holder.style.width = `${SLIDE.width * scale}px`;
  holder.style.height = `${SLIDE.height * scale}px`;
  const slide = $("slide");
  slide.style.transformOrigin = "top left";
  slide.style.transform = `scale(${scale})`;
  $("zoom-label").textContent = `${Math.round(scale * 100)}%`;
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
      y: card.y + card.h - 11,
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
  $("slide").addEventListener("pointerdown", onSlidePointerDown);
  window.addEventListener("pointermove", onPointerMove);
  window.addEventListener("pointerup", onPointerUp);
  window.addEventListener("keydown", onKey);
  window.addEventListener("resize", applyScale);
}

bind();
render();
