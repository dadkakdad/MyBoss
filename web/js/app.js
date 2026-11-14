import {
  COLUMN_LABELS,
  buildOrg,
  cardStyle,
  categories,
  displayName,
  headcountLabel,
  isManager,
  visibleReports,
  moderated,
  peopleInCategory,
  pickerLabel,
  VACANCIES,
  personFullName,
  statisticsLines,
  freezeCards,
  shapePerson,
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
  dropSubordinateShapes,
  layoutDepartmentShades,
  layoutManualBranch,
  layoutSingleCard,
  makeMetrics,
  slideLinkTop,
  hiddenListShape,
  statsBoxHeight,
} from "./layout.js";
import { deckToPptx, slideToCanvas } from "./export.js";
import {
  anchorNotes,
  boxesIntersect,
  cardNoteRecords,
  contentFrame,
  hideCostCenterStamps,
  noteBox,
  noteHeightFor,
  noteTypeSize,
  NOTE_COLORS,
  offsetSlideTargets,
  placeCardNotes,
  prependCharts,
  removeSlide,
  resizeShapes,
  scaleShapes,
  selectionBounds,
  slideMarkdown,
  shapeBox,
  showCostCenterStamps,
  stickNotes,
} from "./editor.js";
import {
  formatMoscowUploadTime,
  refreshSavedEntries,
  saveBody,
  workMenuLabel,
  workToReopen,
} from "./persist.js";
import {
  blankDraftRow,
  cellText,
  checkSheet,
  draftFromMatrix,
  flagValue,
  isFlagOn,
  pickSheet,
  problemRows,
  sheetProblem,
  trimTrailingBlankRows,
  vacancyDraftRow,
} from "./sheet.js";

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
  account: null,
  workId: null,
  fileBytes: null,
  saveLabel: "",
  hydrating: false,
  editingNote: null,
  noteEditReady: false,
  noteTargetId: "",
  deleteNoteId: "",
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
  scheduleSave();
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

const libraryLoads = new Map();

function assetUrl(name) {
  const node = document.getElementById("myboss-assets");
  if (!node) return "";
  try {
    return JSON.parse(node.textContent)[name] || "";
  } catch {
    return "";
  }
}

function loadScript(url) {
  if (!url) return Promise.reject(new Error("A required library is missing."));
  const existing = libraryLoads.get(url);
  if (existing) return existing;
  const loading = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = url;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("Could not load a library."));
    document.head.appendChild(script);
  });
  libraryLoads.set(url, loading);
  return loading;
}

function ensureXlsx() {
  if (window.XLSX) return Promise.resolve();
  return loadScript(assetUrl("xlsx"));
}

function ensurePptx() {
  if (window.PptxGenJS) return Promise.resolve();
  return loadScript(assetUrl("pptx"));
}

async function readWorkbook(buffer, fileName) {
  if (!/\.(xlsx|xls|xlsm)$/i.test(fileName)) {
    toast("File extension should be .xls, .xlsx, or .xlsm.");
    return;
  }
  await ensureXlsx();
  const workbook = window.XLSX.read(buffer, { type: "array", cellDates: true });
  state.fileName = fileName;
  state.imported = false;
  state.org = null;
  state.workbookSheets = workbook.SheetNames.map((name) => ({
    name,
    matrix: window.XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "", raw: true }),
  }));
  autoImport();
}

/** Run the checks on every tab, so the pickers can say which ones are usable. */
function checkSheets() {
  for (const sheet of state.workbookSheets || []) Object.assign(sheet, checkSheet(sheet.matrix));
}

/** Import the first tab that passes the checks (`prefer` first, if it still does). Nothing usable: stay blocked and say why. */
function autoImport(prefer = "") {
  checkSheets();
  const sheets = state.workbookSheets || [];
  const { usable, shown } = pickSheet(sheets, prefer);
  if (usable) {
    importSheet(usable.name, { checked: true });
    return;
  }
  state.sheetName = shown?.name || "";
  state.matrix = shown?.matrix || null;
  state.org = shown?.org || null;
  state.imported = false;
  state.category = "CEO";
  state.personKey = "";
  fillPeople();
  fillSheetPickers();
  showImportProblem();
  document.body.dataset.stage = "file";
  render();
  if (state.hydrating) return;
  toast(
    !sheets.length
      ? "This file has no sheets."
      : sheets.length > 1
        ? "None of the tabs in this file passes the checks. Pick another file, or fix the rows."
        : "This file does not pass the checks. Pick another file, or fix the rows.",
  );
}

/** Both pickers list every tab. Choosing a broken one in the ribbon opens it in Edit data instead of importing it. */
function fillSheetPickers() {
  const fill = (select, brokenLabel) => {
    select.innerHTML = "";
    for (const sheet of state.workbookSheets || []) {
      const option = document.createElement("option");
      option.value = sheet.name;
      option.textContent = sheet.problem ? `${sheet.name} (${brokenLabel})` : sheet.name;
      option.title = sheet.problem || "";
      select.appendChild(option);
    }
    select.value = state.sheetName;
  };
  fill($("sheet"), "has errors, opens Edit data");
  fill($("sheet-fix"), "has errors");
  const name = $("file-name");
  if (name) name.value = state.fileName || "";
}

/** `checked`: autoImport just ran the checks on every tab, so they are not run again. */
function importSheet(name = state.sheetName, { checked = false } = {}) {
  const sheet = state.workbookSheets?.find((item) => item.name === name);
  if (!sheet) {
    toast("Choose an Excel file first.");
    return;
  }
  // The matrix may have changed in Edit data, so every tab's label is refreshed too.
  if (!checked) checkSheets();
  state.sheetName = sheet.name;
  state.matrix = sheet.matrix;
  state.org = sheet.org;
  state.category = "CEO";
  state.personKey = state.org.ceo || state.org.order[0] || "";
  fillPeople();
  fillSheetPickers();
  showImportProblem();
  if (state.org.fatal) {
    state.imported = false;
    document.body.dataset.stage = "file";
    render();
    if (!state.hydrating) toast("This sheet has problems. The chart stays blocked until they are fixed.");
    return;
  }
  state.imported = true;
  document.body.dataset.stage = "ready";
  render();
  if (state.hydrating) return;
  const size = state.org.companySize === 1 ? "1 person" : `${state.org.companySize} people`;
  toast(`Imported ${state.sheetName}. CEO ${personFullName(state.org.people.get(state.org.ceo))}, ${size}.`);
}

/** A failed import is explained next to the sheet picker, where the next move (other sheet, other file, edit rows) is. */
function showImportProblem() {
  const box = $("import-problem");
  const fix = $("edit-data-fix");
  const problem = sheetProblem(state.org);
  box.hidden = !problem;
  const noneUsable = (state.workbookSheets?.length || 0) > 1 && state.workbookSheets.every((sheet) => sheet.problem);
  box.textContent = problem
    ? `${state.sheetName}: ${problem} ${noneUsable ? "No tab in this file passes the checks. " : ""}Pick another file, or fix the rows.`
    : "";
  fix.hidden = !problem || !sheetByName(state.sheetName);
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
    option.textContent = pickerLabel(org.people.get(key));
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
  const records = cardNoteRecords(state.slides, state.active);
  remember();
  const fresh = buildSlides(state.org, key, mode);
  for (const slide of fresh) placeCardNotes(slide.shapes, records);
  const deck = { id: nid("deck"), mode, rootKey: key };
  for (const slide of fresh) slide.deck = deck;
  state.slides = prependCharts(state.slides, fresh);
  state.active = 0;
  state.selection.clear();
  state.editingNote = null;
  state.renaming = null;
  if (state.costCenters) showCostCenterStamps(state.org, state.slides, nid);
  const hidden = hiddenTeams(state.slides[0]);
  render();
  const placed = fresh.length === 1 ? "Added 1 slide in front." : `Added ${fresh.length} slides in front.`;
  const fit = mode === "infinity" ? "" : hidden ? ` ${hidden} team${hidden === 1 ? "" : "s"} folded to fit.` : " Chart fits on one slide.";
  toast(`${placed}${fit}`);
  scheduleSave();
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
  stickNotes(slide.shapes);
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
        h: statsBoxHeight(statisticsLines(state.org.people.get(key), visibleReports(state.org, key).length)),
        z: 5,
        heading: "Department statistics",
        lines: statisticsLines(state.org.people.get(key), visibleReports(state.org, key).length),
      },
    );
    const hiddenList = hiddenListShape(state.org, key);
    if (hiddenList) slide.shapes.push({ ...hiddenList, id: nid("hiddenlist") });
  }
  freezeCards(state.org, slide.shapes);
  if (!slide.rootKey) {
    slide.rootKey = key;
    slide.rootLabel = displayName(state.org.people.get(key));
  }
  render();
}

function titleFor(key) {
  const person = state.org.people.get(key);
  const direct = visibleReports(state.org, key).length;
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
  if (visibleReports(state.org, key).length === 0) {
    toast(`${personFullName(person)} has no direct reports.`);
    return;
  }
  remember();
  const slide = current();
  slide.shapes = dropSubordinateShapes(state.org, slide.shapes, key);
  for (const id of [...state.selection]) {
    if (!slide.shapes.some((shape) => shape.id === id)) state.selection.delete(id);
  }
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
  freezeCards(state.org, slide.shapes);
  stickNotes(slide.shapes);
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
    slide.shapes.filter((shape) => shape.type === "wire"),
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
    ["ordinary", "important", "intern", "tbh"].forEach((placeholder, index) => {
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
    const noteIds = new Set(selectedShapes().filter((shape) => shape.type === "note").map((shape) => shape.id));
    const slideNow = current();
    for (const slide of state.slides) {
      slide.shapes = slide.shapes.filter((shape) => {
        if (slide === slideNow && state.selection.has(shape.id)) return false;
        if (shape.type === "note" && noteIds.has(shape.id)) return false;
        return true;
      });
    }
    state.selection.clear();
    if (noteIds.has(state.editingNote)) state.editingNote = null;
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
  anchorNotes(current().shapes, new Set(boxes.map((shape) => shape.id)));
  mirrorNotesFrom(current().shapes);
  render();
}

function nudge(dx, dy) {
  const shapes = selectedShapes();
  if (!shapes.length) return;
  remember();
  for (const shape of shapes) moveShape(shape, dx, dy);
  anchorNotes(current().shapes, new Set(shapes.map((shape) => shape.id)));
  mirrorNotesFrom(current().shapes);
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
      : "<div><strong>Empty slide</strong>Select an Excel file, then create the chart.</div>";
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
  const selected = selectedShapes();
  const bounds = selectionBounds(selected);
  if (bounds) root.appendChild(selectionFrame(bounds, selected));
}

function selectionFrame(bounds, shapes = []) {
  const frame = document.createElement("div");
  frame.className = "selection-frame";
  frame.style.left = `${viewX(bounds.x)}px`;
  frame.style.top = `${viewY(bounds.y)}px`;
  frame.style.width = `${bounds.w}px`;
  frame.style.height = `${bounds.h}px`;
  const shadesOnly = shapes.length > 0 && shapes.every((shape) => shape.type === "shade");
  if (shadesOnly) {
    for (const edge of ["n", "s", "e", "w", "nw", "ne", "sw", "se"]) frame.appendChild(resizeHandle(edge));
    return frame;
  }
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

function resizeHandle(edge) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `resize-handle resize-${edge}`;
  button.title = "Drag to resize the department";
  button.setAttribute("aria-label", `Resize department ${edge}`);
  const unit = 1 / Math.max(0.4, baseScale() * state.zoom);
  const across = edge === "n" || edge === "s" || edge === "e" || edge === "w";
  const along = edge === "n" || edge === "s";
  button.style.width = `${(along ? 28 : across ? 10 : 12) * unit}px`;
  button.style.height = `${(edge === "e" || edge === "w" ? 28 : across ? 10 : 12) * unit}px`;
  const spot = {
    n: ["50%", "0"],
    s: ["50%", "100%"],
    e: ["100%", "50%"],
    w: ["0", "50%"],
    nw: ["0", "0"],
    ne: ["100%", "0"],
    sw: ["0", "100%"],
    se: ["100%", "100%"],
  }[edge];
  button.style.left = spot[0];
  button.style.top = spot[1];
  button.addEventListener("pointerdown", (event) => beginResize(event, edge));
  return button;
}

function renderShape(shape) {
  const editing = shape.type === "note" && state.editingNote === shape.id;
  const node = editing ? document.createElement("div") : document.createElement("button");
  if (!editing) node.type = "button";
  node.className = `shape ${shape.type}`;
  node.dataset.id = shape.id;
  node.style.left = `${viewX(shape.x)}px`;
  node.style.top = `${viewY(shape.y)}px`;
  node.style.width = `${shape.w}px`;
  node.style.height = `${shape.h}px`;
  node.style.zIndex = String(shape.z ?? 2);
  if (state.selection.has(shape.id)) node.classList.add("selected");

  if (shape.type === "card") {
    const person = shapePerson(state.org, shape);
    const kind = person ? cardStyle(person) : shape.placeholder || "ordinary";
    node.classList.add(kind);
    if (person?.tbh || (!person && kind === "tbh")) node.classList.add("vacant");
    const lines = person
      ? person.tbh
        ? shape.h >= 32
          ? ["To be hired", "", moderated(person.job)]
          : ["To be hired", moderated(person.job)]
        : [moderated(person.surname), moderated(person.name), moderated(person.job)]
      : shape.lines || [];
    const spec = CARDS[shape.size] || CARDS.Medium;
    const fontScale = spec.h ? shape.h / spec.h : 1;
    node.style.fontSize = `${spec.name * fontScale}px`;
    node.style.setProperty("--padx", `${(spec.name * fontScale * 0.65).toFixed(1)}px`);
    const badge = person ? headcountLabel(person) : "";
    const footer = person?.location || "";
    node.classList.toggle("no-loc", !footer);
    const badgeHtml = badge ? `<span class="hc">${escapeHtml(badge)}</span>` : "";
    node.innerHTML = `<span class="who">${escapeHtml(lines[0] || "")}</span>
      <span class="who">${escapeHtml(lines[1] || "")}</span>
      ${shape.h >= 32 && lines[2] ? `<span class="job" style="font-size:${spec.title * fontScale}px">${escapeHtml(lines[2])}</span>` : ""}
      ${footer ? `<span class="loc"><span class="loc-text">${escapeHtml(footer)}</span>${badgeHtml}</span>` : badge ? `<span class="hc corner">${escapeHtml(badge)}</span>` : ""}${person?.newcomer ? `<span class="new-badge" title="Newcomer${person.newcomerDate ? ` from ${escapeHtml(person.newcomerDate)}` : ""}">New${person.newcomerDate ? `<small>${escapeHtml(person.newcomerDate)}</small>` : ""}</span>` : ""}`;
    node.setAttribute("aria-label", person ? personFullName(person) : "Empty card");
    node.addEventListener("dblclick", (event) => {
      event.stopPropagation();
      focusPerson(shape.key);
    });
  } else if (shape.type === "title") {
    node.innerHTML = `<div class="name">${escapeHtml(shape.name || "")}</div>
      <div class="dept">${escapeHtml(shape.department || "Organisation")}</div>`;
  } else if (shape.type === "legend") {
    node.innerHTML = `<div class="legend"><div class="legend-title">Indicators</div><div><i class="ordinary"></i>Ordinary employee</div><div><i class="important"></i>Important employee</div><div><i class="intern"></i>Intern</div></div>`;
  } else if (shape.type === "hiddenlist") {
    node.textContent = shape.text || "";
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
    node.style.fontSize = `${shape.fs || Math.max(5, (shape.h || 11) * 0.64)}px`;
  } else if (shape.type === "comment" || shape.type === "link" || shape.type === "nav") {
    node.textContent = shape.text || "";
    if (shape.type === "link") {
      node.style.background = LINK_STYLE.fill;
      node.style.color = LINK_STYLE.color;
    }
  } else if (shape.type === "note") {
    const card = cardForNoteShape(shape);
    const fs = noteTypeSize(card);
    shape.fs = fs;
    node.style.fontSize = `${fs}px`;
    node.style.background = shape.fill || NOTE_COLORS[0];
    if (editing) {
      const input = document.createElement("textarea");
      input.className = "note-input";
      input.value = shape.text || "";
      input.placeholder = "Note";
      input.setAttribute("aria-label", "Note");
      input.addEventListener("pointerdown", (event) => event.stopPropagation());
      input.addEventListener("input", () => onNoteInput(shape, input.value, node));
      input.addEventListener("blur", () => {
        if (state.editingNote !== shape.id) return;
        state.editingNote = null;
        // Rebuild after the click that left the note, unless that click is still dragging.
        setTimeout(() => {
          if (pointer.id != null || state.editingNote) return;
          render();
        }, 0);
      });
      input.addEventListener("keydown", (event) => {
        event.stopPropagation();
        if (event.key === "Escape") {
          event.preventDefault();
          state.editingNote = null;
          render();
        }
      });
      node.appendChild(input);
      queueMicrotask(() => {
        input.focus();
        const end = input.value.length;
        input.setSelectionRange(end, end);
      });
    } else {
      node.textContent = shape.text || "Note";
      node.classList.toggle("empty", !shape.text);
      node.setAttribute("aria-label", shape.text ? `Note: ${shape.text}` : "Empty note");
      node.addEventListener("dblclick", (event) => {
        event.stopPropagation();
        event.preventDefault();
        beginNoteEdit(shape);
      });
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

function beginNoteEdit(shape) {
  state.editingNote = shape.id;
  state.noteEditReady = false;
  state.selection.clear();
  state.selection.add(shape.id);
  render();
}

function onNoteInput(shape, value, node) {
  if (!state.noteEditReady) {
    remember();
    state.noteEditReady = true;
  }
  const height = noteHeightFor(value, shape.w, shape.fs || noteTypeSize(cardForNoteShape(shape)));
  for (const slide of state.slides) {
    for (const item of slide.shapes) {
      if (item.type !== "note" || item.id !== shape.id) continue;
      item.text = value;
      item.h = height;
    }
  }
  node.style.height = `${height}px`;
  scheduleSave();
}

/** Keep the same note, on every slide it was copied to, in step with the one being edited. */
function mirrorNotesFrom(shapes) {
  const notes = (shapes || []).filter((shape) => shape.type === "note" && shape.id);
  if (!notes.length) return;
  for (const slide of state.slides) {
    if (slide.shapes === shapes) continue;
    let touched = false;
    for (const source of notes) {
      for (const shape of slide.shapes) {
        if (shape.type !== "note" || shape.id !== source.id) continue;
        shape.text = source.text;
        shape.fill = source.fill;
        shape.dx = source.dx;
        shape.dy = source.dy;
        shape.w = source.w;
        shape.h = source.h;
        shape.fs = source.fs;
        touched = true;
      }
    }
    if (touched) stickNotes(slide.shapes);
  }
}

function addCardNote(card) {
  const slide = current();
  if (!slide || !card) return;
  remember();
  const index = slide.shapes.filter((shape) => {
    if (shape.type !== "note") return false;
    if (shape.cardId) return shape.cardId === card.id;
    return Boolean(shape.key && shape.key === card.key);
  }).length;
  const box = noteBox(card, index);
  const note = {
    id: nid("note"),
    type: "note",
    key: card.key || "",
    cardId: card.id,
    text: "",
    fill: NOTE_COLORS[index % NOTE_COLORS.length],
    z: 8,
    ...box,
  };
  slide.shapes.push(note);
  state.selection.clear();
  state.selection.add(note.id);
  state.editingNote = note.id;
  state.noteEditReady = false;
  render();
}

function cardForNoteShape(shape) {
  const slide = current();
  if (!slide || shape?.type !== "note") return null;
  if (shape.cardId) {
    const exact = slide.shapes.find((item) => item.type === "card" && item.id === shape.cardId);
    if (exact) return exact;
  }
  return slide.shapes.find((item) => item.type === "card" && shape.key && item.key === shape.key) || null;
}

function noteCardFor(shape) {
  const slide = current();
  if (!slide) return null;
  if (shape?.type === "card") return shape;
  if (shape?.type === "note") {
    const linked = slide.shapes.find((item) => item.type === "card" && (item.id === shape.cardId || (shape.key && item.key === shape.key)));
    if (linked) return linked;
  }
  return slide.shapes.find((item) => state.selection.has(item.id) && item.type === "card") || null;
}

function closeContextMenu() {
  const menu = $("context-menu");
  if (menu) menu.hidden = true;
}

function openContextMenu(card, note, x, y) {
  const add = $("context-add-note");
  const remove = $("context-delete-note");
  add.hidden = !card;
  remove.hidden = !note;
  if (add.hidden && remove.hidden) return;
  closeMenus();
  state.noteTargetId = card?.id || "";
  state.deleteNoteId = note?.id || "";
  const menu = $("context-menu");
  menu.hidden = false;
  menu.style.left = `${x}px`;
  menu.style.top = `${y}px`;
  const rect = menu.getBoundingClientRect();
  if (rect.right > window.innerWidth - 4) menu.style.left = `${Math.max(4, x - rect.width)}px`;
  if (rect.bottom > window.innerHeight - 4) menu.style.top = `${Math.max(4, y - rect.height)}px`;
}

function onSlideContextMenu(event) {
  const slide = current();
  if (!slide) return;
  const host = event.target.closest?.("[data-id]");
  const shape = host ? slide.shapes.find((item) => item.id === host.dataset.id) : null;
  if (shape?.type === "note") {
    state.selection.clear();
    state.selection.add(shape.id);
    state.editingNote = null;
    if (shape.key) focusPerson(shape.key);
    render();
  } else if (shape?.type === "card" && !state.selection.has(shape.id)) {
    state.selection.clear();
    selectGroup(shape, false);
    if (shape.key) focusPerson(shape.key);
    state.editingNote = null;
    render();
  }
  const note = shape?.type === "note"
    ? shape
    : shape
      ? null
      : slide.shapes.find((item) => state.selection.has(item.id) && item.type === "note") || null;
  const card = noteCardFor(shape);
  if (!card && !note) return;
  event.preventDefault();
  openContextMenu(card, note, event.clientX, event.clientY);
}

function focusPerson(key) {
  if (!key || !state.org?.people.has(key)) return;
  state.personKey = key;
  const wanted = state.org.people.get(key).tbh ? VACANCIES : key.slice(0, 1).toUpperCase();
  // A letter with no category of its own (a digit, say) still lands on the person, under All, A–Z.
  state.category = categories(state.org).includes(wanted) ? wanted : "AZ";
  fillPeople();
  $("employee").value = key;
}

const pointer = { id: null, x: 0, y: 0, moved: false, shape: null, marquee: null, scale: null, resize: null };

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
  if (event.button !== 0) return;
  if (state.editingNote && state.editingNote !== shape?.id) state.editingNote = null;
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
  pointer.resize = null;
  if (shape && !state.selection.has(shape.id)) selectGroup(shape, event.shiftKey);
  if ((shape?.type === "card" || shape?.type === "note") && shape.key) focusPerson(shape.key);
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
  pointer.resize = null;
  pointer.marquee?.remove();
  pointer.marqueeShape = shape || null;
  pointer.marqueeAdditive = event.shiftKey;
  if (!event.shiftKey && !shape) state.selection.clear();
  pointer.marquee = document.createElement("div");
  pointer.marquee.className = "marquee";
  $("slide").appendChild(pointer.marquee);
  $("slide").setPointerCapture?.(event.pointerId);
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
  pointer.resize = null;
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

function beginResize(event, edge) {
  event.stopPropagation();
  event.preventDefault();
  const live = selectedShapes().filter((shape) => shape.type === "shade");
  if (!live.length) return;
  const point = localPoint(event);
  pointer.marquee?.remove();
  pointer.id = event.pointerId;
  pointer.moved = false;
  pointer.shape = null;
  pointer.marquee = null;
  pointer.marqueeShape = null;
  pointer.scale = null;
  pointer.resize = {
    edge,
    anchor: point,
    snapshot: live.map((shape) => ({ ...shape })),
    live,
    remembered: false,
  };
  event.currentTarget.setPointerCapture?.(event.pointerId);
}

function onHolderPointerDown(event) {
  if (event.button !== 0) return;
  if (event.target.closest?.(".shape, .scale-handle, .resize-handle, .wires line")) return;
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

function applyResizeGesture(point) {
  const gesture = pointer.resize;
  const snapshot = gesture.snapshot.map((shape) => ({ ...shape }));
  resizeShapes(snapshot, gesture.edge, point.x - gesture.anchor.x, point.y - gesture.anchor.y);
  gesture.live.forEach((shape, index) => {
    const next = snapshot[index];
    shape.x = next.x;
    shape.y = next.y;
    shape.w = next.w;
    shape.h = next.h;
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
  anchorNotes(current().shapes, new Set(gesture.live.map((shape) => shape.id)));
  for (const shape of current().shapes) if (shape.type === "note") paintShape(shape);
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
  const anchorX = pointer.scale?.anchor.x ?? pointer.resize?.anchor.x ?? pointer.x;
  const anchorY = pointer.scale?.anchor.y ?? pointer.resize?.anchor.y ?? pointer.y;
  if (!pointer.moved && Math.hypot(point.x - anchorX, point.y - anchorY) < 2) return;
  if (pointer.resize) {
    if (!pointer.resize.remembered) {
      remember();
      pointer.resize.remembered = true;
    }
    pointer.moved = true;
    applyResizeGesture(point);
    return;
  }
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
  anchorNotes(slide.shapes, new Set(moving.map((shape) => shape.id)));
  pointer.x = point.x;
  pointer.y = point.y;
  for (const shape of moving) paintShape(shape);
  for (const shape of slide.shapes) if (shape.type === "note") paintShape(shape);
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
  const moved = pointer.moved;
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
    if (state.editingNote && !state.selection.has(state.editingNote)) state.editingNote = null;
    pointer.marquee.remove();
    pointer.marquee = null;
    pointer.marqueeShape = null;
    render();
  } else if (pointer.scale || pointer.resize) {
    pointer.scale = null;
    pointer.resize = null;
    render();
  } else if (!moved && pointer.shape?.type === "note" && state.editingNote !== pointer.shape.id && state.tool === "move") {
    beginNoteEdit(pointer.shape);
  }
  if (moved) mirrorNotesFrom(current()?.shapes || []);
  pointer.id = null;
  pointer.shape = null;
}

function slideCaption(slide, index) {
  const cards = slide.shapes.filter((shape) => shape.type === "card").length;
  return {
    title: slide.title || `Slide ${index + 1}`,
    detail: slide.rootLabel || (cards ? `${cards} cards` : "Empty"),
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

/** People the slides leave out, once there is more than one slide (a single slide may fold teams on purpose). */
function missingFromSlides() {
  if (!state.org || state.org.fatal || state.slides.length < 2) return 0;
  const drawn = new Set();
  for (const slide of state.slides) {
    for (const shape of slide.shapes) if (shape.type === "card" && shape.key) drawn.add(shape.key);
  }
  if (!drawn.size) return 0;
  return [...state.org.people.keys()].filter((key) => !drawn.has(key) && !state.org.people.get(key).hidden).length;
}

function renderStatus() {
  const fit = current()?.fit;
  const fitText = fit ? `${fit.size} · tick ${Math.round(fit.gap)} pt` : "";
  const missing = missingFromSlides();
  const left = $("status-left");
  const saved = state.account && state.saveLabel && state.fileName ? ` · ${state.saveLabel}` : "";
  left.textContent = state.org
    ? `${state.fileName || "Workbook"} · ${state.sheetName} · ${state.org.companySize} people${fitText ? ` · ${fitText}` : ""}${missing ? ` · ${missing} not on any slide` : ""}${saved}`
    : state.fileName
      ? `${state.fileName}${saved}`
      : "No workbook yet.";
  left.classList.toggle("warn", missing > 0);
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
  const width = `${frame.w * scale}px`;
  const height = `${frame.h * scale}px`;
  const transform = `scale(${scale})`;
  if (holder.style.width === width && holder.style.height === height && $("scaler").style.transform === transform) return;
  holder.style.width = width;
  holder.style.height = height;
  const scaler = $("scaler");
  scaler.style.width = `${frame.w}px`;
  scaler.style.height = `${frame.h}px`;
  scaler.style.transform = transform;
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

/** Copy the selected shapes' text, or the whole slide's when nothing is selected. */
async function copySlideText() {
  const shapes = state.selection.size ? selectedShapes() : current().shapes;
  const text = slideMarkdown(state.org, shapes, { number: state.active + 1, total: state.slides.length });
  if (!text) {
    toast("Nothing to copy on this slide.");
    return;
  }
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const area = document.createElement("textarea");
    area.value = text;
    area.style.cssText = "position:fixed;opacity:0";
    document.body.appendChild(area);
    area.select();
    document.execCommand("copy");
    area.remove();
  }
  toast(`Copied the ${state.selection.size ? "selection" : "slide"} as Markdown.`);
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
    await ensurePptx();
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

// ---- Sheet editor -------------------------------------------------------

const sheetEdit = { sheetName: "", draft: [], firstRow: 1, hasTitles: true, dirty: false, org: null, timer: 0, sort: null };
const FLAG_COLUMNS = new Set([10, 11, 14, 15]);

function sheetByName(name) {
  return state.workbookSheets?.find((item) => item.name === name);
}

/** Edit any tab, including one that does not pass the checks yet. Defaults to the imported one. */
function openSheetEditor(name = state.sheetName) {
  const sheet = sheetByName(typeof name === "string" ? name : state.sheetName);
  if (!sheet) {
    toast("Import a sheet first.");
    return;
  }
  sheetEdit.sheetName = sheet.name;
  sheetEdit.hasTitles = sheet.hasTitles !== false;
  sheetEdit.firstRow = sheetEdit.hasTitles ? 1 : 0;
  sheetEdit.draft = trimTrailingBlankRows(draftFromMatrix(sheet.matrix), sheetEdit.firstRow);
  if (sheetEdit.hasTitles && !sheetEdit.draft.length) sheetEdit.draft.push(COLUMN_LABELS.slice());
  sheetEdit.dirty = false;
  sheetEdit.sort = null;
  syncSortControls();
  $("sheet-filter").value = "";
  $("sheet-sub").textContent = `${state.fileName || "Workbook"} · ${sheet.name}`;
  $("sheet-dialog").hidden = false;
  document.body.classList.add("sheet-open");
  renderSheetGrid();
  validateSheet();
}

function closeSheetEditor() {
  clearTimeout(sheetEdit.timer);
  $("sheet").value = state.sheetName; // a broken tab picked in the ribbon only opened the editor
  $("sheet-dialog").hidden = true;
  document.body.classList.remove("sheet-open");
}

function renderSheetGrid() {
  const grid = $("sheet-grid");
  const sort = sheetEdit.sort;
  const head = `<thead><tr><th class="num">#</th>${COLUMN_LABELS.map((label, c) => {
    const dir = sort && sort.col === c ? sort.dir : "";
    const arrow = dir === "asc" ? "▲" : dir === "desc" ? "▼" : "⇅";
    const hint = dir === "asc" ? "Click for Z–A" : dir === "desc" ? "Click to restore sheet order" : "Click to sort A–Z";
    return `<th class="sortable" data-sort="${c}" tabindex="0" role="button" aria-sort="${dir === "asc" ? "ascending" : dir === "desc" ? "descending" : "none"}" title="Column ${String.fromCharCode(65 + c)}: ${escapeHtml(label)}. ${hint}">${escapeHtml(label)}<span class="sort-ind${dir ? " on" : ""}">${arrow}</span></th>`;
  }).join("")}<th></th></tr></thead>`;
  const rows = [];
  for (const r of sheetRowOrder()) {
    const row = sheetEdit.draft[r];
    const cells = COLUMN_LABELS.map((label, c) => {
      if (FLAG_COLUMNS.has(c)) {
        return `<td class="flag"><input type="checkbox" data-c="${c}" aria-label="${escapeHtml(label)}"${isFlagOn(row[c]) ? " checked" : ""}></td>`;
      }
      return `<td><input type="text" data-c="${c}" aria-label="${escapeHtml(label)}" value="${escapeHtml(cellText(row[c]))}"></td>`;
    }).join("");
    rows.push(
      `<tr data-r="${r}"><td class="num">${r + 1}</td>${cells}<td class="del"><button type="button" class="row-del" aria-label="Delete row ${r + 1}" title="Delete row">×</button></td></tr>`,
    );
  }
  grid.innerHTML = `${head}<tbody>${rows.join("")}</tbody>`;
  applySheetFilter();
}

/** Draft row indexes in display order. Sorting is a view only: the sheet order and row numbers do not change. */
function sheetRowOrder() {
  const order = [];
  for (let r = sheetEdit.firstRow; r < sheetEdit.draft.length; r += 1) order.push(r);
  const sort = sheetEdit.sort;
  if (!sort) return order;
  const value = (r) => {
    const cell = sheetEdit.draft[r][sort.col];
    return FLAG_COLUMNS.has(sort.col) ? (isFlagOn(cell) ? "1" : "0") : cellText(cell).trim();
  };
  const sign = sort.dir === "asc" ? 1 : -1;
  // Array.prototype.sort is stable, so equal cells keep their sheet order. Blanks stay at the bottom both ways.
  return order.sort((a, b) => {
    const left = value(a);
    const right = value(b);
    if (!left && !right) return 0;
    if (!left) return 1;
    if (!right) return -1;
    return sign * left.localeCompare(right, undefined, { numeric: true, sensitivity: "base" });
  });
}

function cycleSheetSort(col) {
  const sort = sheetEdit.sort;
  if (!sort || sort.col !== col) setSheetSort(col, "asc");
  else if (sort.dir === "asc") setSheetSort(col, "desc");
  else setSheetSort(null);
}

/** col null = back to the order in the sheet. */
function setSheetSort(col, dir) {
  sheetEdit.sort = col == null ? null : { col, dir };
  syncSortControls();
  const wrap = $("sheet-wrap");
  const left = wrap.scrollLeft;
  renderSheetGrid();
  validateSheet();
  wrap.scrollLeft = left;
  wrap.scrollTop = 0;
}

function syncSortControls() {
  const sort = sheetEdit.sort;
  $("sheet-sort-col").value = sort ? String(sort.col) : "";
  $("sheet-sort-asc").classList.toggle("on", sort?.dir === "asc");
  $("sheet-sort-desc").classList.toggle("on", sort?.dir === "desc");
  $("sheet-sort-reset").disabled = !sort;
}

function applySheetFilter() {
  const needle = $("sheet-filter").value.trim().toLowerCase();
  for (const tr of $("sheet-grid").tBodies[0]?.rows || []) {
    const row = sheetEdit.draft[Number(tr.dataset.r)] || [];
    tr.hidden = !!needle && !row.some((cell) => cellText(cell).toLowerCase().includes(needle));
  }
}

function validateSheet() {
  const org = buildOrg(sheetEdit.draft, { hasTitles: sheetEdit.hasTitles });
  sheetEdit.org = org;
  const bad = problemRows(org);
  for (const tr of $("sheet-grid").tBodies[0]?.rows || []) {
    const messages = bad.get(Number(tr.dataset.r) + 1);
    tr.classList.toggle("bad", !!messages);
    if (messages) tr.title = messages.join(" ");
    else tr.removeAttribute("title");
  }
  const box = $("sheet-issues");
  if (org.fatal) {
    box.innerHTML = `<ul>${org.issues.map((issue) => `<li class="issue">${escapeHtml(issue.message)}</li>`).join("")}</ul>`;
  } else {
    const ceo = org.people.get(org.ceo);
    const notes = org.issues.map((issue) => `<li>${escapeHtml(issue.message)}</li>`).join("");
    box.innerHTML = `<p class="ok">${org.companySize} people${ceo ? ` · CEO ${escapeHtml(personFullName(ceo))}` : ""}. Ready to apply.</p>${notes ? `<ul>${notes}</ul>` : ""}`;
  }
  $("sheet-apply").disabled = org.fatal;
}

function scheduleSheetValidation() {
  sheetEdit.dirty = true;
  clearTimeout(sheetEdit.timer);
  sheetEdit.timer = setTimeout(validateSheet, 250);
}

function addSheetRow(row) {
  sheetEdit.draft.push(row);
  $("sheet-filter").value = "";
  renderSheetGrid();
  validateSheet();
  sheetEdit.dirty = true;
  const last = $("sheet-grid").tBodies[0].querySelector(`tr[data-r="${sheetEdit.draft.length - 1}"]`);
  last?.scrollIntoView({ block: "nearest" });
  const input = last?.querySelector('input[data-c="0"]');
  if (input) {
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);
  }
}

function onSheetInput(event) {
  const input = event.target;
  const tr = input.closest("tr[data-r]");
  if (!tr || input.dataset.c === undefined) return;
  const c = Number(input.dataset.c);
  sheetEdit.draft[Number(tr.dataset.r)][c] = input.type === "checkbox" ? flagValue(input.checked) : input.value;
  scheduleSheetValidation();
}

function onSheetClick(event) {
  const header = event.target.closest("th[data-sort]");
  if (header) {
    cycleSheetSort(Number(header.dataset.sort));
    return;
  }
  const button = event.target.closest(".row-del");
  if (!button) return;
  const r = Number(button.closest("tr").dataset.r);
  const wrap = $("sheet-wrap");
  const top = wrap.scrollTop;
  sheetEdit.draft.splice(r, 1);
  renderSheetGrid();
  validateSheet();
  sheetEdit.dirty = true;
  wrap.scrollTop = top;
}

function onSheetKeydown(event) {
  if ((event.key === "Enter" || event.key === " ") && event.target.matches?.("th[data-sort]")) {
    event.preventDefault();
    cycleSheetSort(Number(event.target.dataset.sort));
    return;
  }
  if (event.key === "Escape") {
    event.stopPropagation();
    if (sheetEdit.dirty) toast("You have unsaved edits. Use Apply or Cancel.");
    else closeSheetEditor();
    return;
  }
  if (event.key !== "Enter" || event.target.type !== "text") return;
  event.preventDefault();
  const c = event.target.dataset.c;
  const step = event.shiftKey ? "previousElementSibling" : "nextElementSibling";
  let tr = event.target.closest("tr")[step];
  while (tr && tr.hidden) tr = tr[step];
  tr?.querySelector(`input[data-c="${c}"]`)?.focus();
}

function applySheetEdits() {
  validateSheet();
  if (sheetEdit.org.fatal) return;
  const sheet = sheetByName(sheetEdit.sheetName);
  if (!sheet) return;
  sheet.matrix = trimTrailingBlankRows(sheetEdit.draft, sheetEdit.firstRow);
  const sameSheet = sheet.name === state.sheetName;
  const keep = { category: state.category, personKey: state.personKey };
  closeSheetEditor();
  importSheet(sheet.name);
  if (state.org && !state.org.fatal) {
    if (sameSheet) {
      state.category = keep.category;
      state.personKey = keep.personKey;
      fillPeople();
    }
    render();
    const size = state.org.companySize === 1 ? "1 person" : `${state.org.companySize} people`;
    const note = state.slides.some((slide) => slide.shapes.some((shape) => shape.type === "card"))
      ? " Existing slides are unchanged: click Create org chart for new ones."
      : "";
    toast(`Data applied, ${size}.${note}`);
  }
  rewriteFileBytes()
    .then(() => enqueueSave(() => saveWork({ file: true, deck: true })))
    .catch((error) => toast(error.message || "Could not save."));
}

async function downloadSheetXlsx() {
  await ensureXlsx();
  const XLSX = window.XLSX;
  const workbook = XLSX.utils.book_new();
  for (const sheet of state.workbookSheets || []) {
    const rows =
      sheet.name === sheetEdit.sheetName ? trimTrailingBlankRows(sheetEdit.draft, sheetEdit.firstRow) : sheet.matrix;
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows, { cellDates: true }), sheet.name);
  }
  const base = (state.fileName || "MyBoss-data").replace(/\.(xlsx|xls|xlsm)$/i, "");
  XLSX.writeFile(workbook, `${base}-edited.xlsx`);
}

function bindSheetEditor() {
  $("edit-data").addEventListener("click", openSheetEditor);
  $("sheet-cancel").addEventListener("click", closeSheetEditor);
  $("sheet-apply").addEventListener("click", applySheetEdits);
  $("sheet-download").addEventListener("click", downloadSheetXlsx);
  $("sheet-add").addEventListener("click", () => addSheetRow(blankDraftRow()));
  $("sheet-add-tbh").addEventListener("click", () => addSheetRow(vacancyDraftRow()));
  $("sheet-filter").addEventListener("input", applySheetFilter);
  const sortSelect = $("sheet-sort-col");
  sortSelect.innerHTML =
    '<option value="">(sheet order)</option>' +
    COLUMN_LABELS.map((label, c) => `<option value="${c}">${String.fromCharCode(65 + c)} · ${escapeHtml(label)}</option>`).join("");
  const sortFromControls = (dir) => {
    const picked = sortSelect.value;
    if (picked === "") {
      if (dir) toast("Pick a column to sort by.");
      return;
    }
    setSheetSort(Number(picked), dir || sheetEdit.sort?.dir || "asc");
  };
  sortSelect.addEventListener("change", () => (sortSelect.value === "" ? setSheetSort(null) : sortFromControls()));
  $("sheet-sort-asc").addEventListener("click", () => sortFromControls("asc"));
  $("sheet-sort-desc").addEventListener("click", () => sortFromControls("desc"));
  $("sheet-sort-reset").addEventListener("click", () => setSheetSort(null));
  const grid = $("sheet-grid");
  grid.addEventListener("input", onSheetInput);
  grid.addEventListener("change", onSheetInput);
  grid.addEventListener("click", onSheetClick);
  $("sheet-dialog").addEventListener("keydown", onSheetKeydown);
}

function onKey(event) {
  if (document.body.classList.contains("sheet-open") || document.body.classList.contains("account-open")) return;
  const typing = ["INPUT", "SELECT", "TEXTAREA"].includes(event.target.tagName);
  if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "z") {
    event.preventDefault();
    if (event.shiftKey) restore(state.redo, state.undo);
    else restore(state.undo, state.redo);
    return;
  }
  if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "c" && !typing && !window.getSelection()?.toString()) {
    event.preventDefault();
    copySlideText();
    return;
  }
  if (typing) return;
  if (event.key === "Delete" || event.key === "Backspace") {
    event.preventDefault();
    arrange("delete");
  } else if (event.key === "Escape") {
    if (!$("context-menu").hidden) {
      closeContextMenu();
      return;
    }
    state.editingNote = null;
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

function closeMenus() {
  document.querySelectorAll(".menu.open").forEach((node) => {
    node.classList.remove("open");
    node.style.position = "";
    node.style.top = "";
    node.style.left = "";
    node.style.zIndex = "";
  });
}

function toggleMenu(menuId, event) {
  event.stopPropagation();
  const menu = $(menuId);
  const willOpen = !menu.classList.contains("open");
  closeMenus();
  if (!willOpen) return;
  menu.classList.add("open");
  // The tool strip scrolls, so a menu anchored inside it would be clipped.
  // Pin it to the screen at the button that opened it.
  const host = menu.parentElement.getBoundingClientRect();
  const left = Math.max(8, Math.min(host.left, window.innerWidth - menu.offsetWidth - 8));
  menu.style.position = "fixed";
  menu.style.top = `${host.bottom - 2}px`;
  menu.style.left = `${left}px`;
  menu.style.zIndex = "40";
}

function bind() {
  $("open-file").addEventListener("click", () => $("file").click());
  $("pick-again").addEventListener("click", () => $("file").click());
  $("file").addEventListener("change", async () => {
    const file = $("file").files[0];
    if (!file) return;
    try {
      const bytes = await file.arrayBuffer();
      state.fileBytes = bytes;
      state.workId = null;
      state.saveLabel = "";
      await readWorkbook(bytes, file.name);
      enqueueSave(() => saveWork({ file: true })).catch((error) => toast(error.message || "Could not save."));
    } catch (error) {
      toast(`Could not read ${file.name}: ${error.message || "not an Excel file"}.`);
    } finally {
      $("file").value = ""; // so choosing the same file again, after fixing it, still fires "change"
    }
  });
  $("refresh").addEventListener("click", () => $("file").click());
  $("pick-sheet").addEventListener("click", () => {
    const select = $("sheet");
    select.focus();
    try {
      select.showPicker?.();
    } catch {
      // Older browsers refuse to open a <select> from script; focus is enough there.
    }
  });
  $("refresh-more").addEventListener("click", (event) => {
    toggleMenu("menu-refresh", event);
    if ($("menu-refresh").classList.contains("open")) fillRefreshSavedMenu();
  });
  $("create-main").addEventListener("click", (event) => toggleMenu("menu-create", event));
  $("manager-more").addEventListener("click", (event) => toggleMenu("menu-manager", event));
  $("reports-more").addEventListener("click", (event) => toggleMenu("menu-reports", event));
  $("elements-main").addEventListener("click", (event) => toggleMenu("menu-elements", event));
  $("arrange-main").addEventListener("click", (event) => toggleMenu("menu-arrange", event));
  document.addEventListener("click", () => {
    closeMenus();
    closeContextMenu();
  });
  $("context-add-note").addEventListener("click", (event) => {
    event.stopPropagation();
    const card = current()?.shapes.find((shape) => shape.id === state.noteTargetId && shape.type === "card");
    closeContextMenu();
    if (card) addCardNote(card);
  });
  $("context-delete-note").addEventListener("click", (event) => {
    event.stopPropagation();
    const id = state.deleteNoteId;
    closeContextMenu();
    if (!id) return;
    state.editingNote = null;
    state.selection.clear();
    state.selection.add(id);
    arrange("delete");
  });
  $("slide").addEventListener("contextmenu", onSlideContextMenu);
  $("sheet").addEventListener("change", () => {
    const sheet = sheetByName($("sheet").value);
    if (!sheet) return;
    if (sheet.problem) {
      openSheetEditor(sheet.name);
      toast(`${sheet.name}: ${sheet.problem} Fix it here, then Apply.`);
      return;
    }
    importSheet(sheet.name);
  });
  $("sheet-fix").addEventListener("change", () => {
    const sheet = state.workbookSheets?.find((item) => item.name === $("sheet-fix").value);
    if (!sheet) return;
    state.sheetName = sheet.name;
    state.matrix = sheet.matrix;
    state.org = sheet.org;
    showImportProblem();
    render();
  });

  $("pick-other-file").addEventListener("click", () => $("file").click());
  $("edit-data-fix").addEventListener("click", openSheetEditor);
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
  $("copy-text").addEventListener("click", copySlideText);
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
  $("account").addEventListener("click", () => openAccount());
  $("account-form").addEventListener("submit", (event) => {
    event.preventDefault();
    submitAccount(apiUrl("/api/login"));
  });
  $("account-register").addEventListener("click", () => submitAccount(apiUrl("/api/register")));
  $("account-logout").addEventListener("click", logoutAccount);
  $("account-close").addEventListener("click", closeAccount);
  $("holder").addEventListener("pointerdown", onHolderPointerDown);
  window.addEventListener("pointermove", onPointerMove);
  window.addEventListener("pointerup", onPointerUp);
  window.addEventListener("keydown", onKey);
  window.addEventListener("resize", applyScale);
  window.addEventListener("pagehide", flushSave);
  if (window.ResizeObserver) new ResizeObserver(() => applyScale()).observe($("stage"));
  requestAnimationFrame(() => applyScale());
}

function apiUrl(path) {
  const here = window.location.pathname || "";
  const prefix = here === "/myboss" || here.startsWith("/myboss/") ? "/myboss" : "";
  return prefix + path;
}

function bytesToBase64(bytes) {
  const view = bytes instanceof ArrayBuffer ? new Uint8Array(bytes) : new Uint8Array(bytes);
  let binary = "";
  const size = 0x8000;
  for (let index = 0; index < view.length; index += size) {
    binary += String.fromCharCode(...view.subarray(index, index + size));
  }
  return btoa(binary);
}

function base64ToBytes(value) {
  const binary = atob(value);
  const out = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) out[index] = binary.charCodeAt(index);
  return out;
}

async function rewriteFileBytes() {
  if (!state.workbookSheets?.length) return;
  await ensureXlsx();
  const workbook = window.XLSX.utils.book_new();
  for (const sheet of state.workbookSheets) {
    window.XLSX.utils.book_append_sheet(
      workbook,
      window.XLSX.utils.aoa_to_sheet(sheet.matrix, { cellDates: true }),
      sheet.name || "Sheet",
    );
  }
  state.fileBytes = window.XLSX.write(workbook, { bookType: "xlsx", type: "array" });
}

let saveChain = Promise.resolve();

function enqueueSave(job) {
  saveChain = saveChain.then(job, job);
  return saveChain;
}

function scheduleSave() {
  if (state.hydrating || !state.account) return;
  clearTimeout(scheduleSave.timer);
  scheduleSave.timer = setTimeout(() => {
    enqueueSave(() => saveWork({ deck: true })).catch((error) => toast(error.message || "Could not save."));
  }, 50);
}

function flushSave() {
  clearTimeout(scheduleSave.timer);
  if (state.hydrating || !state.account) return;
  if (!state.fileBytes && !state.workId) return;
  const includeFile = !state.workId;
  const body = saveBody({
    workId: state.workId,
    fileName: state.fileName,
    fileBase64: includeFile && state.fileBytes ? bytesToBase64(state.fileBytes) : "",
    sheetName: state.sheetName,
    slides: state.slides,
    active: state.active,
    includeFile,
  });
  const json = JSON.stringify(body);
  if (json.length > 60000) return;
  const path = apiUrl(state.workId ? `/api/works/${state.workId}` : "/api/works");
  fetch(path, {
    method: state.workId ? "PUT" : "POST",
    credentials: "same-origin",
    keepalive: true,
    headers: { "content-type": "application/json" },
    body: json,
  }).catch(() => {});
}

async function saveWork({ file = false, deck = false } = {}) {
  if (!state.account || state.hydrating) return;
  if (!state.fileBytes && !state.workId) return;
  const includeFile = file || !state.workId;
  if (includeFile && !state.fileBytes) return;
  const body = saveBody({
    workId: includeFile ? "" : state.workId,
    fileName: state.fileName,
    fileBase64: includeFile ? bytesToBase64(state.fileBytes) : "",
    sheetName: state.sheetName,
    slides: deck || !state.workId ? state.slides : undefined,
    active: state.active,
    includeFile,
  });
  if (!deck && state.workId && !includeFile) return;
  state.saveLabel = "saving";
  renderStatus();
  const path = apiUrl(state.workId ? `/api/works/${state.workId}` : "/api/works");
  const response = await fetch(path, {
    method: state.workId ? "PUT" : "POST",
    credentials: "same-origin",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(body),
  });
  const payload = await response.json().catch(() => ({}));
  if (response.status === 401) {
    state.account = null;
    state.saveLabel = "";
    paintAccount();
    throw new Error("Log in again to keep this file.");
  }
  if (!response.ok) throw new Error(payload.error || "Could not save.");
  state.workId = payload.id || state.workId;
  state.saveLabel = "saved";
  renderStatus();
}

function paintAccount() {
  const button = $("account");
  if (button) button.textContent = state.account ? state.account.email : "Log in";
}

function closeAccount() {
  $("account-dialog").hidden = true;
  document.body.classList.remove("account-open");
}

async function openAccount() {
  $("account-dialog").hidden = false;
  document.body.classList.add("account-open");
  const loggedIn = Boolean(state.account);
  $("account-form").hidden = loggedIn;
  $("account-home").hidden = !loggedIn;
  if (!loggedIn) {
    $("account-email").focus();
    return;
  }
  $("account-who").textContent = state.account.email;
  await refreshWorkList();
}

async function refreshWorkList() {
  const list = $("account-files");
  list.replaceChildren();
  const response = await fetch(apiUrl("/api/works"), { credentials: "same-origin" });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    toast(payload.error || "Could not load saved files.");
    return;
  }
  if (!payload.works?.length) {
    const item = document.createElement("li");
    item.className = "account-empty";
    item.textContent = "No saved files yet. Import a workbook while you are logged in.";
    list.appendChild(item);
    return;
  }
  for (const work of payload.works) {
    const item = document.createElement("li");
    item.className = "saved-file-row";
    const open = document.createElement("button");
    open.type = "button";
    open.className = "saved-file-open";
    const label = document.createElement("span");
    label.className = "saved-file-label";
    label.textContent = workMenuLabel(work);
    open.appendChild(label);
    const when = formatMoscowUploadTime(work.createdAt);
    if (when) {
      const time = document.createElement("span");
      time.className = "saved-file-when";
      time.textContent = when;
      open.appendChild(time);
    }
    open.addEventListener("click", () => openWork(work.id));
    item.appendChild(open);
    item.appendChild(deleteSavedButton(work.id, workMenuLabel(work), () => refreshWorkList()));
    list.appendChild(item);
  }
}

/** Put previously uploaded workbooks under Refresh so a newer upload does not hide them. */
async function fillRefreshSavedMenu() {
  const section = $("refresh-saved");
  const list = $("refresh-saved-list");
  if (!section || !list) return;
  if (!state.account) {
    section.hidden = true;
    list.replaceChildren();
    return;
  }
  const response = await fetch(apiUrl("/api/works"), { credentials: "same-origin" });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    section.hidden = true;
    list.replaceChildren();
    return;
  }
  const entries = refreshSavedEntries(payload.works, state.workId);
  list.replaceChildren();
  if (!entries.length) {
    section.hidden = true;
    return;
  }
  for (const entry of entries) {
    const row = document.createElement("div");
    row.className = "saved-file-row";
    const open = document.createElement("button");
    open.type = "button";
    open.className = "saved-file-open";
    if (entry.current) open.classList.add("menu-current");
    open.title = entry.current ? "Currently open" : `Open ${entry.label}`;
    const label = document.createElement("span");
    label.className = "saved-file-label";
    label.textContent = entry.label;
    open.appendChild(label);
    if (entry.when) {
      const time = document.createElement("span");
      time.className = "saved-file-when";
      time.textContent = entry.when;
      open.appendChild(time);
    }
    open.addEventListener("click", () => openWork(entry.id));
    row.appendChild(open);
    row.appendChild(deleteSavedButton(entry.id, entry.label, () => fillRefreshSavedMenu()));
    list.appendChild(row);
  }
  section.hidden = false;
}

function deleteSavedButton(id, label, afterDelete) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "saved-file-delete";
  button.setAttribute("aria-label", `Delete ${label}`);
  button.title = `Delete ${label}`;
  button.textContent = "×";
  button.addEventListener("click", async (event) => {
    event.stopPropagation();
    await deleteSavedWork(id, afterDelete);
  });
  return button;
}

async function deleteSavedWork(id, afterDelete) {
  const response = await fetch(apiUrl(`/api/works/${id}`), {
    method: "DELETE",
    credentials: "same-origin",
  });
  if (response.status === 401) {
    state.account = null;
    paintAccount();
    toast("Log in again to manage saved files.");
    return;
  }
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    toast(payload.error || "Could not delete that file.");
    return;
  }
  const wasCurrent = state.workId === id;
  if (wasCurrent) {
    state.workId = null;
    state.saveLabel = "";
  }
  toast("Deleted.");
  if (typeof afterDelete === "function") await afterDelete();
  if (!wasCurrent) return;
  const restored = await restoreLatest();
  if (!restored) clearImportedWork();
}

function clearImportedWork() {
  state.fileName = "";
  state.sheetName = "";
  state.matrix = null;
  state.org = null;
  state.workbookSheets = null;
  state.fileBytes = null;
  state.workId = null;
  state.saveLabel = "";
  state.imported = false;
  state.slides = [blankSlide()];
  state.active = 0;
  state.selection.clear();
  document.body.dataset.stage = "empty";
  render();
}

async function openWork(id, { quiet = false } = {}) {
  const response = await fetch(apiUrl(`/api/works/${id}`), { credentials: "same-origin" });
  const work = await response.json().catch(() => ({}));
  if (!response.ok) {
    toast(work.error || "Could not open that file.");
    return;
  }
  state.hydrating = true;
  state.fileBytes = base64ToBytes(work.dataBase64);
  state.workId = work.id;
  state.saveLabel = "saved";
  try {
    await readWorkbook(state.fileBytes, work.name);
    const sheet = state.workbookSheets?.find((item) => item.name === work.sheetName && !item.problem);
    if (sheet) importSheet(sheet.name);
    if (Array.isArray(work.slides) && work.slides.length) {
      state.slides = work.slides;
      state.active = Math.min(Math.max(work.active || 0, 0), work.slides.length - 1);
      state.selection.clear();
    }
  } finally {
    state.hydrating = false;
  }
  closeAccount();
  render();
  if (!quiet) toast(`Opened ${work.name}`);
  return work.name;
}

async function restoreLatest() {
  const response = await fetch(apiUrl("/api/works"), { credentials: "same-origin" });
  if (!response.ok) return "";
  const payload = await response.json().catch(() => ({}));
  const latest = workToReopen(payload.works);
  if (!latest) return "";
  return openWork(latest.id, { quiet: true });
}

async function submitAccount(path) {
  const response = await fetch(path, {
    method: "POST",
    credentials: "same-origin",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      email: $("account-email").value,
      password: $("account-password").value,
    }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    toast(payload.error || "Could not log in.");
    return;
  }
  state.account = { email: payload.email };
  $("account-password").value = "";
  paintAccount();
  closeAccount();
  try {
    if (state.fileBytes) {
      await enqueueSave(() => saveWork({ file: true, deck: true }));
      toast(`Saved to ${payload.email}`);
    } else {
      const restored = await restoreLatest();
      toast(restored ? `Restored ${restored}` : `Logged in as ${payload.email}`);
    }
  } catch (error) {
    toast(error.message || "Could not save.");
  }
}

async function logoutAccount() {
  await fetch(apiUrl("/api/logout"), { method: "POST", credentials: "same-origin" });
  state.account = null;
  state.workId = null;
  state.saveLabel = "";
  paintAccount();
  closeAccount();
  renderStatus();
}

async function restoreAccount() {
  try {
    const response = await fetch(apiUrl("/api/me"), { credentials: "same-origin" });
    if (!response.ok) return;
    const me = await response.json();
    state.account = { email: me.email };
    paintAccount();
    const restored = await restoreLatest();
    if (restored) toast(`Restored ${restored}`);
  } catch {
    // Opening the page from disk has no account server.
  }
}

bindSheetEditor();
bind();
render();
restoreAccount();
