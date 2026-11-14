/**
 * Employee model, Excel import, and the checks MyBoss runs before it draws.
 * Column order matches the VBA template (A–M):
 * surname, name, job, department, solid-line surname, solid-line name,
 * dotted-line surname, dotted-line name, location, termination, intern
 * flag, important flag, cost center.
 */

export const HOME_CITY = "Russia, Moscow";
export const COLUMN_COUNT = 13;

export const COLUMN_LABELS = [
  "Employee surname",
  "Employee name",
  "Job title",
  "Department",
  "Reports-to surname",
  "Reports-to name",
  "Dotted-line surname",
  "Dotted-line name",
  "Location",
  "Termination date",
  "Intern flag",
  "Important flag",
  "Cost center",
  "HR ID",
  "Hide flag",
  "Newcomer flag",
  "Newcomer date",
];

export function fullNameKey(surname, name, duplicateId = "") {
  return `${surname || ""} ${name || ""}`.trim().toLowerCase().replace(/\s+/g, "") + (duplicateId || "");
}

export function textInSquareBrackets(value) {
  const match = String(value ?? "").match(/\[[^\]]*\]/);
  return match ? match[0] : "";
}

export function moderated(value) {
  const raw = String(value ?? "").trim();
  const tag = textInSquareBrackets(raw);
  return tag ? raw.replace(tag, "").trim() : raw;
}

export function personFullName(person) {
  return `${person.surname} ${person.name}`.trim();
}

export function displayName(person) {
  if (person.tbh) return moderated(person.name) || person.name;
  return `${moderated(person.surname)} ${moderated(person.name)}`.trim();
}

function cellString(value) {
  if (value == null) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toLocaleDateString("en-GB", {
      day: "numeric",
      month: "long",
      year: "numeric",
      timeZone: "UTC",
    });
  }
  if (typeof value === "number" && value > 20000 && value < 80000) {
    const epoch = Date.UTC(1899, 11, 30);
    const date = new Date(epoch + Math.round(value) * 86400000);
    return date.toLocaleDateString("en-GB", {
      day: "numeric",
      month: "long",
      year: "numeric",
      timeZone: "UTC",
    });
  }
  return String(value).replace(/\r/g, "").trim();
}

/** Day and month only, "16.11", for the newcomer badge. Reads Excel dates, serials, dd.mm.yyyy, dd/mm/yyyy and yyyy-mm-dd; anything else is kept as typed. */
export function compactDate(value) {
  const two = (n) => String(n).padStart(2, "0");
  const fromDate = (date) => `${two(date.getUTCDate())}.${two(date.getUTCMonth() + 1)}`;
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? "" : fromDate(value);
  if (typeof value === "number" && value > 20000 && value < 80000) {
    return fromDate(new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000));
  }
  const text = String(value ?? "").trim();
  let match = text.match(/^(\d{1,2})[./-](\d{1,2})(?:[./-]\d{2,4})?$/);
  if (match) return `${two(match[1])}.${two(match[2])}`;
  match = text.match(/^\d{4}-(\d{1,2})-(\d{1,2})/);
  if (match) return `${two(match[2])}.${two(match[1])}`;
  return text;
}

export function flagOn(value) {
  if (value === true || value === 1) return true;
  const text = cellString(value).replace(",", ".");
  return text === "1" || text === "1.0";
}

function blankRow(row) {
  for (let i = 0; i < COLUMN_COUNT; i += 1) {
    if (cellString(row?.[i]) !== "") return false;
  }
  return true;
}

function emptyStats() {
  return {
    cards: 0,
    subordinates: 0,
    home: 0,
    ordinary: 0,
    important: 0,
    intern: 0,
    hidden: 0,
    hiddenOrdinary: 0,
    hiddenImportant: 0,
    hiddenIntern: 0,
  };
}

function addStats(target, extra) {
  target.cards += extra.cards;
  target.subordinates += extra.subordinates;
  target.home += extra.home;
  target.ordinary += extra.ordinary;
  target.important += extra.important;
  target.intern += extra.intern;
  target.hidden += extra.hidden;
  target.hiddenOrdinary += extra.hiddenOrdinary;
  target.hiddenImportant += extra.hiddenImportant;
  target.hiddenIntern += extra.hiddenIntern;
}

/**
 * @param {unknown[][]} matrix rows from the sheet, including a title row when hasTitles is set
 * @param {{ hasTitles?: boolean }} [options]
 */
export function buildOrg(matrix, options = {}) {
  const hasTitles = options.hasTitles !== false;
  const issues = [];
  const people = new Map();
  const reports = new Map();
  const firstRow = hasTitles ? 1 : 0;
  const rows = Array.isArray(matrix) ? matrix : [];
  const widest = rows.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0);
  if (widest < COLUMN_COUNT) {
    return {
      people,
      reports,
      order: [],
      ceo: null,
      companySize: 0,
      issues: [
        {
          level: "error",
          message: `This sheet doesn't have enough columns for the template. It needs columns A–M (${COLUMN_COUNT}).`,
        },
      ],
      fatal: true,
      hasTitles,
    };
  }

  const ensureReports = (key) => {
    if (!reports.has(key)) reports.set(key, []);
    return reports.get(key);
  };

  const halt = (message, row) => ({
    people,
    reports,
    order: [],
    ceo: null,
    companySize: 0,
    issues: [{ level: "error", message, ...(row ? { row } : {}) }],
    fatal: true,
    hasTitles,
  });

  for (let index = firstRow; index < rows.length; index += 1) {
    const row = rows[index] || [];
    const excelRow = index + 1;
    if (blankRow(row)) continue;

    const surname = cellString(row[0]);
    const name = cellString(row[1]);
    if (!surname || !name) {
      return halt(`Row ${excelRow} is missing a surname or a name. Every person needs both.`, excelRow);
    }

    const solidSurname = cellString(row[4]);
    const solidName = cellString(row[5]);
    const dottedSurname = cellString(row[6]);
    const dottedName = cellString(row[7]);
    let duplicateId = "";
    const baseKey = fullNameKey(surname, name);
    if (people.has(baseKey)) {
      const original = people.get(baseKey);
      original.duplicates += 1;
      duplicateId = String(original.duplicates + 1);
    }

    const key = fullNameKey(surname, name, duplicateId);
    const person = {
      key,
      duplicateId,
      duplicates: 0,
      surname,
      name,
      job: cellString(row[2]),
      department: cellString(row[3]),
      departmentKey: cellString(row[3]).toLowerCase().replace(/\s+/g, ""),
      solidSurname,
      solidName,
      solidKey: fullNameKey(solidSurname, solidName),
      dottedSurname,
      dottedName,
      dottedKey: fullNameKey(dottedSurname, dottedName),
      location: cellString(row[8]),
      termination: cellString(row[9]),
      intern: flagOn(row[10]),
      important: flagOn(row[11]),
      costCenter: cellString(row[12]),
      hrId: cellString(row[13]),
      hidden: flagOn(row[14]),
      newcomer: flagOn(row[15]),
      newcomerDate: compactDate(row[16]),
      row: excelRow,
      tbh: personFullName({ surname, name }).startsWith("TBH_"),
      stats: emptyStats(),
    };
    people.set(key, person);
    ensureReports(key);
    ensureReports(person.solidKey);
    reports.get(person.solidKey).push(key);
    if (person.dottedKey) {
      ensureReports(person.dottedKey);
      reports.get(person.dottedKey).push(key);
    }
  }

  const ceos = [...people.values()].filter((person) => person.solidKey === "");
  let ceo = null;
  if (ceos.length === 0) {
    return halt("No CEO found. Leave the solid-line manager empty for exactly one person.");
  }
  if (ceos.length > 1) {
    return halt(
      `Several people have no solid-line manager: ${ceos
        .map((person) => `${personFullName(person)} (row ${person.row})`)
        .join(", ")}. Only the CEO should.`,
    );
  }
  ceo = ceos[0].key;

  for (const person of people.values()) {
    if (person.key === ceo) continue;
    if (!person.solidKey || !people.has(person.solidKey)) {
      const named = `${person.solidSurname} ${person.solidName}`.trim();
      return halt(
        named
          ? `Row ${person.row}: ${personFullName(person)} reports to “${named}”, who is not in this sheet.`
          : `Row ${person.row}: ${personFullName(person)} has no solid-line manager.`,
        person.row,
      );
    }
    if (person.dottedKey) {
      if (person.dottedKey === person.solidKey) {
        return halt(
          `Row ${person.row}: ${personFullName(person)} has the same solid-line and dotted-line manager.`,
          person.row,
        );
      }
      if (!people.has(person.dottedKey)) {
        return halt(
          `Row ${person.row}: dotted-line manager “${person.dottedSurname} ${person.dottedName}” is not in this sheet.`,
          person.row,
        );
      }
    }
    for (const managerKey of [person.solidKey, person.dottedKey]) {
      if (!managerKey || !people.has(managerKey)) continue;
      const manager = people.get(managerKey);
      if (manager.duplicates > 0) {
        return halt(
          `Row ${person.row}: “${personFullName(manager)}” is a manager, but that name appears ${manager.duplicates + 1} times. Make the names unique.`,
          person.row,
        );
      }
    }
  }

  for (const person of people.values()) {
    if (!chainReachesCeo(people, person.key, ceo, "solid") || !chainReachesCeo(people, person.key, ceo, "dotted")) {
      return halt(
        `${personFullName(person)} (row ${person.row}) never reaches the CEO. Look for a loop or a break in the reporting line.`,
        person.row,
      );
    }
  }

  // Hidden people are still checked above. They are only left off the slides and the counts, so the
  // visible chart must not depend on them: the CEO and anyone with a visible report stay shown.
  const hiddenCeo = people.get(ceo);
  if (hiddenCeo?.hidden) {
    return halt(`Row ${hiddenCeo.row}: the CEO ${personFullName(hiddenCeo)} cannot be hidden.`, hiddenCeo.row);
  }
  for (const person of people.values()) {
    if (!person.hidden) continue;
    for (const childKey of reports.get(person.key) || []) {
      const child = people.get(childKey);
      if (child && !child.hidden) {
        return halt(
          `Row ${person.row}: ${personFullName(person)} is hidden, but ${personFullName(child)} (row ${child.row}) reports to them and is not. Hide them too or change their manager.`,
          person.row,
        );
      }
    }
  }

  const fatal = issues.some((issue) => issue.level === "error");
  if (!fatal) computeStats(people, reports);

  const order = [...people.keys()].filter((key) => !people.get(key).hidden).sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));
  let companySize = 0;
  for (const person of people.values()) if (!person.tbh && !person.hidden) companySize += 1;

  return {
    people,
    reports,
    order,
    ceo,
    companySize,
    issues,
    fatal,
    hasTitles,
  };
}

/**
 * The person, then each manager up the line. The first hop uses `regime`
 * (the line the card is hanging on). After that a two-manager person is
 * followed through the dotted key, matching extractReportingSequence.
 */
export function reportingChain(people, startKey, regime) {
  const chain = [];
  const start = people.get(startKey);
  if (!start) return chain;
  chain.push(startKey);
  let next = start.dottedKey ? (regime === "dotted" ? start.dottedKey : start.solidKey) : start.solidKey;
  const seen = new Set(chain);
  let guard = 0;
  while (next && !seen.has(next) && guard < 10000) {
    const node = people.get(next);
    if (!node) break;
    chain.push(next);
    seen.add(next);
    next = node.dottedKey ? node.dottedKey : node.solidKey;
    guard += 1;
  }
  return chain;
}

function chainReachesCeo(people, startKey, ceoKey, regime) {
  return reportingChain(people, startKey, regime).includes(ceoKey);
}

/** True when `managerKey` is above `personKey` on the given line. The person is not their own manager. */
export function reportsTo(org, personKey, managerKey, regime) {
  if (!personKey || !managerKey || personKey === managerKey) return false;
  return reportingChain(org.people, personKey, regime).includes(managerKey);
}

function computeStats(people, reports) {
  const done = new Set();
  const visiting = new Set();

  function walk(key) {
    const person = people.get(key);
    if (!person) return emptyStats();
    if (done.has(key)) return person.stats;
    if (visiting.has(key)) return emptyStats();
    visiting.add(key);
    const stats = emptyStats();
    for (const childKey of reports.get(key) || []) {
      const child = people.get(childKey);
      if (!child || managerType(child, key) !== "Solid") continue;
      addStats(stats, walk(childKey));
    }
    if (person.hidden) {
      // Tallied on their own; nothing here reaches the visible counts.
      stats.hidden += 1;
      if (!(person.intern || person.important)) stats.hiddenOrdinary += 1;
      if (person.intern) stats.hiddenIntern += 1;
      if (person.important) stats.hiddenImportant += 1;
    } else {
      stats.cards += 1;
      stats.subordinates += 1;
      if (person.location === HOME_CITY) stats.home += 1;
      if (!(person.intern || person.important)) stats.ordinary += 1;
      if (person.intern) stats.intern += 1;
      if (person.important) stats.important += 1;
    }
    person.stats = stats;
    done.add(key);
    visiting.delete(key);
    return stats;
  }

  for (const key of people.keys()) walk(key);
}

/**
 * What a card needs to be drawn, copied at the moment the slide is made. A slide is a picture of the
 * data at that time: importing another sheet or file later must not change cards already on slides.
 */
export function snapshotPerson(person) {
  return {
    surname: person.surname,
    name: person.name,
    job: person.job,
    department: person.department,
    location: person.location,
    termination: person.termination,
    intern: person.intern,
    important: person.important,
    tbh: person.tbh,
    newcomer: person.newcomer,
    newcomerDate: person.newcomerDate,
    costCenter: person.costCenter,
    key: person.key,
    solidKey: person.solidKey,
    dottedKey: person.dottedKey,
    solidName: personFullName({ surname: person.solidSurname, name: person.solidName }),
    dottedName: personFullName({ surname: person.dottedSurname, name: person.dottedName }),
    stats: { subordinates: person.stats?.subordinates ?? 0, intern: person.stats?.intern ?? 0 },
  };
}

/** The person a card shows: its own snapshot, else (cards from older sessions) the loaded sheet. */
export function shapePerson(org, shape) {
  if (shape?.person) return shape.person;
  return shape?.key ? org?.people?.get(shape.key) || null : null;
}

/** Give every person card on `shapes` its own copy of the person's data. */
export function freezeCards(org, shapes) {
  for (const shape of shapes) {
    if (shape.type !== "card" || !shape.key || shape.person) continue;
    const person = org?.people?.get(shape.key);
    if (person) shape.person = snapshotPerson(person);
  }
}

export function managerType(person, managerKey) {
  if (!person || !managerKey) return "";
  if (person.solidKey === managerKey) return "Solid";
  if (person.dottedKey === managerKey) return "Dotted";
  return "";
}

export function hasTwoManagers(person) {
  return Boolean(person?.dottedKey);
}

/** Reports that are drawn: hidden people are left out. */
export function visibleReports(org, key) {
  return (org.reports.get(key) || []).filter((id) => org.people.has(id) && !org.people.get(id).hidden);
}

export function isManager(org, key) {
  return visibleReports(org, key).length > 0;
}

export function comparePeople(org, a, b) {
  const left = org.people.get(a)?.stats.subordinates ?? 0;
  const right = org.people.get(b)?.stats.subordinates ?? 0;
  if (left !== right) return right - left;
  return a < b ? -1 : a > b ? 1 : 0;
}

export function directReports(org, key) {
  const all = visibleReports(org, key);
  const sorted = [...all].sort((a, b) => comparePeople(org, a, b));
  const solid = [];
  const dotted = [];
  const workers = [];
  const managers = [];
  for (const id of sorted) {
    const person = org.people.get(id);
    const kind = managerType(person, key);
    if (kind === "Solid") solid.push(id);
    else if (kind === "Dotted") dotted.push(id);
    if (isManager(org, id)) managers.push(id);
    else workers.push(id);
  }
  return { all: sorted, solid, dotted, workers, managers };
}

/** Departments in first-seen order. Empty department names stay together. */
export function departmentsOf(org, keys) {
  const groups = [];
  const index = new Map();
  for (const key of keys) {
    const person = org.people.get(key);
    const deptKey = person.departmentKey || "";
    if (!index.has(deptKey)) {
      index.set(deptKey, groups.length);
      groups.push({
        key: deptKey,
        name: person.department || "",
        people: [],
      });
    }
    groups[index.get(deptKey)].people.push(key);
  }
  return groups;
}

/** Wording from the Department statistics shape on the template slide. Hidden people are summarised in the slide footer instead. */
export function statisticsLines(person, directCount) {
  const stats = person?.stats || {};
  const first = moderated(person?.name) || "This person";
  const lines = [
    `${first} has ${directCount} direct reports`,
    `Team Size – ${stats.subordinates || 0} people`,
    `${stats.ordinary || 0} ordinary employees`,
    `${stats.important || 0} non-ordinary employees`,
    `${stats.intern || 0} interns`,
  ];
  return lines;
}

export function headcountLabel(person) {
  const count = person?.stats?.subordinates ?? 0;
  if (count <= 1) return "";
  const intern = person.stats.intern || 0;
  return intern > 0 ? `${count} (${intern})` : String(count);
}

export function cardStyle(person) {
  if (!person) return "ordinary";
  // A vacancy takes the colors of the flags it carries. Plain white is only for a
  // vacancy with no flag; the dashed outline (person.tbh) is what marks it as open.
  if (person.important && person.intern) return "split";
  if (person.important) return "important";
  if (person.intern) return "intern";
  return person.tbh ? "tbh" : "ordinary";
}

/** Everyone reached from `rootKey` by solid-line links, including the root. */
export function solidTree(org, rootKey) {
  const seen = new Set();
  const walk = (key) => {
    if (!key || seen.has(key) || !org.people.has(key)) return;
    seen.add(key);
    for (const child of org.reports.get(key) || []) {
      const person = org.people.get(child);
      if (person && !person.hidden && person.solidKey === key) walk(child);
    }
  };
  walk(rootKey);
  return seen;
}

export const VACANCIES = "Vacancies";

/** Picker text: a vacancy is "Asyat (vacancy)", not the raw TBH_ code. */
export function pickerLabel(person) {
  if (!person) return "";
  if (!person.tbh) return displayName(person);
  const label = moderated(person.name).replace(/^TBH_/i, "").trim() || moderated(person.surname).replace(/^TBH_/i, "").trim();
  return `${label || "Unnamed"} (vacancy)`;
}

export function categories(org) {
  const letters = new Set();
  let vacancies = false;
  for (const key of org.order) {
    if (org.people.get(key).tbh) vacancies = true;
    else letters.add(key.slice(0, 1).toUpperCase());
  }
  // Any alphabet: a Cyrillic surname gets its own letter, like a Latin one.
  const list = [...letters].filter((letter) => /\p{L}/u.test(letter)).sort((x, y) => x.localeCompare(y));
  return ["CEO", "Managers", "AZ", ...list, ...(vacancies ? [VACANCIES] : [])];
}

export function peopleInCategory(org, category) {
  if (!org?.order) return [];
  if (category === "CEO") return org.ceo ? [org.ceo] : [];
  if (category === "Managers") return org.order.filter((key) => isManager(org, key));
  if (category === "AZ") return org.order;
  if (category === VACANCIES) return org.order.filter((key) => org.people.get(key).tbh);
  const letter = String(category || "").toLowerCase();
  return org.order.filter((key) => !org.people.get(key).tbh && key.startsWith(letter));
}
