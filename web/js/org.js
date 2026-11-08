/**
 * Employee model, Excel import, and the checks MyBoss runs before it draws.
 * Column order matches the VBA template (A–M):
 * surname, name, job, department, solid-line surname, solid-line name,
 * dotted-line surname, dotted-line name, location, termination, outsourcer
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
  "Outsourcer flag",
  "Important flag",
  "Cost center",
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

function flagOn(value) {
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
  return { cards: 0, subordinates: 0, home: 0, ordinary: 0, important: 0, outsource: 0 };
}

function addStats(target, extra) {
  target.cards += extra.cards;
  target.subordinates += extra.subordinates;
  target.home += extra.home;
  target.ordinary += extra.ordinary;
  target.important += extra.important;
  target.outsource += extra.outsource;
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

  for (let index = firstRow; index < rows.length; index += 1) {
    const row = rows[index] || [];
    const excelRow = index + 1;
    if (blankRow(row)) continue;

    const surname = cellString(row[0]);
    const name = cellString(row[1]);
    if (!surname || !name) {
      issues.push({
        level: "error",
        row: excelRow,
        message: `Row ${excelRow} is missing a surname or a name. Every person needs both.`,
      });
      continue;
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
      outsource: flagOn(row[10]),
      important: flagOn(row[11]),
      costCenter: cellString(row[12]),
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
    issues.push({
      level: "error",
      message: "No CEO found. Leave the solid-line manager empty for exactly one person.",
    });
  } else if (ceos.length > 1) {
    issues.push({
      level: "error",
      message: `Several people have no solid-line manager: ${ceos
        .map((person) => `${personFullName(person)} (row ${person.row})`)
        .join(", ")}. Only the CEO should.`,
    });
  } else {
    ceo = ceos[0].key;
  }

  for (const person of people.values()) {
    if (person.key === ceo) continue;
    if (!person.solidKey || !people.has(person.solidKey)) {
      const named = `${person.solidSurname} ${person.solidName}`.trim();
      issues.push({
        level: "error",
        row: person.row,
        message: named
          ? `Row ${person.row}: ${personFullName(person)} reports to “${named}”, who is not in this sheet.`
          : `Row ${person.row}: ${personFullName(person)} has no solid-line manager.`,
      });
    }
    if (person.dottedKey) {
      if (person.dottedKey === person.solidKey) {
        issues.push({
          level: "error",
          row: person.row,
          message: `Row ${person.row}: ${personFullName(person)} has the same solid-line and dotted-line manager.`,
        });
      } else if (!people.has(person.dottedKey)) {
        issues.push({
          level: "error",
          row: person.row,
          message: `Row ${person.row}: dotted-line manager “${person.dottedSurname} ${person.dottedName}” is not in this sheet.`,
        });
      }
    }
    for (const managerKey of [person.solidKey, person.dottedKey]) {
      if (!managerKey || !people.has(managerKey)) continue;
      const manager = people.get(managerKey);
      if (manager.duplicates > 0) {
        issues.push({
          level: "error",
          row: person.row,
          message: `Row ${person.row}: “${personFullName(manager)}” is a manager, but that name appears ${manager.duplicates + 1} times. Make the names unique.`,
        });
      }
    }
  }

  if (ceo && !issues.some((issue) => issue.level === "error")) {
    for (const person of people.values()) {
      if (!reachesCeo(people, person.key, ceo)) {
        issues.push({
          level: "error",
          row: person.row,
          message: `${personFullName(person)} (row ${person.row}) never reaches the CEO. Look for a loop or a break in the reporting line.`,
        });
        break;
      }
    }
  }

  const fatal = issues.some((issue) => issue.level === "error");
  if (!fatal) computeStats(people, reports);

  const order = [...people.keys()].sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));
  let companySize = 0;
  for (const person of people.values()) if (!person.tbh) companySize += 1;

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

function reachesCeo(people, startKey, ceoKey) {
  const seen = new Set();
  let key = startKey;
  while (key && !seen.has(key)) {
    if (key === ceoKey) return true;
    seen.add(key);
    const person = people.get(key);
    if (!person) return false;
    key = person.solidKey;
  }
  return false;
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
    stats.cards += 1;
    if (!person.tbh) {
      stats.subordinates += 1;
      if (person.location === HOME_CITY) stats.home += 1;
      if (person.outsource) stats.outsource += 1;
      else if (person.important) stats.important += 1;
      else stats.ordinary += 1;
    }
    person.stats = stats;
    done.add(key);
    visiting.delete(key);
    return stats;
  }

  for (const key of people.keys()) walk(key);
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

export function isManager(org, key) {
  return (org.reports.get(key) || []).length > 0;
}

export function comparePeople(org, a, b) {
  const left = org.people.get(a)?.stats.subordinates ?? 0;
  const right = org.people.get(b)?.stats.subordinates ?? 0;
  if (left !== right) return right - left;
  return a < b ? -1 : a > b ? 1 : 0;
}

export function directReports(org, key) {
  const all = [...(org.reports.get(key) || [])].filter((id) => org.people.has(id));
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

/** Wording from the Department statistics shape on the template slide. */
export function statisticsLines(person, directCount) {
  const stats = person?.stats || {};
  const first = moderated(person?.name) || "This person";
  return [
    `${first} has ${directCount} direct reports`,
    `Team Size – ${stats.subordinates || 0} people`,
    `${stats.ordinary || 0} ordinary employees`,
    `${stats.important || 0} non-ordinary employees`,
    `${stats.outsource || 0} outsourcers`,
  ];
}

export function headcountLabel(person) {
  const count = person?.stats?.subordinates ?? 0;
  if (count <= 1) return "";
  const external = person.stats.outsource || 0;
  return external > 0 ? `${count} (${external})` : String(count);
}

export function cardStyle(person) {
  if (!person) return "ordinary";
  if (person.tbh) return "tbh";
  if (person.important) return "important";
  if (person.outsource) return "outsource";
  return "ordinary";
}

export function categories(org) {
  const letters = new Set();
  for (const key of org.order) letters.add(key.slice(0, 1).toUpperCase());
  const list = [...letters].filter((letter) => /[A-Z]/.test(letter)).sort();
  return ["CEO", "Managers", "AZ", ...list];
}

export function peopleInCategory(org, category) {
  if (!org?.order) return [];
  if (category === "CEO") return org.ceo ? [org.ceo] : [];
  if (category === "Managers") return org.order.filter((key) => isManager(org, key));
  if (category === "AZ") return org.order.slice(0, 1000);
  const letter = String(category || "").toLowerCase();
  return org.order.filter((key) => key.startsWith(letter));
}
