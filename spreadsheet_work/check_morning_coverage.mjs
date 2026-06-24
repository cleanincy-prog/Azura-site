import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

const inputPath = (await exists("/Users/dmitro/Downloads/azura_guide_v2/outputs/schedule_break_fix/grafik_novyi_break_fixed.xlsx"))
  ? "/Users/dmitro/Downloads/azura_guide_v2/outputs/schedule_break_fix/grafik_novyi_break_fixed.xlsx"
  : "/Users/dmitro/Downloads/ГРАФИК НОВЫЙ.xlsx";
const extractDir = "/Users/dmitro/Downloads/azura_guide_v2/spreadsheet_work/morning_coverage_xml";
const outputDir = "/Users/dmitro/Downloads/azura_guide_v2/outputs/morning_coverage";

const requiredPositions = [
  "Waiter",
  "Somellier",
  "Runner",
  "Pasta",
  "Hot",
  "Hostess",
  "Dessert",
  "Cook assistant",
  "Cook asisstant",
  "Cold",
  "Bar",
];

function exists(filePath) {
  return fs
    .access(filePath)
    .then(() => true)
    .catch(() => false);
}

function decodeXml(text) {
  return String(text ?? "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function parseAttrs(attrsText) {
  const attrs = {};
  for (const match of attrsText.matchAll(/\s?([A-Za-z_:][\w:.-]*)="([^"]*)"/g)) {
    attrs[match[1]] = decodeXml(match[2]);
  }
  return attrs;
}

function parseSharedStrings(xml) {
  const strings = [];
  for (const match of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
    strings.push(
      [...match[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)]
        .map((part) => decodeXml(part[1]))
        .join(""),
    );
  }
  return strings;
}

function parseWorkbookSheets(workbookXml, relsXml) {
  const rels = new Map();
  for (const match of relsXml.matchAll(/<Relationship\b([^>]*)\/>/g)) {
    const attrs = parseAttrs(match[1]);
    rels.set(attrs.Id, attrs.Target);
  }

  const sheets = [];
  for (const match of workbookXml.matchAll(/<sheet\b([^>]*)\/>/g)) {
    const attrs = parseAttrs(match[1]);
    const target = rels.get(attrs["r:id"]);
    if (!target) {
      continue;
    }
    sheets.push({
      name: attrs.name,
      path: path.join(extractDir, target.startsWith("/") ? target.slice(1) : `xl/${target}`),
    });
  }
  return sheets;
}

function columnNameToIndex(column) {
  return [...column].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function columnIndexToName(indexZeroBased) {
  let n = indexZeroBased + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function parseAddress(address) {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    return null;
  }
  return { col: columnNameToIndex(match[1]), row: Number(match[2]) };
}

function cellAddress(row, col) {
  return `${columnIndexToName(col)}${row}`;
}

function cellValue(attrs, body, sharedStrings) {
  if (attrs.t === "inlineStr") {
    return [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)]
      .map((part) => decodeXml(part[1]))
      .join("");
  }
  const rawValue = body.match(/<v>([\s\S]*?)<\/v>/)?.[1];
  if (rawValue === undefined) {
    return "";
  }
  if (attrs.t === "s") {
    return sharedStrings[Number(rawValue)] ?? "";
  }
  return decodeXml(rawValue);
}

function parseCells(sheetXml, sharedStrings) {
  const cells = new Map();
  for (const match of sheetXml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
    const attrs = parseAttrs(match[1]);
    if (!attrs.r) {
      continue;
    }
    cells.set(attrs.r, cellValue(attrs, match[2], sharedStrings));
  }
  return cells;
}

function sheetMonthNumber(sheetName) {
  const months = {
    JANUARY: 1,
    FEBRUARY: 2,
    MARCH: 3,
    APRIL: 4,
    MAY: 5,
  };
  return months[sheetName.toUpperCase()] ?? null;
}

function excelSerialToIso(serial) {
  const epoch = Date.UTC(1899, 11, 30);
  return new Date(epoch + Number(serial) * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
}

function normalizeDate(value, sheetName) {
  const text = String(value ?? "").trim();
  if (!text) {
    return null;
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) {
    return text.slice(0, 10);
  }
  if (/^\d+(?:\.\d+)?$/.test(text)) {
    const num = Number(text);
    if (num >= 40000 && num <= 70000) {
      return excelSerialToIso(num);
    }
  }
  const dotted = text.match(/^(\d{1,2})\.(\d{1,2})/);
  if (dotted) {
    return `2026-${dotted[2].padStart(2, "0")}-${dotted[1].padStart(2, "0")}`;
  }
  const month = sheetMonthNumber(sheetName);
  const day = text.match(/^(\d{1,2})\b/);
  if (month && day) {
    return `2026-${String(month).padStart(2, "0")}-${day[1].padStart(2, "0")}`;
  }
  return text;
}

function displayDate(isoDate) {
  const match = String(isoDate).match(/^2026-(\d{2})-(\d{2})$/);
  return match ? `${match[2]}.${match[1]}` : isoDate;
}

function dateMatchesSheet(date, sheetName) {
  const month = sheetMonthNumber(sheetName);
  if (!month) {
    return true;
  }
  const match = String(date).match(/^2026-(\d{2})-\d{2}$/);
  return !match || Number(match[1]) === month;
}

function toHours(time) {
  const [hours, minutes] = time.split(":").map(Number);
  return hours + minutes / 60;
}

function cleanPosition(raw) {
  return raw.replace(/\s+/g, " ").replace(/[-–—:]+$/g, "").trim();
}

function parseShift(value) {
  const text = String(value ?? "").trim();
  const match = text.match(/^\s*(.*?)\s+(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})(?:\D.*)?$/);
  if (!match) {
    return null;
  }
  const position = cleanPosition(match[1]);
  if (!position || /^off|vacation$/i.test(position)) {
    return null;
  }
  return {
    position,
    start: match[2].padStart(5, "0"),
    end: match[3].padStart(5, "0"),
    text,
  };
}

function headerRows(cells, sheetName) {
  const rows = new Map();
  for (const [address, value] of cells.entries()) {
    const parsed = parseAddress(address);
    if (!parsed) {
      continue;
    }
    const row = rows.get(parsed.row) ?? new Map();
    row.set(parsed.col, String(value ?? ""));
    rows.set(parsed.row, row);
  }

  const headers = [];
  for (const [rowNumber, row] of rows.entries()) {
    const a = row.get(0)?.trim().toLowerCase();
    const c = row.get(2)?.trim().toLowerCase();
    if (a !== "name" || !c?.startsWith("total hours")) {
      continue;
    }
    const dates = new Map();
    for (const [col, value] of row.entries()) {
      if (col < 3) {
        continue;
      }
      const normalized = normalizeDate(value, sheetName);
      if (normalized) {
        dates.set(col, normalized);
      }
    }
    headers.push({ row: rowNumber, dates });
  }
  return headers.sort((a, b) => a.row - b.row);
}

function findHeader(headers, row) {
  let selected = null;
  for (const header of headers) {
    if (header.row < row) {
      selected = header;
    } else {
      break;
    }
  }
  return selected;
}

function formatPeople(records) {
  return records
    .map((record) => `${record.employee || "(без имени)"} ${record.start}-${record.end} ${record.cell}`)
    .join("; ");
}

await fs.rm(extractDir, { recursive: true, force: true });
await fs.mkdir(extractDir, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });
await execFileAsync("/usr/bin/unzip", ["-q", inputPath, "-d", extractDir]);

const sharedStrings = parseSharedStrings(await fs.readFile(path.join(extractDir, "xl/sharedStrings.xml"), "utf8"));
const sheets = parseWorkbookSheets(
  await fs.readFile(path.join(extractDir, "xl/workbook.xml"), "utf8"),
  await fs.readFile(path.join(extractDir, "xl/_rels/workbook.xml.rels"), "utf8"),
);

const morningByDayPosition = new Map();
const allDatesBySheet = new Map();

for (const sheet of sheets) {
  const cells = parseCells(await fs.readFile(sheet.path, "utf8"), sharedStrings);
  const headers = headerRows(cells, sheet.name);
  const sheetDates = new Set();
  for (const header of headers) {
    for (const date of header.dates.values()) {
      if (dateMatchesSheet(date, sheet.name)) {
        sheetDates.add(date);
      }
    }
  }
  allDatesBySheet.set(sheet.name, sheetDates);

  for (const [address, value] of cells.entries()) {
    const parsedAddress = parseAddress(address);
    if (!parsedAddress || parsedAddress.col < 3) {
      continue;
    }
    const shift = parseShift(value);
    if (!shift || !requiredPositions.includes(shift.position) || toHours(shift.start) >= 17) {
      continue;
    }
    const header = findHeader(headers, parsedAddress.row);
    const date = header?.dates.get(parsedAddress.col);
    if (!date || !dateMatchesSheet(date, sheet.name)) {
      continue;
    }
    const employee = String(cells.get(cellAddress(parsedAddress.row, 0)) ?? "").replace(/\s+/g, " ").trim();
    const key = `${sheet.name}\u0000${date}\u0000${shift.position}`;
    const list = morningByDayPosition.get(key) ?? [];
    list.push({
      sheet: sheet.name,
      date,
      position: shift.position,
      employee,
      start: shift.start,
      end: shift.end,
      cell: address,
      text: shift.text,
    });
    morningByDayPosition.set(key, list);
  }
}

const violations = [];
let checked = 0;
let ok = 0;

for (const [sheetName, dates] of allDatesBySheet.entries()) {
  for (const date of [...dates].sort()) {
    for (const position of requiredPositions) {
      checked += 1;
      const records = morningByDayPosition.get(`${sheetName}\u0000${date}\u0000${position}`) ?? [];
      if (records.length === 1) {
        ok += 1;
        continue;
      }
      violations.push({
        sheet: sheetName,
        date,
        position,
        count: records.length,
        people: formatPeople(records),
      });
    }
  }
}

violations.sort(
  (a, b) =>
    a.date.localeCompare(b.date) ||
    a.position.localeCompare(b.position) ||
    a.sheet.localeCompare(b.sheet),
);

const byType = {
  missing: violations.filter((item) => item.count === 0),
  extra: violations.filter((item) => item.count > 1),
};

const groupedMissing = new Map();
for (const item of byType.missing) {
  const key = `${item.position}\u0000${item.sheet}`;
  const group = groupedMissing.get(key) ?? { position: item.position, sheet: item.sheet, dates: [] };
  group.dates.push(displayDate(item.date));
  groupedMissing.set(key, group);
}

const result = {
  inputPath,
  checked,
  ok,
  violationCount: violations.length,
  missingCount: byType.missing.length,
  extraCount: byType.extra.length,
  groupedMissing: [...groupedMissing.values()].sort((a, b) => a.position.localeCompare(b.position) || a.sheet.localeCompare(b.sheet)),
  extra: byType.extra.map((item) => ({
    ...item,
    displayDate: displayDate(item.date),
  })),
};

await fs.writeFile(path.join(outputDir, "morning_coverage_result.json"), JSON.stringify(result, null, 2));

console.log(JSON.stringify(result, null, 2));
