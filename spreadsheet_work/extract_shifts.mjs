import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const execFileAsync = promisify(execFile);

const preferredInput =
  "/Users/dmitro/Downloads/azura_guide_v2/outputs/schedule_break_fix/grafik_novyi_break_fixed.xlsx";
const fallbackInput = "/Users/dmitro/Downloads/ГРАФИК НОВЫЙ.xlsx";
const inputPath = await exists(preferredInput) ? preferredInput : fallbackInput;
const extractDir = "/Users/dmitro/Downloads/azura_guide_v2/spreadsheet_work/shift_extract_xml";
const outputDir = "/Users/dmitro/Downloads/azura_guide_v2/outputs/shift_extract";
const outputPath = path.join(outputDir, "shift_list_by_section.xlsx");

function exists(filePath) {
  return fs
    .access(filePath)
    .then(() => true)
    .catch(() => false);
}

function decodeXml(text) {
  return text
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function parseAttrs(attrsText) {
  const attrs = {};
  for (const match of attrsText.matchAll(/\s([A-Za-z_:][\w:.-]*)="([^"]*)"/g)) {
    attrs[match[1]] = decodeXml(match[2]);
  }
  return attrs;
}

function parseSharedStrings(xml) {
  const strings = [];
  for (const match of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
    const parts = [...match[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((part) =>
      decodeXml(part[1]),
    );
    strings.push(parts.join(""));
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
    const relId = attrs["r:id"];
    const target = rels.get(relId);
    if (!target) {
      continue;
    }
    const normalizedTarget = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
    sheets.push({
      name: attrs.name,
      sheetId: attrs.sheetId,
      path: path.join(extractDir, normalizedTarget),
    });
  }
  return sheets;
}

function columnNameToIndex(column) {
  return [...column].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function parseAddress(address) {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    return null;
  }
  return { col: columnNameToIndex(match[1]), row: Number(match[2]) };
}

function cellValue(attrs, body, sharedStrings) {
  const valueRaw = body.match(/<v>([\s\S]*?)<\/v>/)?.[1];
  if (attrs.t === "inlineStr") {
    const parts = [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((part) =>
      decodeXml(part[1]),
    );
    return parts.join("");
  }
  if (valueRaw === undefined) {
    return "";
  }
  if (attrs.t === "s") {
    return sharedStrings[Number(valueRaw)] ?? "";
  }
  return decodeXml(valueRaw);
}

function parseSheetCells(sheetXml, sharedStrings) {
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

function toHours(time) {
  const [hh, mm] = time.split(":").map(Number);
  return hh + mm / 60;
}

function durationHours(start, end) {
  let endHours = toHours(end);
  const startHours = toHours(start);
  if (endHours <= startHours) {
    endHours += 24;
  }
  return Number((endHours - startHours).toFixed(2));
}

function excelSerialToDate(serial) {
  const epoch = Date.UTC(1899, 11, 30);
  return new Date(epoch + serial * 24 * 60 * 60 * 1000);
}

function displayHeaderDate(value) {
  const text = String(value ?? "").trim();
  if (/^\d+(?:\.\d+)?$/.test(text)) {
    const number = Number(text);
    if (number >= 40000 && number <= 70000) {
      return excelSerialToDate(number).toISOString().slice(0, 10);
    }
  }
  return text;
}

function cleanSection(raw) {
  return raw
    .replace(/\s+/g, " ")
    .replace(/[-–—:]+$/g, "")
    .trim();
}

function parseShift(value) {
  if (typeof value !== "string") {
    return null;
  }
  const match = value.match(/^\s*(.*?)\s+(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\s*$/);
  if (!match) {
    return null;
  }
  const section = cleanSection(match[1]);
  if (!section || /^off|vacation$/i.test(section)) {
    return null;
  }
  return {
    section,
    start: match[2].padStart(5, "0"),
    end: match[3].padStart(5, "0"),
  };
}

function uniqueSorted(values) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function summarizeEmployees(values, max = 8) {
  const employees = uniqueSorted(values);
  if (employees.length <= max) {
    return employees.join(", ");
  }
  return `${employees.slice(0, max).join(", ")} +${employees.length - max}`;
}

function a1(row, col) {
  let n = col + 1;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return `${letters}${row}`;
}

function writeMatrix(sheet, startRow, startCol, matrix) {
  if (!matrix.length) {
    return;
  }
  const endRow = startRow + matrix.length - 1;
  const endCol = startCol + matrix[0].length - 1;
  sheet.getRange(`${a1(startRow, startCol)}:${a1(endRow, endCol)}`).values = matrix;
}

function styleTable(sheet, rangeAddress, headerAddress, numberColumns = []) {
  sheet.getRange(headerAddress).format = {
    fill: "#0F5D75",
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true,
  };
  sheet.getRange(rangeAddress).format.borders = {
    insideHorizontal: { style: "thin", color: "#D8E2E7" },
    insideVertical: { style: "thin", color: "#D8E2E7" },
    bottom: { style: "thin", color: "#9AB4C0" },
  };
  for (const column of numberColumns) {
    sheet.getRange(column).format.numberFormat = "#,##0.00";
  }
}

await fs.rm(extractDir, { recursive: true, force: true });
await fs.mkdir(extractDir, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });

await execFileAsync("/usr/bin/unzip", ["-q", inputPath, "-d", extractDir]);

const sharedStrings = parseSharedStrings(
  await fs.readFile(path.join(extractDir, "xl/sharedStrings.xml"), "utf8"),
);
const sheets = parseWorkbookSheets(
  await fs.readFile(path.join(extractDir, "xl/workbook.xml"), "utf8"),
  await fs.readFile(path.join(extractDir, "xl/_rels/workbook.xml.rels"), "utf8"),
);

const records = [];

for (const sheet of sheets) {
  const cells = parseSheetCells(await fs.readFile(sheet.path, "utf8"), sharedStrings);
  for (const [address, value] of cells.entries()) {
    const parsedAddress = parseAddress(address);
    if (!parsedAddress || parsedAddress.col < 3) {
      continue;
    }
    const shift = parseShift(String(value));
    if (!shift) {
      continue;
    }

    const employee = String(cells.get(`A${parsedAddress.row}`) ?? "").trim();
    const skills = String(cells.get(`B${parsedAddress.row}`) ?? "").trim();
    const dateHeader = displayHeaderDate(cells.get(a1(1, parsedAddress.col)));

    records.push({
      month: sheet.name,
      cell: address,
      employee,
      skills,
      date: dateHeader,
      shiftText: String(value).trim(),
      section: shift.section,
      start: shift.start,
      end: shift.end,
      duration: durationHours(shift.start, shift.end),
      paidHours: Math.max(0, durationHours(shift.start, shift.end) - 1),
    });
  }
}

records.sort((a, b) =>
  a.section.localeCompare(b.section) ||
  a.start.localeCompare(b.start) ||
  a.end.localeCompare(b.end) ||
  a.month.localeCompare(b.month) ||
  a.employee.localeCompare(b.employee),
);

const summaryMap = new Map();
for (const record of records) {
  const key = `${record.section}\u0000${record.start}\u0000${record.end}`;
  const current =
    summaryMap.get(key) ??
    {
      section: record.section,
      start: record.start,
      end: record.end,
      duration: record.duration,
      paidHours: record.paidHours,
      count: 0,
      months: [],
      employees: [],
      skills: [],
    };
  current.count += 1;
  current.months.push(record.month);
  current.employees.push(record.employee);
  current.skills.push(record.skills);
  summaryMap.set(key, current);
}

const summary = [...summaryMap.values()].map((item) => ({
  ...item,
  months: uniqueSorted(item.months).join(", "),
  employees: summarizeEmployees(item.employees),
  skills: uniqueSorted(item.skills).join(", "),
}));

summary.sort(
  (a, b) =>
    a.section.localeCompare(b.section) ||
    a.start.localeCompare(b.start) ||
    a.end.localeCompare(b.end),
);

const workbook = Workbook.create();
const listSheet = workbook.worksheets.add("Список смен");
const detailSheet = workbook.worksheets.add("Все записи");
const byTimeSheet = workbook.worksheets.add("По времени");

listSheet.showGridLines = false;
detailSheet.showGridLines = false;
byTimeSheet.showGridLines = false;

listSheet.getRange("A1:H1").merge();
listSheet.getRange("A1").values = [["Уникальные смены по цехам / позициям"]];
listSheet.getRange("A1").format = {
  fill: "#E7F3F7",
  font: { bold: true, size: 14, color: "#0B3A4A" },
};
listSheet.getRange("A2:H2").merge();
listSheet.getRange("A2").values = [[`Источник: ${path.basename(inputPath)}. Найдено записей: ${records.length}. Уникальных смен: ${summary.length}.`]];
listSheet.getRange("A2").format = { fill: "#F5FAFC", font: { color: "#315766" } };

const summaryRows = [
  [
    "Цех / позиция",
    "Начало",
    "Конец",
    "Длительность, ч",
    "Часы минус перерыв",
    "Сколько раз",
    "Месяцы",
    "Сотрудники",
  ],
  ...summary.map((item) => [
    item.section,
    item.start,
    item.end,
    item.duration,
    item.paidHours,
    item.count,
    item.months,
    item.employees,
  ]),
];
writeMatrix(listSheet, 4, 0, summaryRows);
styleTable(listSheet, `A4:H${summaryRows.length + 3}`, "A4:H4", ["D5:E200"]);
const listLastRow = summaryRows.length + 3;
listSheet.getRange(`A5:C${listLastRow}`).format = { horizontalAlignment: "left" };
listSheet.getRange(`F5:F${listLastRow}`).format.numberFormat = "#,##0";
listSheet.getRange(`G5:H${listLastRow}`).format.wrapText = true;
listSheet.freezePanes.freezeRows(4);
listSheet.getRange(`A1:A${listLastRow}`).format.columnWidth = 22;
listSheet.getRange(`B1:C${listLastRow}`).format.columnWidth = 11;
listSheet.getRange(`D1:F${listLastRow}`).format.columnWidth = 15;
listSheet.getRange(`G1:G${listLastRow}`).format.columnWidth = 28;
listSheet.getRange(`H1:H${listLastRow}`).format.columnWidth = 60;

const detailRows = [
  [
    "Месяц",
    "Дата / колонка",
    "Сотрудник",
    "Навыки / отдел",
    "Текст смены",
    "Цех / позиция",
    "Начало",
    "Конец",
    "Длительность, ч",
    "Часы минус перерыв",
    "Ячейка",
  ],
  ...records.map((record) => [
    record.month,
    record.date,
    record.employee,
    record.skills,
    record.shiftText,
    record.section,
    record.start,
    record.end,
    record.duration,
    record.paidHours,
    record.cell,
  ]),
];
writeMatrix(detailSheet, 1, 0, detailRows);
styleTable(detailSheet, `A1:K${detailRows.length}`, "A1:K1", ["I2:J5000"]);
const detailLastRow = detailRows.length;
detailSheet.getRange(`I2:J${detailLastRow}`).format.numberFormat = "#,##0.00";
detailSheet.freezePanes.freezeRows(1);
detailSheet.getRange(`A1:A${detailLastRow}`).format.columnWidth = 13;
detailSheet.getRange(`B1:B${detailLastRow}`).format.columnWidth = 14;
detailSheet.getRange(`C1:D${detailLastRow}`).format.columnWidth = 24;
detailSheet.getRange(`E1:F${detailLastRow}`).format.columnWidth = 24;
detailSheet.getRange(`G1:H${detailLastRow}`).format.columnWidth = 11;
detailSheet.getRange(`I1:J${detailLastRow}`).format.columnWidth = 16;
detailSheet.getRange(`K1:K${detailLastRow}`).format.columnWidth = 10;

const byTimeMap = new Map();
for (const item of summary) {
  const key = `${item.start}\u0000${item.end}`;
  const current =
    byTimeMap.get(key) ??
    {
      start: item.start,
      end: item.end,
      duration: item.duration,
      paidHours: item.paidHours,
      sections: [],
      count: 0,
    };
  current.sections.push(item.section);
  current.count += item.count;
  byTimeMap.set(key, current);
}
const byTime = [...byTimeMap.values()]
  .map((item) => ({
    ...item,
    sections: uniqueSorted(item.sections).join(", "),
  }))
  .sort((a, b) => a.start.localeCompare(b.start) || a.end.localeCompare(b.end));

const byTimeRows = [
  ["Начало", "Конец", "Длительность, ч", "Часы минус перерыв", "Цеха / позиции", "Сколько раз"],
  ...byTime.map((item) => [
    item.start,
    item.end,
    item.duration,
    item.paidHours,
    item.sections,
    item.count,
  ]),
];
writeMatrix(byTimeSheet, 1, 0, byTimeRows);
styleTable(byTimeSheet, `A1:F${byTimeRows.length}`, "A1:F1", ["C2:D200"]);
const byTimeLastRow = byTimeRows.length;
byTimeSheet.getRange(`F2:F${byTimeLastRow}`).format.numberFormat = "#,##0";
byTimeSheet.getRange(`E2:E${byTimeLastRow}`).format.wrapText = true;
byTimeSheet.freezePanes.freezeRows(1);
byTimeSheet.getRange(`A1:B${byTimeLastRow}`).format.columnWidth = 12;
byTimeSheet.getRange(`C1:D${byTimeLastRow}`).format.columnWidth = 17;
byTimeSheet.getRange(`E1:E${byTimeLastRow}`).format.columnWidth = 60;
byTimeSheet.getRange(`F1:F${byTimeLastRow}`).format.columnWidth = 14;

listSheet.getRange(`A5:H${listLastRow}`).format.font = { name: "Arial", size: 10 };
detailSheet.getRange(`A2:K${detailLastRow}`).format.font = { name: "Arial", size: 10 };
byTimeSheet.getRange(`A2:F${byTimeLastRow}`).format.font = { name: "Arial", size: 10 };

const preview = await workbook.render({
  sheetName: "Список смен",
  range: `A1:H${Math.min(summaryRows.length + 3, 35)}`,
  scale: 1,
  format: "png",
});
await fs.writeFile(
  path.join(outputDir, "shift_list_preview.png"),
  new Uint8Array(await preview.arrayBuffer()),
);

const check = await workbook.inspect({
  kind: "table",
  range: "Список смен!A1:H15",
  include: "values",
  tableMaxRows: 15,
  tableMaxCols: 8,
  maxChars: 12000,
});
console.log(check.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

await fs.writeFile(
  path.join(outputDir, "shift_extract_stats.json"),
  JSON.stringify(
    {
      inputPath,
      records: records.length,
      uniqueShiftsBySection: summary.length,
      uniqueTimes: byTime.length,
      sections: uniqueSorted(summary.map((item) => item.section)),
    },
    null,
    2,
  ),
);

console.log(`Records: ${records.length}`);
console.log(`Unique shifts by section: ${summary.length}`);
console.log(`Unique start/end times: ${byTime.length}`);
console.log(`Output: ${outputPath}`);
