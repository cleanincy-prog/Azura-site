import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

const inputPath = "/Users/dmitro/Downloads/ГРАФИК НОВЫЙ1.xlsx";
const extractDir = "/Users/dmitro/Downloads/azura_guide_v2/spreadsheet_work/vacation_weekends_xml";
const outputDir = "/Users/dmitro/Downloads/azura_guide_v2/outputs/vacation_weekends_off";
const outputPath = path.join(outputDir, "grafik_novyi1_weekend_vacation_off.xlsx");

function decodeXml(text) {
  return String(text ?? "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function encodeXmlText(text) {
  return String(text ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
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

function updateSharedStringCount(xml, nextCount, nextUniqueCount) {
  return xml
    .replace(/\bcount="[^"]*"/, `count="${nextCount}"`)
    .replace(/\buniqueCount="[^"]*"/, `uniqueCount="${nextUniqueCount}"`);
}

function ensureSharedString(xml, strings, value) {
  const existingIndex = strings.findIndex((item) => item.toUpperCase() === value.toUpperCase());
  if (existingIndex >= 0) {
    return { xml, strings, index: existingIndex };
  }
  const nextStrings = [...strings, value];
  const insert = `<si><t>${encodeXmlText(value)}</t></si>`;
  const nextXml = updateSharedStringCount(
    xml.replace("</sst>", `${insert}</sst>`),
    nextStrings.length,
    nextStrings.length,
  );
  return { xml: nextXml, strings: nextStrings, index: nextStrings.length - 1 };
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
    if (!target) continue;
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
  if (!match) return null;
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
  if (rawValue === undefined) return "";
  if (attrs.t === "s") return sharedStrings[Number(rawValue)] ?? "";
  return decodeXml(rawValue);
}

function parseCells(sheetXml, sharedStrings) {
  const cells = new Map();
  for (const match of sheetXml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
    const attrs = parseAttrs(match[1]);
    if (!attrs.r) continue;
    cells.set(attrs.r, cellValue(attrs, match[2], sharedStrings));
  }
  return cells;
}

function sheetMonthNumber(sheetName) {
  const months = { JANUARY: 1, FEBRUARY: 2, MARCH: 3, APRIL: 4, MAY: 5 };
  return months[String(sheetName).toUpperCase()] ?? null;
}

function excelSerialToIso(serial) {
  const epoch = Date.UTC(1899, 11, 30);
  return new Date(epoch + Number(serial) * 86400000).toISOString().slice(0, 10);
}

function normalizeDate(value, sheetName) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);
  if (/^\d+(?:\.\d+)?$/.test(text)) {
    const num = Number(text);
    if (num >= 40000 && num <= 70000) return excelSerialToIso(num);
  }
  const dotted = text.match(/^(\d{1,2})\.(\d{1,2})/);
  if (dotted) return `2026-${dotted[2].padStart(2, "0")}-${dotted[1].padStart(2, "0")}`;
  const month = sheetMonthNumber(sheetName);
  const day = text.match(/^(\d{1,2})\b/);
  if (month && day) return `2026-${String(month).padStart(2, "0")}-${day[1].padStart(2, "0")}`;
  return text;
}

function dateMatchesSheet(date, sheetName) {
  const month = sheetMonthNumber(sheetName);
  if (!month) return true;
  const match = String(date).match(/^2026-(\d{2})-\d{2}$/);
  return !match || Number(match[1]) === month;
}

function isWeekend(date) {
  const day = new Date(`${date}T00:00:00Z`).getUTCDay();
  return day === 0 || day === 6;
}

function displayDate(isoDate) {
  const match = String(isoDate).match(/^2026-(\d{2})-(\d{2})$/);
  return match ? `${match[2]}.${match[1]}` : isoDate;
}

function headerRows(cells, sheetName) {
  const rows = new Map();
  for (const [address, value] of cells.entries()) {
    const parsed = parseAddress(address);
    if (!parsed) continue;
    const row = rows.get(parsed.row) ?? new Map();
    row.set(parsed.col, String(value ?? ""));
    rows.set(parsed.row, row);
  }
  const headers = [];
  for (const [rowNumber, row] of rows.entries()) {
    const a = row.get(0)?.trim().toLowerCase();
    const c = row.get(2)?.trim().toLowerCase();
    if (a !== "name" || !c?.startsWith("total hours")) continue;
    const dates = new Map();
    for (const [col, value] of row.entries()) {
      if (col < 3) continue;
      const date = normalizeDate(value, sheetName);
      if (date && dateMatchesSheet(date, sheetName)) dates.set(col, date);
    }
    headers.push({ row: rowNumber, dates });
  }
  return headers.sort((a, b) => a.row - b.row);
}

function findHeader(headers, row) {
  let selected = null;
  for (const header of headers) {
    if (header.row < row) selected = header;
    else break;
  }
  return selected;
}

function valuesInRange(cells, range) {
  const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) return [];
  const [, startCol, startRowText, endCol, endRowText] = match;
  const values = [];
  for (let row = Number(startRowText); row <= Number(endRowText); row += 1) {
    for (let col = columnNameToIndex(startCol); col <= columnNameToIndex(endCol); col += 1) {
      values.push(String(cells.get(cellAddress(row, col)) ?? ""));
    }
  }
  return values;
}

function wildcardToRegex(pattern) {
  const escaped = pattern
    .split("*")
    .map((part) => part.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
    .join(".*");
  return new RegExp(`^${escaped}$`, "i");
}

function evaluateCountIfFormula(formula, cells) {
  let total = 0;
  const termRegex = /COUNTIF\(([A-Z]+\d+:[A-Z]+\d+),"([^"]+)"\)\s*\*\s*([0-9]+(?:\.[0-9]+)?)/g;
  for (const match of formula.matchAll(termRegex)) {
    const [, range, pattern, multiplierText] = match;
    const test = wildcardToRegex(pattern);
    total += valuesInRange(cells, range).filter((value) => test.test(value)).length * Number(multiplierText);
  }
  return Number.isInteger(total) ? String(total) : String(Number(total.toFixed(4)));
}

function replaceCellValue(cellXml, value) {
  if (/<v>[\s\S]*?<\/v>/.test(cellXml)) {
    return cellXml.replace(/<v>[\s\S]*?<\/v>/, `<v>${value}</v>`);
  }
  return cellXml.replace("</c>", `<v>${value}</v></c>`);
}

function updateCalcPr(workbookXml) {
  if (/<calcPr\b[^>]*\/>/.test(workbookXml)) {
    return workbookXml.replace(/<calcPr\b([^>]*)\/>/, (_match, attrs) => {
      let nextAttrs = attrs
        .replace(/\s(?:calcMode|fullCalcOnLoad|forceFullCalc)="[^"]*"/g, "")
        .trim();
      nextAttrs = nextAttrs ? ` ${nextAttrs}` : "";
      return `<calcPr${nextAttrs} calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>`;
    });
  }
  return workbookXml.replace(
    "</workbook>",
    '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>',
  );
}

function removeCalcChainRelationship(relsXml) {
  return relsXml.replace(/<Relationship\b[^>]*Type="[^"]*\/calcChain"[^>]*\/>/g, "");
}

function removeCalcChainContentType(contentTypesXml) {
  return contentTypesXml.replace(/<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, "");
}

await fs.rm(extractDir, { recursive: true, force: true });
await fs.mkdir(extractDir, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });
await fs.rm(outputPath, { force: true });
await execFileAsync("/usr/bin/unzip", ["-q", inputPath, "-d", extractDir]);

const sharedStringsPath = path.join(extractDir, "xl/sharedStrings.xml");
let sharedStringsXml = await fs.readFile(sharedStringsPath, "utf8");
let sharedStrings = parseSharedStrings(sharedStringsXml);
const off = ensureSharedString(sharedStringsXml, sharedStrings, "OFF");
sharedStringsXml = off.xml;
sharedStrings = off.strings;
const offIndex = off.index;
await fs.writeFile(sharedStringsPath, sharedStringsXml);

const workbookPath = path.join(extractDir, "xl/workbook.xml");
const workbookRelsPath = path.join(extractDir, "xl/_rels/workbook.xml.rels");
const contentTypesPath = path.join(extractDir, "[Content_Types].xml");
const sheets = parseWorkbookSheets(
  await fs.readFile(workbookPath, "utf8"),
  await fs.readFile(workbookRelsPath, "utf8"),
);

await fs.writeFile(workbookPath, updateCalcPr(await fs.readFile(workbookPath, "utf8")));
await fs.writeFile(workbookRelsPath, removeCalcChainRelationship(await fs.readFile(workbookRelsPath, "utf8")));
await fs.writeFile(contentTypesPath, removeCalcChainContentType(await fs.readFile(contentTypesPath, "utf8")));
await fs.rm(path.join(extractDir, "xl/calcChain.xml"), { force: true });

const changes = [];
const recalculated = [];

for (const sheet of sheets) {
  const sheetXml = await fs.readFile(sheet.path, "utf8");
  const cells = parseCells(sheetXml, sharedStrings);
  const headers = headerRows(cells, sheet.name);
  const changedRows = new Set();

  const updatedXml = sheetXml.replace(/<c\b([^>]*)>([\s\S]*?)<\/c>/g, (cellXml, attrsText, body) => {
    const attrs = parseAttrs(attrsText);
    if (!attrs.r) return cellXml;
    const parsed = parseAddress(attrs.r);
    if (!parsed || parsed.col < 3) return cellXml;

    const value = cellValue(attrs, body, sharedStrings);
    if (String(value).trim().toUpperCase() !== "VACATION") return cellXml;

    const header = findHeader(headers, parsed.row);
    const date = header?.dates.get(parsed.col);
    if (!date || !isWeekend(date)) return cellXml;

    cells.set(attrs.r, "OFF");
    changedRows.add(parsed.row);
    changes.push({
      sheet: sheet.name,
      cell: attrs.r,
      date: displayDate(date),
      row: parsed.row,
      employee: String(cells.get(cellAddress(parsed.row, 0)) ?? "").replace(/\s+/g, " ").trim(),
    });

    if (attrs.t === "s") {
      return cellXml.replace(/<v>[\s\S]*?<\/v>/, `<v>${offIndex}</v>`);
    }
    return `<c${attrsText} t="s"><v>${offIndex}</v></c>`;
  });

  const recalculatedXml = updatedXml.replace(/<c\b([^>]*)>([\s\S]*?)<\/c>/g, (cellXml, attrsText, body) => {
    const attrs = parseAttrs(attrsText);
    const parsed = attrs.r ? parseAddress(attrs.r) : null;
    if (!parsed || !changedRows.has(parsed.row)) return cellXml;
    const formula = body.match(/<f\b[^>]*>([\s\S]*?)<\/f>/)?.[1];
    if (!formula || !/COUNTIF\(/i.test(formula)) return cellXml;
    const value = evaluateCountIfFormula(decodeXml(formula), cells);
    recalculated.push({ sheet: sheet.name, cell: attrs.r, value });
    return replaceCellValue(cellXml, value);
  });

  await fs.writeFile(sheet.path, recalculatedXml);
}

await execFileAsync("/usr/bin/zip", ["-qr", outputPath, "."], { cwd: extractDir });

await fs.writeFile(
  path.join(outputDir, "vacation_weekend_changes.json"),
  JSON.stringify({ inputPath, outputPath, changes, recalculated }, null, 2),
);

console.log(JSON.stringify({ outputPath, changedCells: changes.length, recalculatedCells: recalculated.length, changes }, null, 2));
