import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

const inputPath =
  "/Users/dmitro/Downloads/azura_guide_v2/outputs/vacation_weekends_off/grafik_novyi1_weekend_vacation_off.xlsx";
const extractDir = "/Users/dmitro/Downloads/azura_guide_v2/spreadsheet_work/color_positions_xml";
const outputDir = "/Users/dmitro/Downloads/azura_guide_v2/outputs/colored_positions";
const outputPath = path.join(outputDir, "grafik_colored_positions.xlsx");

const positionColors = [
  ["Bar", "E2F0D9"],
  ["Chief", "CFE2F3"],
  ["Cleaner", "EADCF8"],
  ["Cold", "DDEBF7"],
  ["Cook asisstant", "F4CCCC"],
  ["Cook assistant", "FCE4D6"],
  ["Dessert", "FFF2CC"],
  ["Driver", "D9D2E9"],
  ["Hostess", "D0E0E3"],
  ["Hot", "F8CBAD"],
  ["Pasta", "FFE699"],
  ["Pizza", "F4B183"],
  ["Runner", "C6E0B4"],
  ["Somellier", "E4DFEC"],
  ["Waiter", "B7DEE8"],
];

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
    if (!target) continue;
    sheets.push({
      name: attrs.name,
      path: path.join(extractDir, target.startsWith("/") ? target.slice(1) : `xl/${target}`),
    });
  }
  return sheets;
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

function escapeRegex(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function matchedPosition(value) {
  const text = String(value ?? "").trim();
  for (const [position] of [...positionColors].sort((a, b) => b[0].length - a[0].length)) {
    const pattern = new RegExp(`^${escapeRegex(position)}\\s+\\d{1,2}:\\d{2}\\s*-\\s*\\d{1,2}:\\d{2}`, "i");
    if (pattern.test(text)) return position;
  }
  return null;
}

function getSection(xml, tagName) {
  const regex = new RegExp(`<${tagName}\\b([^>]*)>([\\s\\S]*?)<\\/${tagName}>`);
  const match = xml.match(regex);
  if (!match) throw new Error(`Missing ${tagName} section`);
  return {
    full: match[0],
    attrsText: match[1],
    body: match[2],
  };
}

function updateCountAttr(fullXml, count) {
  if (/\bcount="[^"]*"/.test(fullXml)) {
    return fullXml.replace(/\bcount="[^"]*"/, `count="${count}"`);
  }
  return fullXml.replace(/^<([A-Za-z0-9:]+)\b/, `<$1 count="${count}"`);
}

function parseFills(fillsBody) {
  return [...fillsBody.matchAll(/<fill\b[^>]*>[\s\S]*?<\/fill>/g)].map((match) => match[0]);
}

function parseCellXfs(cellXfsBody) {
  return [...cellXfsBody.matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)].map((match) => match[0]);
}

function addOrGetFill(fills, rgb) {
  const normalized = `FF${rgb.toUpperCase()}`;
  const existing = fills.findIndex((fill) => fill.includes(`rgb="${normalized}"`));
  if (existing >= 0) return { fills, fillId: existing };
  return {
    fills: [
      ...fills,
      `<fill><patternFill patternType="solid"><fgColor rgb="${normalized}"/><bgColor indexed="64"/></patternFill></fill>`,
    ],
    fillId: fills.length,
  };
}

function withAttr(xml, name, value) {
  const regex = new RegExp(`\\s${name}="[^"]*"`);
  if (regex.test(xml)) return xml.replace(regex, ` ${name}="${value}"`);
  return xml.replace(/\/>$|>[\s\S]*?<\/xf>$/, (tail) => ` ${name}="${value}"${tail}`);
}

function addOrGetXf(cellXfs, baseStyleId, fillId) {
  const base = cellXfs[baseStyleId] ?? cellXfs[0];
  let next = withAttr(base, "fillId", fillId);
  next = withAttr(next, "applyFill", 1);
  const existing = cellXfs.indexOf(next);
  if (existing >= 0) return { cellXfs, styleId: existing };
  return { cellXfs: [...cellXfs, next], styleId: cellXfs.length };
}

function replaceSection(xml, tagName, section, nextItems) {
  const open = section.full.match(new RegExp(`^<${tagName}\\b[^>]*>`))[0];
  const close = `</${tagName}>`;
  const updated = updateCountAttr(`${open}${nextItems.join("")}${close}`, nextItems.length);
  return xml.replace(section.full, updated);
}

function updateCellStyle(attrsText, styleId) {
  if (/\bs="[^"]*"/.test(attrsText)) {
    return attrsText.replace(/\bs="[^"]*"/, `s="${styleId}"`);
  }
  return `${attrsText} s="${styleId}"`;
}

await fs.rm(extractDir, { recursive: true, force: true });
await fs.mkdir(extractDir, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });
await fs.rm(outputPath, { force: true });
await execFileAsync("/usr/bin/unzip", ["-q", inputPath, "-d", extractDir]);

const sharedStrings = parseSharedStrings(await fs.readFile(path.join(extractDir, "xl/sharedStrings.xml"), "utf8"));
const workbookXml = await fs.readFile(path.join(extractDir, "xl/workbook.xml"), "utf8");
const workbookRelsXml = await fs.readFile(path.join(extractDir, "xl/_rels/workbook.xml.rels"), "utf8");
const sheets = parseWorkbookSheets(workbookXml, workbookRelsXml);

const stylesPath = path.join(extractDir, "xl/styles.xml");
let stylesXml = await fs.readFile(stylesPath, "utf8");
const fillsSection = getSection(stylesXml, "fills");
const cellXfsSection = getSection(stylesXml, "cellXfs");
let fills = parseFills(fillsSection.body);
let cellXfs = parseCellXfs(cellXfsSection.body);

const fillIdsByPosition = new Map();
for (const [position, color] of positionColors) {
  const result = addOrGetFill(fills, color);
  fills = result.fills;
  fillIdsByPosition.set(position, result.fillId);
}

const styleCache = new Map();
const counts = Object.fromEntries(positionColors.map(([position]) => [position, 0]));

for (const sheet of sheets) {
  const sheetXml = await fs.readFile(sheet.path, "utf8");
  const updatedXml = sheetXml.replace(/<c\b([^>]*)>([\s\S]*?)<\/c>/g, (cellXml, attrsText, body) => {
    const attrs = parseAttrs(attrsText);
    const value = cellValue(attrs, body, sharedStrings);
    const position = matchedPosition(value);
    if (!position) return cellXml;

    const baseStyleId = Number(attrs.s ?? 0);
    const cacheKey = `${baseStyleId}\u0000${position}`;
    let styleId = styleCache.get(cacheKey);
    if (styleId === undefined) {
      const result = addOrGetXf(cellXfs, baseStyleId, fillIdsByPosition.get(position));
      cellXfs = result.cellXfs;
      styleId = result.styleId;
      styleCache.set(cacheKey, styleId);
    }
    counts[position] += 1;
    return `<c${updateCellStyle(attrsText, styleId)}>${body}</c>`;
  });
  await fs.writeFile(sheet.path, updatedXml);
}

stylesXml = replaceSection(stylesXml, "fills", fillsSection, fills);
const nextCellXfsSection = getSection(stylesXml, "cellXfs");
stylesXml = replaceSection(stylesXml, "cellXfs", nextCellXfsSection, cellXfs);
await fs.writeFile(stylesPath, stylesXml);

await execFileAsync("/usr/bin/zip", ["-qr", outputPath, "."], { cwd: extractDir });

await fs.writeFile(
  path.join(outputDir, "position_color_counts.json"),
  JSON.stringify({ inputPath, outputPath, positionColors: Object.fromEntries(positionColors), counts }, null, 2),
);

console.log(JSON.stringify({ outputPath, counts, colors: Object.fromEntries(positionColors) }, null, 2));
