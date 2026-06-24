import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

const inputPath = "/Users/dmitro/Downloads/ГРАФИК НОВЫЙ.xlsx";
const extractDir = "/Users/dmitro/Downloads/azura_guide_v2/spreadsheet_work/xlsx_xml";
const outputDir = "/Users/dmitro/Downloads/azura_guide_v2/outputs/schedule_break_fix";
const outputPath = path.join(outputDir, "grafik_novyi_break_fixed.xlsx");

const targetShift = {
  pattern: "10:00-22:00",
  hours: "11",
};

function decodeXml(text) {
  return text
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function encodeXmlText(text) {
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
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

function toHours(value) {
  const [hh, mm] = value.split(":").map(Number);
  return hh + mm / 60;
}

function payableHours(start, end) {
  const startHours = toHours(start);
  let endHours = toHours(end);
  if (endHours <= startHours) {
    endHours += 24;
  }
  const hours = Math.max(0, endHours - startHours - 1);
  return Number.isInteger(hours) ? String(hours) : String(Number(hours.toFixed(4)));
}

function updateShiftMultipliers(formula) {
  let changed = false;
  const updated = formula.replace(
    /COUNTIF\(([^)]*?"\*(\d{1,2}:\d{2})-(\d{1,2}:\d{2})\*"[^)]*)\)(?:\s*\*\s*([0-9]+(?:\.[0-9]+)?))?/g,
    (_match, countIfArgs, start, end, existingMultiplier) => {
      const nextMultiplier = payableHours(start, end);
      if (existingMultiplier !== nextMultiplier) {
        changed = true;
      }
      return `COUNTIF(${countIfArgs})*${nextMultiplier}`;
    },
  );

  return { formula: updated, changed };
}

function firstCountIfRange(formula) {
  const match = formula.match(/COUNTIF\(([A-Z]+\d+:[A-Z]+\d+),/);
  return match?.[1] ?? null;
}

function ensureTargetShift(formula) {
  if (!/COUNTIF\([^)]*"\*\d{1,2}:\d{2}-\d{1,2}:\d{2}\*"/.test(formula)) {
    return { formula, changed: false };
  }
  if (formula.includes(`*${targetShift.pattern}*`)) {
    return { formula, changed: false };
  }
  const range = firstCountIfRange(formula);
  if (!range) {
    return { formula, changed: false };
  }
  return {
    formula: `${formula}+COUNTIF(${range},"*${targetShift.pattern}*")*${targetShift.hours}`,
    changed: true,
  };
}

function updateFormula(formula) {
  const multiplierResult = updateShiftMultipliers(formula);
  const targetShiftResult = ensureTargetShift(multiplierResult.formula);
  return {
    formula: targetShiftResult.formula,
    changed: multiplierResult.changed || targetShiftResult.changed,
  };
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

function cellAddress(column, row) {
  return `${columnIndexToName(column)}${row}`;
}

function parseCells(sheetXml, sharedStrings) {
  const cells = new Map();
  const cellRegex = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
  for (const match of sheetXml.matchAll(cellRegex)) {
    const attrs = match[1];
    const body = match[2];
    const address = attrs.match(/\br="([^"]+)"/)?.[1];
    if (!address) {
      continue;
    }
    const valueRaw = body.match(/<v>([\s\S]*?)<\/v>/)?.[1];
    if (valueRaw === undefined) {
      cells.set(address, "");
      continue;
    }
    if (/\bt="s"/.test(attrs)) {
      cells.set(address, sharedStrings[Number(valueRaw)] ?? "");
    } else if (/\bt="str"/.test(attrs) || /\bt="inlineStr"/.test(attrs)) {
      cells.set(address, decodeXml(valueRaw));
    } else {
      cells.set(address, decodeXml(valueRaw));
    }
  }
  return cells;
}

function wildcardToRegex(pattern) {
  const escaped = pattern
    .split("*")
    .map((part) => part.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
    .join(".*");
  return new RegExp(`^${escaped}$`, "i");
}

function valuesInRange(cells, range) {
  const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) {
    return [];
  }
  const [, startCol, startRowText, endCol, endRowText] = match;
  const startColIndex = columnNameToIndex(startCol);
  const endColIndex = columnNameToIndex(endCol);
  const startRow = Number(startRowText);
  const endRow = Number(endRowText);
  const values = [];
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startColIndex; col <= endColIndex; col += 1) {
      values.push(String(cells.get(cellAddress(col, row)) ?? ""));
    }
  }
  return values;
}

function evaluateCountIfSum(formula, cells) {
  let total = 0;
  const termRegex =
    /COUNTIF\(([A-Z]+\d+:[A-Z]+\d+),"([^"]+)"\)\s*\*\s*([0-9]+(?:\.[0-9]+)?)/g;
  for (const match of formula.matchAll(termRegex)) {
    const [, range, pattern, multiplierText] = match;
    const test = wildcardToRegex(pattern);
    const count = valuesInRange(cells, range).filter((value) => test.test(value)).length;
    total += count * Number(multiplierText);
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
  return relsXml.replace(
    /<Relationship\b[^>]*Type="[^"]*\/calcChain"[^>]*\/>/g,
    "",
  );
}

function removeCalcChainContentType(contentTypesXml) {
  return contentTypesXml.replace(
    /<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g,
    "",
  );
}

await fs.rm(extractDir, { recursive: true, force: true });
await fs.mkdir(extractDir, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });
await fs.rm(outputPath, { force: true });

await execFileAsync("/usr/bin/unzip", ["-q", inputPath, "-d", extractDir]);

const sharedStringsPath = path.join(extractDir, "xl/sharedStrings.xml");
const sharedStrings = parseSharedStrings(await fs.readFile(sharedStringsPath, "utf8"));

const workbookPath = path.join(extractDir, "xl/workbook.xml");
const workbookRelsPath = path.join(extractDir, "xl/_rels/workbook.xml.rels");
const contentTypesPath = path.join(extractDir, "[Content_Types].xml");

await fs.writeFile(workbookPath, updateCalcPr(await fs.readFile(workbookPath, "utf8")));
await fs.writeFile(
  workbookRelsPath,
  removeCalcChainRelationship(await fs.readFile(workbookRelsPath, "utf8")),
);
await fs.writeFile(
  contentTypesPath,
  removeCalcChainContentType(await fs.readFile(contentTypesPath, "utf8")),
);
await fs.rm(path.join(extractDir, "xl/calcChain.xml"), { force: true });

const sheetFiles = ["sheet1.xml", "sheet2.xml", "sheet3.xml", "sheet4.xml", "sheet5.xml"];
const changes = [];

for (const sheetFile of sheetFiles) {
  const sheetPath = path.join(extractDir, "xl/worksheets", sheetFile);
  const originalXml = await fs.readFile(sheetPath, "utf8");
  const cells = parseCells(originalXml, sharedStrings);

  const updatedXml = originalXml.replace(/<c\b([^>]*)>([\s\S]*?)<\/c>/g, (cellXml, attrs, body) => {
    const address = attrs.match(/\br="([^"]+)"/)?.[1];
    const formulaMatch = body.match(/<f\b([^>]*)>([\s\S]*?)<\/f>/);
    if (!address || !formulaMatch) {
      return cellXml;
    }

    const formula = decodeXml(formulaMatch[2]);
    const result = updateFormula(formula);
    if (!result.changed) {
      return cellXml;
    }

    const encodedFormula = encodeXmlText(result.formula);
    const withFormula = cellXml.replace(
      /<f\b([^>]*)>[\s\S]*?<\/f>/,
      `<f${formulaMatch[1]}>${encodedFormula}</f>`,
    );
    const cachedValue = evaluateCountIfSum(result.formula, cells);
    const withValue = replaceCellValue(withFormula, cachedValue);
    changes.push({ sheetFile, address, cachedValue, before: formula, after: result.formula });
    return withValue;
  });

  await fs.writeFile(sheetPath, updatedXml);
}

await execFileAsync("/usr/bin/zip", ["-qr", outputPath, "."], { cwd: extractDir });

const verification = changes.map(({ sheetFile, address, cachedValue }) => ({
  sheetFile,
  address,
  cachedValue,
}));
await fs.writeFile(
  path.join(outputDir, "formula_changes.json"),
  JSON.stringify(verification, null, 2),
);

console.log(`Changed formula cells: ${changes.length}`);
console.log(JSON.stringify(verification.slice(0, 20), null, 2));
console.log(`Output: ${outputPath}`);
