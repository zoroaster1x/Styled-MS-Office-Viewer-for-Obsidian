/*
 * Styled MS Office Viewer, an Obsidian plugin that renders office documents (xlsx,
 * docx, pptx and their relatives) with their real styling, read only.
 *
 * Copyright (C) 2026 Zoroaster1x
 *
 * This program is free software: you can redistribute it and/or modify it under
 * the terms of the GNU General Public License as published by the Free Software
 * Foundation, either version 3 of the License, or (at your option) any later
 * version.
 *
 * This program is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
 * details.
 *
 * You should have received a copy of the GNU General Public License along with
 * this program. If not, see <https://www.gnu.org/licenses/>.
 */

"use strict";

// CSV and TSV become a sheet model, so they open in the same grid as a
// workbook: header row, sorting by click, wrapping, the lot.

const { key } = require("../shared/addr");
const { parseStyles } = require("../spreadsheet/styles");

const DELIMITERS = [",", "\t", ";", "|"];

// RFC 4180 quoting, plus a delimiter guess when the extension does not say.
function parseDelimited(text, delimiter) {
  const content = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  const sep = delimiter || guessDelimiter(content);
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  let i = 0;
  while (i < content.length) {
    const ch = content[i];
    if (quoted) {
      if (ch === '"') {
        if (content[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"' && field === "") {
      quoted = true;
      i++;
      continue;
    }
    if (ch === sep) {
      row.push(field);
      field = "";
      i++;
      continue;
    }
    if (ch === "\r") {
      if (content[i + 1] === "\n") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      i++;
      continue;
    }
    if (ch === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      i++;
      continue;
    }
    field += ch;
    i++;
  }
  if (field !== "" || row.length) {
    row.push(field);
    rows.push(row);
  }
  // A trailing newline produces one empty row; drop it.
  while (rows.length && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === "") rows.pop();
  return rows;
}

// The delimiter that yields the most consistent column count wins.
function guessDelimiter(text) {
  const sample = text.slice(0, 4000);
  let best = { delimiter: ",", score: -1 };
  for (const delimiter of DELIMITERS) {
    const rows = sample.split(/\r?\n/).filter((line) => line.length).slice(0, 20);
    if (!rows.length) continue;
    const counts = rows.map((line) => line.split(delimiter).length);
    const first = counts[0];
    if (first < 2) continue;
    const consistent = counts.filter((count) => count === first).length / counts.length;
    const score = consistent * 100 + first;
    if (score > best.score) best = { delimiter, score };
  }
  return best.delimiter;
}

function isNumeric(value) {
  if (value === "") return false;
  if (!/^[-+]?[\d,]*\.?\d+(?:[eE][-+]?\d+)?$/.test(value)) return false;
  const cleaned = value.replace(/,/g, "");
  return !isNaN(Number(cleaned));
}

// Produces the same model the spreadsheet renderer consumes for .xlsx.
function buildSheetModel(rows, name) {
  const styles = parseStyles(null, null);
  let maxCol = 0;
  const cells = new Map();
  const modelRows = new Map();
  rows.forEach((row, r) => {
    const rowIndex = r + 1;
    modelRows.set(rowIndex, {
      r: rowIndex,
      ht: null,
      hidden: false,
      outline: 0,
      collapsed: false,
      s: null,
      customFormat: false,
    });
    row.forEach((text, c) => {
      const colIndex = c + 1;
      if (colIndex > maxCol) maxCol = colIndex;
      if (text === "") return;
      const numeric = isNumeric(text);
      cells.set(key(rowIndex, colIndex), numeric
        ? { v: Number(text.replace(/,/g, "")), t: "n", s: 0 }
        : { v: text, t: "s", s: 0 });
    });
  });
  const rowCount = Math.max(1, rows.length);
  const colCount = Math.max(1, maxCol);
  return {
    name: name || "Sheet1",
    dims: { r1: 1, c1: 1, r2: rowCount, c2: colCount },
    defaultRowHeightPt: 15,
    defaultRowHeightPx: 20,
    defaultColWidthChars: null,
    cols: [],
    rows: modelRows,
    cells,
    merges: [],
    freeze: null,
    autoFilter: null,
    hyperlinks: new Map(),
    showGridLines: true,
    outlinePr: { summaryBelow: true, summaryRight: true },
    conditional: [],
    tabColor: null,
    styles,
  };
}

module.exports = {
  parseDelimited,
  guessDelimiter,
  buildSheetModel,
  isNumeric,
};
