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

// Address and text helpers that are not XML specific.

function colToLetter(col) {
  let s = "";
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function letterToCol(letters) {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    const code = letters.charCodeAt(i);
    if (code < 65 || code > 90) continue;
    col = col * 26 + (code - 64);
  }
  return col;
}

function parseCellRef(ref) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(ref).trim());
  if (!m) return null;
  return { c: letterToCol(m[1].toUpperCase()), r: parseInt(m[2], 10) };
}

function parseRange(ref) {
  const parts = String(ref).split(":");
  const a = parseCellRef(parts[0]);
  const b = parts.length > 1 ? parseCellRef(parts[1]) : a;
  if (!a || !b) return null;
  return {
    r1: Math.min(a.r, b.r),
    c1: Math.min(a.c, b.c),
    r2: Math.max(a.r, b.r),
    c2: Math.max(a.c, b.c),
  };
}

// Splits "Sheet name!A1:B2" into { sheet, range }. Handles quoted names.
function splitSheetRef(ref) {
  const bang = String(ref).lastIndexOf("!");
  if (bang === -1) return { sheet: null, ref: String(ref) };
  let sheet = String(ref).slice(0, bang).trim();
  if (sheet.startsWith("'") && sheet.endsWith("'")) {
    sheet = sheet.slice(1, -1).replace(/''/g, "'");
  }
  return { sheet, ref: String(ref).slice(bang + 1) };
}

// Excel column width (in characters) to pixels, using the usual 7px digit width.
function pxFromCharWidth(chars) {
  return Math.round(chars * 7) + 5;
}

// Numbers first, then text, both in a human order.
function naturalCompare(a, b) {
  const sa = String(a == null ? "" : a);
  const sb = String(b == null ? "" : b);
  const na = sa.trim() === "" ? NaN : Number(sa);
  const nb = sb.trim() === "" ? NaN : Number(sb);
  const aNum = !isNaN(na);
  const bNum = !isNaN(nb);
  if (aNum && bNum) return na - nb;
  if (aNum) return -1;
  if (bNum) return 1;
  return sa.localeCompare(sb, undefined, { numeric: true, sensitivity: "base" });
}

function key(r, c) {
  return r + ":" + c;
}

// Joins the text of every w:t / a:t / text:p descendant, skipping phonetic
// runs, which the spreadsheet and docx parsers both need.
function textOfRuns(el) {
  let out = "";
  const walk = (node) => {
    const list = node.children || [];
    for (let i = 0; i < list.length; i++) {
      const child = list[i];
      const tag = (child.tagName || "").replace(/^.*:/, "");
      if (tag === "rPh" || tag === "phoneticPr") continue;
      if (tag === "t") out += child.textContent || "";
      else walk(child);
    }
  };
  walk(el);
  return out;
}

module.exports = {
  colToLetter,
  letterToCol,
  parseCellRef,
  parseRange,
  splitSheetRef,
  pxFromCharWidth,
  naturalCompare,
  key,
  textOfRuns,
};
