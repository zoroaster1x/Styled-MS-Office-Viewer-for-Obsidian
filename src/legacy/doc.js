/*
 * Styled MS Office Viewer, an Obsidian plugin that renders office documents
 * (xlsx, docx, pptx and their relatives) with their real styling, read only.
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

// Word 97 (.doc) text, read straight out of the compound file.
//
// The document stream carries a piece table, not a plain string: a list of
// character ranges, each pointing at either 8 bit (cp1252) or 16 bit (UTF-16)
// text somewhere in the same stream. Reading the pieces in order rebuilds the
// text; a file that was never edited straightforwardly (fComplex clear) points
// at one contiguous run instead. Formatting runs are not decoded, so this is a
// text view: paragraphs and their words, no fonts, colours or tables.
//
// Page setup also lives in a table this reader does not walk, so the page is an
// A4 sheet with two centimetre margins; every .doc opens without a crash and
// reads top to bottom.

const { openCompound } = require("./ole");

const WORD_SIGNATURE = 0xa5ec;
const FIB_OFFSETS = {
  flags: 0x0a,
  fcMin: 0x18,
  ccpText: 0x4c,
  fcClx: 0x01a2,
  lcbClx: 0x01a6,
};

const CP1252 = (() => {
  // The 0x80 to 0x9f block differs from latin-1; everything else is the same.
  const high = {
    0x80: "\u20ac", 0x82: "\u201a", 0x83: "\u0192", 0x84: "\u201e", 0x85: "\u2026", 0x86: "\u2020", 0x87: "\u2021",
    0x88: "\u02c6", 0x89: "\u2030", 0x8a: "\u0160", 0x8b: "\u2039", 0x8c: "\u0152", 0x8e: "\u017d",
    0x91: "\u2018", 0x92: "\u2019", 0x93: "\u201c", 0x94: "\u201d", 0x95: "\u2022", 0x96: "\u2013", 0x97: "\u2014",
    0x98: "\u02dc", 0x99: "\u2122", 0x9a: "\u0161", 0x9b: "\u203a", 0x9c: "\u0153", 0x9e: "\u017e", 0x9f: "\u0178",
  };
  const table = new Array(256);
  for (let i = 0; i < 256; i++) table[i] = high[i] || String.fromCharCode(i);
  return table;
})();

function decodeCp1252(bytes) {
  let out = "";
  for (let i = 0; i < bytes.length; i++) out += CP1252[bytes[i]];
  return out;
}

function decodeUtf16(bytes) {
  let out = "";
  for (let i = 0; i + 1 < bytes.length; i += 2) out += String.fromCharCode(bytes[i] | (bytes[i + 1] << 8));
  return out;
}

function readDoc(bytes) {
  const compound = openCompound(bytes);
  const word = compound.stream("WordDocument");
  if (!word || word.length < 0x200) throw new Error("This .doc file has no document stream.");
  const view = new DataView(word.buffer, word.byteOffset, word.byteLength);
  if (view.getUint16(0, true) !== WORD_SIGNATURE) throw new Error("This .doc file has an unexpected header.");
  const flags = view.getUint16(FIB_OFFSETS.flags, true);
  const complex = (flags & 0x0004) !== 0;
  const useOneTable = (flags & 0x0200) !== 0;
  const table = compound.stream(useOneTable ? "1Table" : "0Table") || compound.stream("0Table") || compound.stream("1Table");
  const fcMin = view.getUint32(FIB_OFFSETS.fcMin, true);
  const ccpText = view.getUint32(FIB_OFFSETS.ccpText, true);
  const fcClx = view.getUint32(FIB_OFFSETS.fcClx, true);
  const lcbClx = view.getUint32(FIB_OFFSETS.lcbClx, true);

  let raw = "";
  if (!complex || !table || !lcbClx) {
    // One contiguous run of 8 bit characters.
    const length = Math.min(ccpText || word.length - fcMin, Math.max(0, word.length - fcMin));
    raw = decodeCp1252(word.subarray(fcMin, fcMin + length));
  } else {
    raw = readPieces(word, table, fcClx, lcbClx);
  }
  const text = cleanText(raw);
  const paragraphs = text.split("\n").map((line) => line.replace(/\s+$/, ""));
  const body = [];
  for (const line of paragraphs) {
    if (!line) {
      body.push({ type: "p", runs: [], props: {}, style: null });
      continue;
    }
    // The renderer expects plain text as { type: "text" }; a { type: "run" }
    // wrapper means a nested run list and would recurse into nothing.
    body.push({
      type: "p",
      runs: [{ type: "text", props: {}, text: line }],
      props: {},
      style: null,
    });
  }
  while (body.length && !body[body.length - 1].runs.length) body.pop();

  return {
    kind: "doc",
    body,
    section: {
      pageWidthTw: 11906,
      pageHeightTw: 16838,
      marginTopTw: 1134,
      marginRightTw: 1134,
      marginBottomTw: 1134,
      marginLeftTw: 1134,
      estimated: true,
    },
    styles: { paragraph: new Map(), character: new Map(), table: new Map(), docDefaults: null },
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
    mediaCache: null,
    properties: {},
    mediaUrl() {
      return null;
    },
    paragraphCount: body.filter((block) => block.runs && block.runs.length).length,
    characterCount: text.replace(/\n/g, "").length,
  };
}

// The piece table: run of character positions, each mapped to text in the
// document stream. Positions are inclusive starts and one past the end.
function readPieces(word, table, fcClx, lcbClx) {
  const end = Math.min(table.length, fcClx + lcbClx);
  let at = fcClx;
  let plc = null;
  while (at < end) {
    const token = table[at];
    if (token === 0x01) {
      if (at + 3 > end) break;
      const length = table[at + 1] | (table[at + 2] << 8);
      at += 3 + length;
    } else if (token === 0x02) {
      if (at + 5 > end) break;
      const length = (table[at + 1] | (table[at + 2] << 8) | (table[at + 3] << 16) | (table[at + 4] << 24)) >>> 0;
      plc = { offset: at + 5, length: Math.min(length, end - (at + 5)) };
      break;
    } else {
      break;
    }
  }
  if (!plc || plc.length < 12) return "";
  const view = new DataView(table.buffer, table.byteOffset, table.byteLength);
  const count = Math.floor((plc.length - 4) / 12);
  if (count <= 0) return "";
  const cps = [];
  for (let i = 0; i <= count; i++) cps.push(view.getUint32(plc.offset + i * 4, true) >>> 0);
  const pcdsAt = plc.offset + (count + 1) * 4;
  let out = "";
  for (let i = 0; i < count; i++) {
    const characters = cps[i + 1] - cps[i];
    if (characters <= 0 || characters > 32 * 1024 * 1024) continue;
    const fc = view.getUint32(pcdsAt + i * 8 + 2, true);
    const compressed = (fc & 0x40000000) !== 0;
    const offset = fc & 0x3fffffff;
    if (compressed) {
      const start = offset >> 1;
      if (start >= word.length) continue;
      out += decodeCp1252(word.subarray(start, Math.min(word.length, start + characters)));
    } else {
      if (offset + 2 > word.length) continue;
      out += decodeUtf16(word.subarray(offset, Math.min(word.length, offset + characters * 2)));
    }
  }
  return out;
}

// Word stores control characters in the text: field marks, cell ends, picture
// anchors and the two line separators. Keep the reader's meaning and drop the
// mechanics.
function cleanText(raw) {
  const withFields = raw.replace(/\u0013[\s\S]*?\u0014/g, "").replace(/[\u0013\u0014\u0015\u0001\u0002\u0008]/g, "");
  const mapped = withFields
    .replace(/\u000b/g, "\n")
    .replace(/[\r\u000c\u0007]/g, "\n")
    .replace(/\u001e/g, "-")
    .replace(/\u001f/g, "")
    .replace(/\u00a0/g, " ")
    .replace(/\u0000/g, "");
  return mapped.replace(/\n{3,}/g, "\n\n").replace(/^\n+/, "");
}

module.exports = { readDoc, cleanText };
