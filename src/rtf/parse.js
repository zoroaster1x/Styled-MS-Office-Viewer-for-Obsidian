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

// Rich Text Format reader.
//
// RTF is text: groups in braces, control words after a backslash. The reader
// walks it once, keeping a stack of character and paragraph state, and emits
// the same block model the docx renderer draws. Font and colour tables are
// read in a first pass so the body can resolve \f and \cf indices.

const DEFAULT_SECTION = {
  pageWidthTw: 12240,
  pageHeightTw: 15840,
  marginTopTw: 1440,
  marginRightTw: 1440,
  marginBottomTw: 1440,
  marginLeftTw: 1440,
  headerTw: 708,
  footerTw: 708,
  columns: 1,
  columnSpaceTw: 708,
};

const ALIGNMENT = { ql: "left", qc: "center", qr: "right", qj: "justify", qd: "justify" };

function readRtf(text) {
  if (!text || text.indexOf("{\\rtf") !== 0) {
    throw new Error("This file is not a readable RTF document.");
  }
  const tables = readTables(text);
  const reader = new RtfReader(text, tables);
  const body = reader.read();
  return {
    kind: "rtf",
    body,
    section: reader.section,
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
  };
}

// First pass: font names by index and the colour table.
function readTables(text) {
  const fonts = new Map();
  const colors = [];
  let i = 0;
  while (i < text.length) {
    const at = text.indexOf("\\fonttbl", i);
    if (at === -1) break;
    const group = text.slice(at, findGroupEnd(text, at));
    const fontPattern = /\\f(\d+)[^;]*?([^;\\{}]+);/g;
    let match;
    while ((match = fontPattern.exec(group)) !== null) {
      fonts.set(Number(match[1]), cleanFontName(match[2]));
    }
    i = at + group.length;
  }
  i = 0;
  while (i < text.length) {
    const at = text.indexOf("\\colortbl", i);
    if (at === -1) break;
    const group = text.slice(at, findGroupEnd(text, at));
    const entries = group.split(";");
    for (const entry of entries) {
      if (!/\\red|\\green|\\blue/.test(entry)) continue;
      const r = parseInt((/\\red(\d+)/.exec(entry) || [])[1] || "0", 10);
      const g = parseInt((/\\green(\d+)/.exec(entry) || [])[1] || "0", 10);
      const b = parseInt((/\\blue(\d+)/.exec(entry) || [])[1] || "0", 10);
      colors.push({ r, g, b });
    }
    i = at + group.length;
  }
  return { fonts, colors };
}

// Finds the closing brace of the group that starts at or after `from`.
function findGroupEnd(text, from) {
  let depth = 0;
  for (let i = from; i < text.length; i++) {
    const ch = text[i];
    if (ch === "\\" && (text[i + 1] === "{" || text[i + 1] === "}" || text[i + 1] === "\\")) {
      i++;
      continue;
    }
    if (ch === "{") depth++;
    else if (ch === "}") {
      depth--;
      if (depth === 0) return i + 1;
    }
  }
  return text.length;
}

function cleanFontName(value) {
  return String(value || "").replace(/[{}]/g, "").trim();
}

class RtfReader {
  constructor(text, tables) {
    this.text = text;
    this.pos = 0;
    this.tables = tables;
    this.section = Object.assign({}, DEFAULT_SECTION);
    this.stack = [];
    this.charProps = {};
    this.paragraphProps = {};
    this.fontIndex = 0;
    this.colorIndex = 0;
    this.highlightIndex = 0;
    this.current = newParagraph();
    this.blocks = [];
    this.target = this.blocks;
    this.row = null;
    this.inTable = false;
  }

  read() {
    while (this.pos < this.text.length) {
      const ch = this.text[this.pos];
      if (ch === "\\") {
        this.controlWord();
      } else if (ch === "{") {
        this.openGroup();
      } else if (ch === "}") {
        this.closeGroup();
      } else if (ch === "\r" || ch === "\n") {
        this.pos++;
      } else {
        this.textRun(ch);
      }
    }
    this.flushParagraph();
    return this.blocks;
  }

  openGroup() {
    this.stack.push({
      charProps: Object.assign({}, this.charProps),
      paragraphProps: Object.assign({}, this.paragraphProps),
      fontIndex: this.fontIndex,
      colorIndex: this.colorIndex,
      highlightIndex: this.highlightIndex,
      inTable: this.inTable,
    });
    this.pos++;
  }

  closeGroup() {
    const saved = this.stack.pop();
    if (!saved) {
      this.pos++;
      return;
    }
    this.charProps = saved.charProps;
    this.paragraphProps = saved.paragraphProps;
    this.fontIndex = saved.fontIndex;
    this.colorIndex = saved.colorIndex;
    this.highlightIndex = saved.highlightIndex;
    this.inTable = saved.inTable;
    this.pos++;
  }

  controlWord() {
    this.pos++;
    const start = this.pos;
    while (this.pos < this.text.length && /[a-zA-Z]/.test(this.text[this.pos])) this.pos++;
    const word = this.text.slice(start, this.pos);
    let param = null;
    if (this.text[this.pos] === "-" || /[0-9]/.test(this.text[this.pos] || "")) {
      const numStart = this.pos;
      if (this.text[this.pos] === "-") this.pos++;
      while (this.pos < this.text.length && /[0-9]/.test(this.text[this.pos])) this.pos++;
      param = parseInt(this.text.slice(numStart, this.pos), 10);
    }
    if (this.text[this.pos] === " ") this.pos++;
    if (this.isSkippedDestination(word)) {
      this.skipGroup();
      return;
    }
    this.applyControl(word, param);
  }

  // Font, colour, style and metadata tables are read by the first pass.
  isSkippedDestination(word) {
    return word === "fonttbl" || word === "colortbl" || word === "stylesheet" ||
      word === "info" || word === "listtable" || word === "listoverridetable" ||
      word === "generator" || word === "revtbl" || word === "rsidtbl" ||
      word === "latentstyles" || word === "datastore" || word === "themedata" ||
      word === "colorschememapping" || word === "xmlnstbl";
  }

  skipGroup() {
    // The control word sits inside a group, so consume to the matching brace.
    let depth = 0;
    while (this.pos < this.text.length) {
      const ch = this.text[this.pos];
      if (ch === "\\" && (this.text[this.pos + 1] === "{" || this.text[this.pos + 1] === "}")) {
        this.pos += 2;
        continue;
      }
      if (ch === "{") depth++;
      else if (ch === "}") {
        if (depth === 0) {
          this.pos++;
          return;
        }
        depth--;
      }
      this.pos++;
    }
  }

  applyControl(word, param) {
    switch (word) {
      case "rtf":
      case "ansi":
      case "ansicpg":
      case "deff":
      case "adeff":
      case "uc":
      case "paperw":
        if (word === "paperw") this.section.pageWidthTw = param || DEFAULT_SECTION.pageWidthTw;
        break;
      case "paperh":
        this.section.pageHeightTw = param || DEFAULT_SECTION.pageHeightTw;
        break;
      case "margl":
        this.section.marginLeftTw = param || 0;
        break;
      case "margr":
        this.section.marginRightTw = param || 0;
        break;
      case "margt":
        this.section.marginTopTw = param || 0;
        break;
      case "margb":
        this.section.marginBottomTw = param || 0;
        break;
      case "f":
        this.fontIndex = param || 0;
        break;
      case "fs":
        if (param) this.charProps.sizeHalfPt = param;
        break;
      case "b":
        this.charProps.bold = param !== 0;
        break;
      case "i":
        this.charProps.italic = param !== 0;
        break;
      case "ul":
        this.charProps.underline = param === 0 ? false : "single";
        break;
      case "ulnone":
        this.charProps.underline = false;
        break;
      case "strike":
        this.charProps.strike = param !== 0;
        break;
      case "cf":
        this.colorIndex = param == null ? 0 : param;
        break;
      case "highlight":
        this.highlightIndex = param == null ? 0 : param;
        break;
      case "super":
        this.charProps.vertAlign = "superscript";
        break;
      case "sub":
        this.charProps.vertAlign = "subscript";
        break;
      case "nosupersub":
        this.charProps.vertAlign = null;
        break;
      case "caps":
        this.charProps.caps = param !== 0;
        break;
      case "scaps":
        this.charProps.smallCaps = param !== 0;
        break;
      case "pard":
        this.paragraphProps = {};
        break;
      case "ql":
      case "qc":
      case "qr":
      case "qj":
      case "qd":
        this.paragraphProps.align = ALIGNMENT[word];
        break;
      case "li":
        this.paragraphProps.indentLeftTw = param || 0;
        break;
      case "ri":
        this.paragraphProps.indentRightTw = param || 0;
        break;
      case "fi":
        if (param != null) this.paragraphProps.indentFirstLineTw = param;
        break;
      case "sb":
        this.paragraphProps.spaceBeforePt = (param || 0) / 20;
        break;
      case "sa":
        this.paragraphProps.spaceAfterPt = (param || 0) / 20;
        break;
      case "sl":
        if (param) this.paragraphProps.lineHeightPt = param / 20;
        break;
      case "par":
        this.flushParagraph();
        break;
      case "line":
        this.current.runs.push({ type: "break" });
        break;
      case "tab":
        this.current.runs.push({ type: "tab" });
        break;
      case "page":
        this.flushParagraph();
        break;
      case "intbl":
        this.inTable = true;
        break;
      case "trowd":
        this.startRow();
        break;
      case "cellx":
        if (this.row) this.row.grid.push(param || 0);
        break;
      case "cell":
        this.finishCell();
        break;
      case "row":
        this.finishRow();
        break;
      case "u":
        if (param != null) {
          const code = param < 0 ? param + 65536 : param;
          this.addText(String.fromCharCode(code));
        }
        break;
      case "lquote":
        this.addText("\u2018");
        break;
      case "rquote":
        this.addText("\u2019");
        break;
      case "ldblquote":
        this.addText("\u201c");
        break;
      case "rdblquote":
        this.addText("\u201d");
        break;
      case "endash":
        this.addText("\u2013");
        break;
      case "emdash":
        this.addText("\u2014");
        break;
      case "bullet":
        this.addText("\u2022");
        break;
      default:
        break;
    }
  }

  textRun(ch) {
    this.pos++;
    if (ch === "'") {
      const hex = this.text.slice(this.pos, this.pos + 2);
      this.pos += 2;
      const code = parseInt(hex, 16);
      this.addText(String.fromCharCode(isNaN(code) ? 63 : code));
      return;
    }
    this.addText(ch);
  }

  addText(text) {
    this.current.runs.push({ type: "text", text, props: this.resolvedCharProps() });
  }

  resolvedCharProps() {
    const out = {};
    const font = this.tables.fonts.get(this.fontIndex);
    if (font) out.fontFamily = font;
    if (this.charProps.sizeHalfPt) out.sizeHalfPt = this.charProps.sizeHalfPt;
    if (this.charProps.bold) out.bold = true;
    if (this.charProps.italic) out.italic = true;
    if (this.charProps.underline) out.underline = this.charProps.underline;
    if (this.charProps.strike) out.strike = true;
    if (this.charProps.vertAlign) out.vertAlign = this.charProps.vertAlign;
    if (this.charProps.caps) out.caps = true;
    if (this.charProps.smallCaps) out.smallCaps = true;
    const color = this.tables.colors[this.colorIndex];
    if (color) out.color = "rgb(" + color.r + ", " + color.g + ", " + color.b + ")";
    const highlight = this.tables.colors[this.highlightIndex];
    if (highlight) out.highlight = "rgb(" + highlight.r + ", " + highlight.g + ", " + highlight.b + ")";
    return Object.keys(out).length ? out : null;
  }

  flushParagraph() {
    const runs = this.current.runs;
    const text = runs.map((run) => (run.type === "text" ? run.text : "")).join("");
    const hasVisible = text.trim().length > 0 || runs.some((run) => run.type === "break" || run.type === "tab");
    if (hasVisible) {
      const paragraph = {
        type: "p",
        props: Object.keys(this.paragraphProps).length ? Object.assign({}, this.paragraphProps) : null,
        style: null,
        runs,
        markRunProps: null,
        numbering: null,
        bookmarks: [],
        sectionBreak: null,
      };
      if (this.row) {
        this.currentCell().blocks.push(paragraph);
      } else {
        this.blocks.push(paragraph);
      }
    }
    this.current = newParagraph();
  }

  // ---------- tables ----------

  startRow() {
    if (this.row) this.finishRow();
    this.inTable = true;
    this.row = { props: null, cells: [], grid: [] };
  }

  currentCell() {
    if (!this.row.cells.length) {
      this.row.cells.push({ props: {}, blocks: [], gridSpan: 1, vMerge: null });
    }
    return this.row.cells[this.row.cells.length - 1];
  }

  finishCell() {
    if (!this.row) this.startRow();
    this.flushParagraph();
    if (this.row.cells.length === 0) this.currentCell();
    this.row.cells.push({ props: {}, blocks: [], gridSpan: 1, vMerge: null });
  }

  finishRow() {
    if (!this.row) return;
    this.flushParagraph();
    const cells = this.row.cells.filter((cell) => cell.blocks.length > 0);
    if (cells.length) {
      this.blocks.push({
        type: "table",
        props: { widthPct: 100 },
        grid: this.row.grid.length ? this.row.grid.map(() => 0) : [],
        rows: [{ props: null, cells: cells.length ? cells : this.row.cells }],
        styleId: null,
      });
    }
    this.row = null;
  }
}

function newParagraph() {
  return { runs: [], props: null };
}

module.exports = { readRtf };
