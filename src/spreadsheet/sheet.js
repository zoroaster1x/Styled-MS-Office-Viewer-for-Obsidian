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

const {
  attr, attrInt, attrBool, childrenOf, firstOf, parseXml,
} = require("../shared/xml");
const { textOfRuns } = require("../shared/addr");
const { parseRange, splitSheetRef } = require("../shared/addr");

// Parses one worksheet XML into a plain model the renderer can use.

function parseSheet(xmlText, ctx) {
  const doc = parseXml(xmlText);
  const root = doc.documentElement;
  const sharedStrings = ctx.sharedStrings || [];
  const rels = ctx.rels || new Map();

  const model = {
    name: ctx.name || "Sheet",
    dims: null,
    defaultRowHeightPt: 15,
    defaultRowHeightPx: 20,
    defaultColWidthChars: null,
    cols: [],
    rows: new Map(),
    cells: new Map(),
    merges: [],
    freeze: null,
    autoFilter: null,
    hyperlinks: new Map(),
    showGridLines: true,
    outlinePr: { summaryBelow: true, summaryRight: true },
    conditional: [],
    tabColor: null,
  };

  const sheetPr = firstOf(root, "sheetPr");
  if (sheetPr) {
    const outlinePr = firstOf(sheetPr, "outlinePr");
    if (outlinePr) {
      model.outlinePr.summaryBelow = attrBool(outlinePr, "summaryBelow", true);
      model.outlinePr.summaryRight = attrBool(outlinePr, "summaryRight", true);
    }
    const tabColor = firstOf(sheetPr, "tabColor");
    if (tabColor) model.tabColor = parseColorAttrs(tabColor);
  }

  const views = firstOf(root, "sheetViews");
  if (views) {
    const view = firstOf(views, "sheetView");
    if (view) {
      model.showGridLines = attrBool(view, "showGridLines", true);
      const pane = firstOf(view, "pane");
      if (pane) {
        const state = attr(pane, "state");
        const xSplit = attrInt(pane, "xSplit", 0);
        const ySplit = attrInt(pane, "ySplit", 0);
        if (state === "frozen" || state === "frozenSplit") {
          model.freeze = { rows: ySplit, cols: xSplit, topLeft: attr(pane, "topLeftCell") };
        }
      }
    }
  }

  const formatPr = firstOf(root, "sheetFormatPr");
  if (formatPr) {
    const h = parseFloat(attr(formatPr, "defaultRowHeight"));
    if (!isNaN(h)) {
      model.defaultRowHeightPt = h;
      model.defaultRowHeightPx = Math.round((h * 96) / 72 * 100) / 100;
    }
    const w = parseFloat(attr(formatPr, "defaultColWidth"));
    if (!isNaN(w)) model.defaultColWidthChars = w;
  }

  const colsEl = firstOf(root, "cols");
  if (colsEl) {
    for (const col of childrenOf(colsEl, "col")) {
      model.cols.push({
        min: attrInt(col, "min", 1),
        max: attrInt(col, "max", 1),
        width: attr(col, "width") != null ? parseFloat(attr(col, "width")) : null,
        hidden: attrBool(col, "hidden", false),
        style: attrInt(col, "style", 0),
        outline: attrInt(col, "outlineLevel", 0),
        collapsed: attrBool(col, "collapsed", false),
      });
    }
  }

  let maxRow = 0;
  let maxCol = 0;

  const sheetData = firstOf(root, "sheetData");
  if (sheetData) {
    for (const rowEl of childrenOf(sheetData, "row")) {
      const r = attrInt(rowEl, "r", 0);
      if (!r) continue;
      const row = {
        r,
        ht: attr(rowEl, "ht") != null ? parseFloat(attr(rowEl, "ht")) : null,
        hidden: attrBool(rowEl, "hidden", false),
        outline: attrInt(rowEl, "outlineLevel", 0),
        collapsed: attrBool(rowEl, "collapsed", false),
        s: attr(rowEl, "s") != null ? attrInt(rowEl, "s", 0) : null,
        customFormat: attrBool(rowEl, "customFormat", false),
      };
      model.rows.set(r, row);
      if (r > maxRow) maxRow = r;

      let nextCol = 0;
      for (const cEl of childrenOf(rowEl, "c")) {
        const ref = attr(cEl, "r");
        let c;
        if (ref) {
          const parsed = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref);
          if (!parsed) { nextCol += 1; continue; }
          c = letterToIndex(parsed[1]);
        } else {
          c = nextCol + 1;
        }
        nextCol = c;
        if (c > maxCol) maxCol = c;
        const cell = parseCell(cEl, r, c, sharedStrings);
        if (cell) model.cells.set(r + ":" + c, cell);
      }
    }
  }

  const mergeCells = firstOf(root, "mergeCells");
  if (mergeCells) {
    for (const merge of childrenOf(mergeCells, "mergeCell")) {
      const range = parseRange(attr(merge, "ref") || "");
      if (range) model.merges.push(range);
    }
  }
  model.merges.sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1);

  const autoFilter = firstOf(root, "autoFilter");
  if (autoFilter) {
    const range = parseRange(attr(autoFilter, "ref") || "");
    if (range) {
      const columns = new Map();
      for (const fc of childrenOf(autoFilter, "filterColumn")) {
        const colId = attrInt(fc, "colId", 0);
        const def = { colId, showButton: attrBool(fc, "showButton", true) };
        const filters = firstOf(fc, "filters");
        if (filters) {
          const values = [];
          for (const f of childrenOf(filters, "filter")) {
            const v = attr(f, "val");
            if (v != null) values.push(v);
          }
          def.values = values;
          def.blank = attrBool(filters, "blank", false);
        }
        const custom = firstOf(fc, "customFilters");
        if (custom) {
          const list = [];
          for (const f of childrenOf(custom, "customFilter")) {
            list.push({ operator: attr(f, "operator") || "equal", val: attr(f, "val") });
          }
          def.custom = { and: attrBool(custom, "and", false), filters: list };
        }
        columns.set(colId, def);
      }
      model.autoFilter = { range, columns };
    }
  }

  const hyperlinks = firstOf(root, "hyperlinks");
  if (hyperlinks) {
    for (const link of childrenOf(hyperlinks, "hyperlink")) {
      const ref = attr(link, "ref") || "";
      const range = parseRange(ref);
      if (!range) continue;
      const rid = link.getAttribute("r:id") || link.getAttribute("id");
      const rel = rid ? rels.get(rid) : null;
      let target = null;
      let mode = null;
      if (rel) {
        target = rel.target;
        mode = rel.mode;
      }
      const location = attr(link, "location");
      const entry = {
        target,
        mode,
        location,
        display: attr(link, "display"),
        tooltip: attr(link, "tooltip"),
      };
      for (let r = range.r1; r <= range.r2; r++) {
        for (let c = range.c1; c <= range.c2; c++) {
          model.hyperlinks.set(r + ":" + c, entry);
        }
      }
    }
  }

  const cfEl = firstOf(root, "conditionalFormatting");
  const cfList = root.children || [];
  for (let i = 0; i < cfList.length; i++) {
    const el = cfList[i];
    if ((el.tagName || "").indexOf("conditionalFormatting") !== -1) {
      parseConditionalFormatting(el, model);
    }
  }

  const dim = firstOf(root, "dimension");
  let dims = dim ? parseRange(attr(dim, "ref") || "") : null;
  if (!dims) {
    let minR = maxRow, minC = maxCol;
    for (const cellKey of model.cells.keys()) {
      const parts = cellKey.split(":");
      const r = parseInt(parts[0], 10);
      const c = parseInt(parts[1], 10);
      if (r < minR) minR = r;
      if (c < minC) minC = c;
    }
    dims = { r1: minR || 1, c1: minC || 1, r2: maxRow || 1, c2: maxCol || 1 };
  }
  for (const merge of model.merges) {
    if (merge.r2 > dims.r2) dims.r2 = merge.r2;
    if (merge.c2 > dims.c2) dims.c2 = merge.c2;
    if (merge.r1 < dims.r1) dims.r1 = merge.r1;
    if (merge.c1 < dims.c1) dims.c1 = merge.c1;
  }
  model.dims = dims;
  return model;
}

function letterToIndex(letters) {
  let col = 0;
  const up = letters.toUpperCase();
  for (let i = 0; i < up.length; i++) col = col * 26 + (up.charCodeAt(i) - 64);
  return col;
}

function parseCell(cEl, r, c, sharedStrings) {
  const t = attr(cEl, "t") || "n";
  const s = attr(cEl, "s") != null ? attrInt(cEl, "s", 0) : null;
  const fEl = firstOf(cEl, "f");
  const f = fEl ? fEl.textContent : null;

  let value = null;
  if (t === "inlineStr") {
    const isEl = firstOf(cEl, "is");
    value = isEl ? textOfRuns(isEl) : "";
  } else {
    const vEl = firstOf(cEl, "v");
    const raw = vEl ? vEl.textContent : null;
    if (raw == null) {
      // Style-only cells (no cached value) still carry borders and fills, so
      // keep them. Excel writes these all over merged regions and styled areas.
      if (f || s != null) return { r, c, s, t, v: null, f };
      return null;
    }
    if (t === "s") {
      const idx = parseInt(raw, 10);
      value = sharedStrings[idx] != null ? sharedStrings[idx] : "";
    } else if (t === "str") {
      value = raw;
    } else if (t === "b") {
      value = raw === "1" || raw === "true";
    } else if (t === "e") {
      value = raw;
    } else {
      const n = Number(raw);
      value = isNaN(n) ? raw : n;
    }
  }
  if (value === null && !f) return null;
  return { r, c, s, t, v: value, f };
}

function parseColorAttrs(el) {
  const color = {};
  if (attr(el, "rgb")) color.rgb = attr(el, "rgb");
  if (attr(el, "theme") != null) color.theme = attrInt(el, "theme", 0);
  if (attr(el, "indexed") != null) color.indexed = attrInt(el, "indexed", 0);
  if (attr(el, "tint") != null) color.tint = parseFloat(attr(el, "tint"));
  return Object.keys(color).length ? color : null;
}

function parseConditionalFormatting(el, model) {
  const sqref = attr(el, "sqref") || "";
  const ranges = [];
  for (const piece of sqref.split(/\s+/)) {
    const range = parseRange(piece);
    if (range) ranges.push(range);
  }
  const rules = [];
  for (const rule of childrenOf(el, "cfRule")) {
    const type = attr(rule, "type");
    const entry = {
      type,
      operator: attr(rule, "operator"),
      dxfId: attr(rule, "dxfId") != null ? attrInt(rule, "dxfId", 0) : null,
      priority: attrInt(rule, "priority", 0),
      stopIfTrue: attrBool(rule, "stopIfTrue", false),
      text: attr(rule, "text"),
      formulas: [],
    };
    for (const f of childrenOf(rule, "formula")) entry.formulas.push(f.textContent || "");
    const cs = firstOf(rule, "colorScale");
    if (cs) {
      entry.colorScale = {
        values: childrenOf(cs, "cfvo").map((v) => ({
          type: attr(v, "type") || "min",
          val: attr(v, "val"),
        })),
        colors: childrenOf(cs, "color").map((cc) => parseColorAttrs(cc)),
      };
    }
    const db = firstOf(rule, "dataBar");
    if (db) {
      entry.dataBar = {
        color: parseColorAttrs(firstOf(db, "color")),
        values: childrenOf(db, "cfvo").map((v) => ({ type: attr(v, "type") || "min", val: attr(v, "val") })),
      };
    }
    const icons = firstOf(rule, "iconSet");
    if (icons) {
      entry.iconSet = {
        set: attr(icons, "iconSet") || "3TrafficLights1",
        reverse: attrBool(icons, "reverse", false),
        showValue: attr(icons, "showValue") !== "0",
        values: childrenOf(icons, "cfvo").map((v) => ({ type: attr(v, "type") || "percent", val: attr(v, "val") })),
      };
    }
    // top10 covers bottom ten too, and rank/percent change what it means.
    entry.rank = attrInt(rule, "rank", 0) || null;
    entry.percent = attrBool(rule, "percent", false);
    entry.bottom = attrBool(rule, "bottom", false);
    rules.push(entry);
  }
  model.conditional.push({ ranges, rules });
}

module.exports = { parseSheet };
