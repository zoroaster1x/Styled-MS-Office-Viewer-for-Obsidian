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

// Renders one parsed sheet into a DOM grid: styled cells, merged ranges,
// frozen panes, gridlines, wrapped text, filter dropdowns and outline groups.

const { colToLetter, key, naturalCompare, parseRange, splitSheetRef } = require("../shared/addr");
const { formatValue } = require("./numfmt");
const cf = require("./conditional");
const { strongerBorder } = require("./styles");
const { isAutoColor, resolveColor } = require("../shared/color");

const ROW_HEADER_WIDTH = 46;
const COL_HEADER_HEIGHT = 20;
const GUTTER_WIDTH = 14;
const PAD_X = 3;
const PAD_Y = 2;

// Stacking layers. Headers must sit above frozen cells, frozen cells above
// scrolling content, and frozen row numbers above the scrolling numbers.
const Z = {
  cell: 1,
  overflow: 2,
  scrollHead: 3,
  frozen: 4,
  overflowFrozen: 5,
  frozenHead: 7,
  colHead: 7,
  corner: 8,
};

function createRenderer(opts) {
  const container = opts.container;
  const doc = container.ownerDocument || document;
  const sheet = opts.sheet;
  const styles = opts.styles;
  const theme = opts.theme;
  const callbacks = {
    onNavigate: opts.onNavigate || function () {},
    onOpenExternal: opts.onOpenExternal || function (url) { window.open(url, "_blank"); },
    onFilterChange: opts.onFilterChange || function () {},
    onSelect: opts.onSelect || function () {},
    showContextMenu: opts.showContextMenu || null,
  };
  const settings = Object.assign({
    sheetBackground: "white",
    showGridlines: true,
    showHeaders: true,
    maxRows: 2000,
    scale: 1,
  }, opts.settings || {});
  const filterState = opts.filterState || new Map();
  const outlineCollapsed = opts.outlineCollapsed || new Map();
  // Column widths and row heights the reader has dragged. They are kept per
  // sheet by the controller and replayed on every render, because the grid is
  // measured from the file each time.
  const sizeOverrides = {
    cols: new Map(Object.entries((opts.sizeOverrides && opts.sizeOverrides.cols) || {}).map(([k, v]) => [Number(k), Number(v)])),
    rows: new Map(Object.entries((opts.sizeOverrides && opts.sizeOverrides.rows) || {}).map(([k, v]) => [Number(k), Number(v)])),
  };
  const RESIZE_MIN_COL = 24;
  const RESIZE_MAX_COL = 900;
  const RESIZE_MIN_ROW = 12;
  const RESIZE_MAX_ROW = 400;
  const date1904 = Boolean(opts.date1904);

  let rootEl = null;
  let scrollEl = null;
  let gridEl = null;
  let popupEl = null;
  let destroyed = false;
  let itemEls = [];
  let resolveItemEls = new Map();
  let selectedEl = null;
  const selection = { ref: null, r: null, c: null, text: "" };
  // The selected block: an anchor cell and the cell the pointer or keyboard
  // last reached. Both are visible row and column numbers.
  const anchor = { r: null, c: null };
  const focus = { r: null, c: null };
  let dragging = false;
  let textPopupEl = null;
  let G = null;
  // Rows outside the viewport are not built. The grid keeps its full row
  // template, so the scroll height and the row positions do not change; only
  // the cells that are on screen (plus a few rows of overscan) exist in the
  // DOM. This is what keeps scrolling smooth on a long sheet: a 250 row
  // timetable carries about sixty rows of elements instead of all of them.
  let windowStart = 0;
  let windowEnd = 0;
  let scrollPending = false;
  let styleCache = new Map();
  let displayCache = new Map();
  let measureCanvas = null;
  let measureCtx = null;
  const measureCache = new Map();

  // ---------- static per-run geometry ----------

  function styleIndexFor(r, c) {
    const cell = sheet.cells.get(key(r, c));
    if (cell && cell.s != null) return cell.s;
    const row = sheet.rows.get(r);
    if (row && row.customFormat && row.s != null) return row.s;
    const colDef = colDefFor(c);
    if (colDef && colDef.style) return colDef.style;
    return 0;
  }

  let colDefMap = null;
  function colDefFor(c) {
    if (!colDefMap) {
      colDefMap = new Array((G ? G.maxCol : sheet.dims.c2) + 2).fill(null);
      for (const def of sheet.cols) {
        for (let c2 = def.min; c2 <= def.max; c2++) {
          colDefMap[c2] = def;
          if (c2 >= colDefMap.length) break;
        }
      }
    }
    return colDefMap[c] || null;
  }

  function resolvedStyle(r, c) {
    const idx = styleIndexFor(r, c);
    const cacheKey = r + ":" + c;
    if (styleCache.has(cacheKey)) return styleCache.get(cacheKey);
    const st = styles.resolveXf(idx);
    styleCache.set(cacheKey, st);
    return st;
  }

  function cellValue(r, c) {
    const cell = sheet.cells.get(key(r, c));
    return cell ? cell.v : null;
  }

  function cellValueType(r, c) {
    const cell = sheet.cells.get(key(r, c));
    return cell ? cell.t : "n";
  }

  // Returns { text, color, extra } for the cell, applying number format and
  // conditional formatting on top of the raw value.
  function displayOf(r, c) {
    const cacheKey = r + ":" + c;
    if (displayCache.has(cacheKey)) return displayCache.get(cacheKey);
    const cell = sheet.cells.get(key(r, c));
    let out = { text: "", color: null, extra: null };
    if (cell) {
      const st = styles.resolveXf(styleIndexFor(r, c));
      const t = cell.t;
      if (t === "s" || t === "str" || t === "is" || t === "inlineStr" || t === "e") {
        out = { text: cell.v == null ? "" : String(cell.v), color: null, extra: null };
      } else {
        const formatted = formatValue(cell.v, st.fmt, { date1904 });
        out = { text: formatted.text, color: formatted.color || null, extra: null };
      }
      if (typeof cell.v === "number") {
        const cf = conditionalOverride(r, c, cell.v, out.text);
        if (cf) {
          out = Object.assign({}, out, { extra: cf });
          if (cf.fill) out.extraFill = cf.fill;
          if (cf.fontColor) out.color = cf.fontColor;
        }
      }
    }
    displayCache.set(cacheKey, out);
    return out;
  }

  // Conditional formatting lives in its own module (src/spreadsheet/conditional.js),
  // because the arithmetic is the interesting part. This wrapper supplies the
  // cell reader and the colour and style resolvers it needs.
  function conditionalOverride(r, c, numValue, textValue) {
    return cf.overrideForCell(sheet, r, c, numValue, textValue, {
      readValue: (row, col) => cellValue(row, col),
      resolveColour: (color, fallback) => resolveColor(color, theme, fallback || null),
      dxfStyle: (id) => styles.dxfStyle(id),
    });
  }

  function mixColors(a, b, t) {
    const pa = parseCss(a);
    const pb = parseCss(b);
    if (!pa || !pb) return a;
    const mix = (x, y) => Math.round(x + (y - x) * t);
    return "rgb(" + mix(pa[0], pb[0]) + ", " + mix(pa[1], pb[1]) + ", " + mix(pa[2], pb[2]) + ")";
  }

  function parseCss(color) {
    if (!color) return null;
    if (color[0] === "#") {
      const hex = color.length === 4
        ? color[1] + color[1] + color[2] + color[2] + color[3] + color[3]
        : color.slice(1);
      return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
    }
    const m = /rgb\((\d+),\s*(\d+),\s*(\d+)\)/.exec(color);
    if (m) return [Number(m[1]), Number(m[2]), Number(m[3])];
    return null;
  }

  // Automatic text colour on a coloured fill: black on light fills, white on
  // dark ones, so multicolour sheets stay readable in either theme.
  function autoTextColorForFill(fillCss) {
    const rgb = parseCss(fillCss);
    if (!rgb) return null;
    const lum = (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) / 255;
    return lum > 0.55 ? "#000000" : "#FFFFFF";
  }

  // ---------- visibility ----------

  function prepare() {
    styleCache = new Map();
    displayCache = new Map();
    colDefMap = null;

    const dims = sheet.dims;
    const maxRow = Math.max(1, dims.r2);
    const maxCol = Math.max(1, dims.c2);
    // Excel always shows the grid from row 1 and column A, even when the used
    // range starts further in.
    const usedR1 = 1;
    const usedC1 = 1;

    const hiddenRows0 = new Set();
    const hiddenCols0 = new Set();
    for (const [r, row] of sheet.rows) if (row.hidden) hiddenRows0.add(r);
    for (const def of sheet.cols) {
      if (!def.hidden) continue;
      for (let c = def.min; c <= def.max; c++) hiddenCols0.add(c);
    }
    // Also treat zero width columns as hidden.
    for (const def of sheet.cols) {
      if (def.width != null && def.width <= 0) {
        for (let c = def.min; c <= def.max; c++) hiddenCols0.add(c);
      }
    }

    const hiddenByFilter = new Set();
    const filterRange = sheet.autoFilter ? sheet.autoFilter.range : null;
    if (filterRange) {
      for (const [colId, sel] of filterState) {
        if (sel == null) continue;
        const col = filterRange.c1 + colId;
        const custom = sel && sel.custom ? sel.custom : null;
        const set = sel instanceof Set ? sel : null;
        for (let r = filterRange.r1 + 1; r <= filterRange.r2 && r <= maxRow; r++) {
          const text = displayTextRaw(r, col);
          let matched;
          if (custom) matched = matchesCustom(custom, text, r, col);
          else if (set) matched = set.has(text);
          else matched = true;
          if (!matched) hiddenByFilter.add(r);
        }
      }
    }

    const rowGroups = buildOutlineGroups("row");
    const colGroups = buildOutlineGroups("col");
    const hiddenByOutline = new Set();
    const hiddenColsByOutline = new Set();
    for (const group of rowGroups) {
      if (!outlineCollapsed.get(group.key)) continue;
      for (let r = group.start; r <= group.end; r++) hiddenByOutline.add(r);
    }
    for (const group of colGroups) {
      if (!outlineCollapsed.get(group.key)) continue;
      for (let c = group.start; c <= group.end; c++) hiddenColsByOutline.add(c);
    }

    const hiddenRows = new Set([...hiddenRows0, ...hiddenByFilter, ...hiddenByOutline]);
    const hiddenCols = new Set([...hiddenCols0, ...hiddenColsByOutline]);

    const maxRowsSetting = settings.maxRows > 0 ? settings.maxRows : 0;
    const renderRowEnd = maxRowsSetting ? Math.min(maxRow, maxRowsSetting) : maxRow;

    const visRows = [];
    const visRowIndex = new Map();
    for (let r = usedR1; r <= renderRowEnd; r++) {
      if (hiddenRows.has(r)) continue;
      visRowIndex.set(r, visRows.length);
      visRows.push(r);
    }
    const visCols = [];
    const visColIndex = new Map();
    for (let c = usedC1; c <= maxCol; c++) {
      if (hiddenCols.has(c)) continue;
      visColIndex.set(c, visCols.length);
      visCols.push(c);
    }

    const colWidths = new Array(maxCol + 2).fill(0);
    const defaultColWidth = sheet.defaultColWidthChars != null
      ? Math.round(sheet.defaultColWidthChars * 7) + 5
      : 64;
    for (let c = 1; c <= maxCol; c++) {
      const def = colDefFor(c);
      colWidths[c] = def && def.width != null ? Math.round(def.width * 7) + 5 : defaultColWidth;
    }

    const rowHeights = new Array(maxRow + 2).fill(0);
    for (let r = 1; r <= maxRow; r++) {
      const row = sheet.rows.get(r);
      rowHeights[r] = row && row.ht != null && row.ht > 0
        ? Math.round((row.ht * 96) / 72 * 100) / 100
        : sheet.defaultRowHeightPx;
    }

    // A reader-set minimum, off by default, then anything dragged by hand.
    const minCol = Number(settings.minColumnWidth) > 0 ? Number(settings.minColumnWidth) : 0;
    const minRow = Number(settings.minRowHeight) > 0 ? Number(settings.minRowHeight) : 0;
    if (minCol) {
      for (let c = 1; c <= maxCol; c++) colWidths[c] = Math.max(colWidths[c], minCol);
    }
    if (minRow) {
      for (let r = 1; r <= maxRow; r++) rowHeights[r] = Math.max(rowHeights[r], minRow);
    }
    for (const [c, px] of sizeOverrides.cols) {
      if (c >= 1 && c <= maxCol && px > 0) colWidths[c] = px;
    }
    for (const [r, px] of sizeOverrides.rows) {
      if (r >= 1 && r <= maxRow && px > 0) rowHeights[r] = px;
    }

    const colLeftData = new Array(visCols.length + 1).fill(0);
    for (let j = 0; j < visCols.length; j++) {
      colLeftData[j + 1] = colLeftData[j] + colWidths[visCols[j]];
    }
    const rowTopData = new Array(visRows.length + 1).fill(0);
    for (let i = 0; i < visRows.length; i++) {
      rowTopData[i + 1] = rowTopData[i] + rowHeights[visRows[i]];
    }

    const coverMap = new Map();
    for (const merge of sheet.merges) {
      for (let r = merge.r1; r <= Math.min(merge.r2, maxRow); r++) {
        for (let c = merge.c1; c <= Math.min(merge.c2, maxCol); c++) {
          coverMap.set(key(r, c), merge);
        }
      }
    }

    G = {
      maxRow, maxCol, usedR1, usedC1, usedR2: maxRow, usedC2: maxCol,
      hiddenRows, hiddenCols, hiddenRows0, hiddenByFilter, hiddenByOutline,
      hiddenCols0, hiddenColsByOutline,
      visRows, visRowIndex, visCols, visColIndex,
      colWidths, rowHeights, colLeftData, rowTopData,
      coverMap, filterRange, rowGroups, colGroups,
      renderRowEnd, rowCapped: renderRowEnd < maxRow,
      hasRowOutlines: rowGroups.length > 0,
      hasColOutlines: colGroups.length > 0,
    };
  }

  function buildOutlineGroups(kind) {
    const groups = [];
    if (kind === "row") {
      let maxLevel = 0;
      for (const row of sheet.rows.values()) if (row.outline > maxLevel) maxLevel = row.outline;
      for (let level = 1; level <= maxLevel; level++) {
        let start = null;
        for (let r = 1; r <= G_maxRowHint(); r++) {
          const row = sheet.rows.get(r);
          const lvl = row ? row.outline : 0;
          if (lvl >= level) {
            if (start == null) start = r;
          } else if (start != null) {
            groups.push({ kind, level, start, end: r - 1, key: "row:" + level + ":" + start });
            start = null;
          }
        }
        if (start != null) {
          groups.push({ kind, level, start, end: G_maxRowHint(), key: "row:" + level + ":" + start });
        }
      }
    } else {
      const levels = new Map();
      for (const def of sheet.cols) {
        if (def.outline > 0) levels.set(def.min, def.outline);
      }
      let maxLevel = 0;
      for (const lvl of levels.values()) if (lvl > maxLevel) maxLevel = lvl;
      for (let level = 1; level <= maxLevel; level++) {
        let start = null;
        const maxCol = G_maxColHint();
        for (let c = 1; c <= maxCol; c++) {
          const def = colDefFor(c);
          const lvl = def ? def.outline : 0;
          if (lvl >= level) {
            if (start == null) start = c;
          } else if (start != null) {
            groups.push({ kind, level, start, end: c - 1, key: "col:" + level + ":" + start });
            start = null;
          }
        }
        if (start != null) groups.push({ kind, level, start, end: maxCol, key: "col:" + level + ":" + start });
      }
    }
    return groups;
  }

  // Small indirection so buildOutlineGroups can run before G exists.
  function G_maxRowHint() {
    const dims = sheet.dims;
    let max = dims.r2 || 1;
    for (const r of sheet.rows.keys()) if (r > max) max = r;
    return max;
  }
  function G_maxColHint() {
    const dims = sheet.dims;
    let max = dims.c2 || 1;
    for (const merge of sheet.merges) if (merge.c2 > max) max = merge.c2;
    return max;
  }

  function displayTextRaw(r, c) {
    return displayOf(r, c).text;
  }

  function matchesCustom(custom, text, r, c) {
    const num = Number(text);
    let result = null;
    for (const filter of custom.filters) {
      let hit;
      const val = filter.val != null ? filter.val : "";
      const numVal = Number(val);
      switch (filter.operator) {
        case "greaterThan": hit = !isNaN(num) && num > numVal; break;
        case "greaterThanOrEqual": hit = !isNaN(num) && num >= numVal; break;
        case "lessThan": hit = !isNaN(num) && num < numVal; break;
        case "lessThanOrEqual": hit = !isNaN(num) && num <= numVal; break;
        case "equal": hit = text === val; break;
        case "notEqual": hit = text !== val; break;
        default: hit = true;
      }
      if (result == null) result = hit;
      else result = custom.and ? (result && hit) : (result || hit);
    }
    if (result == null) return true;
    if (custom.and) return result;
    return result;
  }

  // ---------- text measuring ----------

  function fontString(style) {
    return (style.font.italic ? "italic " : "") + (style.font.bold ? "bold " : "")
      + style.font.sizePx + "px " + style.font.family;
  }

  function measureText(text, style) {
    const font = fontString(style);
    const cacheKey = font + "\u0000" + text;
    if (measureCache.has(cacheKey)) return measureCache.get(cacheKey);
    try {
      if (!measureCtx) {
        measureCanvas = doc.createElement("canvas");
        measureCtx = measureCanvas.getContext ? measureCanvas.getContext("2d") : null;
        if (!measureCtx) return null;
      }
      measureCtx.font = font;
      const w = measureCtx.measureText(text).width;
      measureCache.set(cacheKey, w);
      return w;
    } catch (err) {
      return null;
    }
  }

  // ---------- borders ----------

  function rawBorder(r, c, side) {
    const st = styles.resolveXf(styleIndexFor(r, c));
    return st.borders[side];
  }

  function cellEdgeBorder(r, c, side) {
    const merge = G.coverMap.get(key(r, c));
    if (!merge) return rawBorder(r, c, side);
    let best = null;
    if (side === "left" && c === merge.c1) {
      for (let rr = merge.r1; rr <= merge.r2; rr++) best = strongerBorder(best, rawBorder(rr, merge.c1, "left"));
      return best;
    }
    if (side === "right" && c === merge.c2) {
      for (let rr = merge.r1; rr <= merge.r2; rr++) best = strongerBorder(best, rawBorder(rr, merge.c2, "right"));
      return best;
    }
    if (side === "top" && r === merge.r1) {
      for (let cc = merge.c1; cc <= merge.c2; cc++) best = strongerBorder(best, rawBorder(merge.r1, cc, "top"));
      return best;
    }
    if (side === "bottom" && r === merge.r2) {
      for (let cc = merge.c1; cc <= merge.c2; cc++) best = strongerBorder(best, rawBorder(merge.r2, cc, "bottom"));
      return best;
    }
    return null;
  }

  function gridlineSpec() {
    return { width: 1, css: "solid", color: "var(--xlsx-gridline)", weight: 0, style: "thin" };
  }

  // Borders come from the cell itself and from its neighbours *as drawn*. A row
  // hidden by a filter, by the reader or by an outline group contributes
  // nothing: Excel never draws the edge of a row that is not displayed, which
  // is what used to leave ghost lines under a filtered block.
  function visibleNeighbourRow(row, step) {
    let r = row + step;
    while (r >= 1 && r <= G.maxRow) {
      if (G.visRowIndex.has(r)) return r;
      r += step;
    }
    return null;
  }

  function visibleNeighbourCol(col, step) {
    let c = col + step;
    while (c >= 1 && c <= G.maxCol) {
      if (G.visColIndex.has(c)) return c;
      c += step;
    }
    return null;
  }

  function pickEdge(r1, c1, r2, c2, side) {
    let best = null;
    if (side === "left" || side === "right") {
      for (let r = r1; r <= r2; r++) {
        if (!G.visRowIndex.has(r)) continue;
        best = strongerBorder(best, cellEdgeBorder(r, c1, side));
      }
      const neighbour = side === "left" ? visibleNeighbourCol(c1, -1) : visibleNeighbourCol(c2, 1);
      if (neighbour != null) {
        const other = side === "left" ? "right" : "left";
        for (let r = r1; r <= r2; r++) {
          if (!G.visRowIndex.has(r)) continue;
          best = strongerBorder(best, cellEdgeBorder(r, neighbour, other));
        }
      }
    } else {
      for (let c = c1; c <= c2; c++) {
        if (!G.visColIndex.has(c)) continue;
        best = strongerBorder(best, cellEdgeBorder(r1, c, side));
      }
      const neighbour = side === "top" ? visibleNeighbourRow(r1, -1) : visibleNeighbourRow(r2, 1);
      if (neighbour != null) {
        const other = side === "top" ? "bottom" : "top";
        for (let c = c1; c <= c2; c++) {
          if (!G.visColIndex.has(c)) continue;
          best = strongerBorder(best, cellEdgeBorder(neighbour, c, other));
        }
      }
    }
    if (!best && sheet.showGridLines !== false && settings.showGridlines) best = gridlineSpec();
    return best;
  }

  function applyBorder(el, which, spec, force) {
    if (!spec) return;
    if (!force && which !== "left" && which !== "top") return;
    el.style["border" + which.charAt(0).toUpperCase() + which.slice(1)] = spec.width + "px " + spec.css + " " + spec.color;
  }

  // ---------- build ----------

  function render(preserveScroll) {
    if (destroyed) return;
    const prev = preserveScroll && scrollEl ? { top: scrollEl.scrollTop, left: scrollEl.scrollLeft } : null;
    if (!preserveScroll) closePopup();
    prepare();

    const showHeaders = settings.showHeaders;
    const gutter = G.hasRowOutlines || G.hasColOutlines ? GUTTER_WIDTH : 0;
    const rowHeaderW = showHeaders ? ROW_HEADER_WIDTH : 0;
    const colHeaderH = showHeaders ? COL_HEADER_HEIGHT : 0;
    const colBase = (gutter ? 1 : 0) + (showHeaders ? 1 : 0);

    const templateCols = [];
    if (gutter) templateCols.push(GUTTER_WIDTH + "px");
    if (showHeaders) templateCols.push(ROW_HEADER_WIDTH + "px");
    for (const c of G.visCols) templateCols.push(G.colWidths[c] + "px");
    templateCols.push("minmax(40px, 1fr)");

    const templateRows = [];
    if (showHeaders) templateRows.push(COL_HEADER_HEIGHT + "px");
    for (const r of G.visRows) templateRows.push(G.rowHeights[r] + "px");

    const frag = doc.createDocumentFragment();
    itemEls = [];
    resolveItemEls = new Map();

    const freezeRows = sheet.freeze ? sheet.freeze.rows : 0;
    const freezeCols = sheet.freeze ? sheet.freeze.cols : 0;
    const overscan = Number(settings.overscanRows) >= 0 ? Number(settings.overscanRows) : 8;
    const frozenCount = Math.min(freezeRows, G.visRows.length);
    G.freezeRows = freezeRows;
    G.freezeCols = freezeCols;
    // Which rows to build for a given scroll offset: the frozen rows always,
    // then the rows crossing the viewport with overscan on both sides.
    G.computeWindow = (top, viewHeight) => {
      const height = viewHeight || (scrollEl && scrollEl.clientHeight) || 900;
      let firstVisible = frozenCount;
      while (firstVisible < G.visRows.length && G.rowTopData[firstVisible + 1] <= top) firstVisible++;
      const start = Math.max(frozenCount, firstVisible - overscan);
      let end = firstVisible;
      const bottom = top + height;
      while (end < G.visRows.length && G.rowTopData[end] < bottom) end++;
      end = Math.min(G.visRows.length, end + overscan);
      if (end <= start) end = Math.min(G.visRows.length, start + overscan + 8);
      return { start, end };
    };
    const initialWindow = G.computeWindow(scrollEl ? scrollEl.scrollTop : 0, 0);
    windowStart = initialWindow.start;
    windowEnd = initialWindow.end;

    const rowTopOf = (r) => {
      const i = G.visRowIndex.get(r);
      return i == null ? null : colHeaderH + G.rowTopData[i];
    };
    const colLeftOf = (c) => {
      const j = G.visColIndex.get(c);
      return j == null ? null : rowHeaderW + gutter + G.colLeftData[j];
    };

    // Column headers and row headers.
    if (showHeaders) {
      const corner = doc.createElement("div");
      corner.className = "xlsx-corner";
      corner.style.gridColumn = "1 / span " + Math.max(colBase, 1);
      corner.style.gridRow = "1";
      corner.style.position = "sticky";
      corner.style.top = "0";
      corner.style.left = "0";
      corner.style.zIndex = String(Z.corner);
      frag.appendChild(corner);

      for (let j = 0; j < G.visCols.length; j++) {
        const c = G.visCols[j];
        const head = doc.createElement("div");
        head.className = "xlsx-col-head";
        head.textContent = colToLetter(c);
        head.style.gridColumn = String(colBase + j + 1);
        head.style.gridRow = "1";
        head.style.position = "sticky";
        head.style.top = "0";
        head.style.zIndex = String(Z.colHead);
        if (c <= freezeCols) head.style.left = (rowHeaderW + gutter + G.colLeftData[j]) + "px";
        const handle = doc.createElement("div");
        handle.className = "xlsx-col-resize";
        handle.title = "Drag to set the column width, double click to fit the text";
        handle.addEventListener("mousedown", (ev) => startColumnResize(ev, c, j));
        handle.addEventListener("dblclick", (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          autoFitColumn(c);
        });
        head.appendChild(handle);
        frag.appendChild(head);
      }
    }

    const headerRows = [];
    for (let i = 0; i < Math.min(freezeRows, G.visRows.length); i++) headerRows.push(i);
    for (let i = windowStart; i < windowEnd; i++) {
      if (i >= Math.min(freezeRows, G.visRows.length)) headerRows.push(i);
    }
    for (const i of headerRows) {
      const r = G.visRows[i];
      const rowHead = doc.createElement("div");
      if (showHeaders) {
        rowHead.className = "xlsx-row-head";
        rowHead.textContent = String(r);
        rowHead.style.gridColumn = String(gutter + 1);
        rowHead.style.gridRow = String(i + 2);
        rowHead.style.position = "sticky";
        rowHead.style.left = "0";
        rowHead.style.zIndex = String(r <= freezeRows ? Z.frozenHead : Z.scrollHead);
        if (r <= freezeRows) rowHead.style.top = (colHeaderH + G.rowTopData[i]) + "px";
        const rowHandle = doc.createElement("div");
        rowHandle.className = "xlsx-row-resize";
        rowHandle.title = "Drag to set the row height, double click to fit the text";
        rowHandle.addEventListener("mousedown", (ev) => startRowResize(ev, r, i));
        rowHandle.addEventListener("dblclick", (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          autoFitRow(r);
        });
        rowHead.appendChild(rowHandle);
        frag.appendChild(rowHead);
      }
    }

    // Outline gutter buttons.
    if (G.hasRowOutlines) {
      for (const group of G.rowGroups) {
        const controlRow = sheet.outlinePr.summaryBelow
          ? Math.min(group.end + 1, G.maxRow)
          : Math.max(group.start - 1, 1);
        let target = controlRow;
        if (!G.visRowIndex.has(target)) {
          target = null;
          for (let r = group.start; r <= group.end; r++) {
            if (G.visRowIndex.has(r)) { target = r; break; }
          }
        }
        if (target == null) continue;
        const btn = doc.createElement("div");
        btn.className = "xlsx-outline-btn";
        btn.textContent = outlineCollapsed.get(group.key) ? "+" : "-";
        btn.style.gridColumn = "1";
        btn.style.gridRow = String(G.visRowIndex.get(target) + 2);
        btn.style.position = "sticky";
        btn.style.left = "0";
        btn.addEventListener("click", (ev) => {
          ev.stopPropagation();
          toggleOutline(group.key);
        });
        frag.appendChild(btn);
      }
    }
    if (G.hasColOutlines) {
      for (const group of G.colGroups) {
        const controlCol = sheet.outlinePr.summaryRight ? Math.min(group.end + 1, G.maxCol) : Math.max(group.start - 1, 1);
        let target = controlCol;
        if (!G.visColIndex.has(target)) {
          target = null;
          for (let c = group.start; c <= group.end; c++) {
            if (G.visColIndex.has(c)) { target = c; break; }
          }
        }
        if (target == null) continue;
        const btn = doc.createElement("div");
        btn.className = "xlsx-outline-btn";
        btn.textContent = outlineCollapsed.get(group.key) ? "+" : "-";
        btn.style.gridColumn = String(colBase + G.visColIndex.get(target) + 1);
        btn.style.gridRow = "1";
        btn.style.position = "sticky";
        btn.style.top = "0";
        btn.addEventListener("click", (ev) => {
          ev.stopPropagation();
          toggleOutline(group.key);
        });
        frag.appendChild(btn);
      }
    }

    // Frozen rows first, then the window, so the frozen band paints last and
    // covers whatever scrolls beneath it.
    const rowOrder = [];
    for (let i = 0; i < Math.min(freezeRows, G.visRows.length); i++) rowOrder.push(i);
    for (let i = windowStart; i < windowEnd; i++) {
      if (i >= Math.min(freezeRows, G.visRows.length)) rowOrder.push(i);
    }

    // One opaque backdrop behind the frozen band, so nothing shows through it
    // while the sheet scrolls underneath.
    if (freezeRows > 0) {
      const backdrop = doc.createElement("div");
      backdrop.className = "xlsx-freeze-backdrop";
      backdrop.style.gridRow = "1 / span " + (Math.min(freezeRows, G.visRows.length) + 1);
      backdrop.style.gridColumn = "1 / -1";
      backdrop.style.position = "sticky";
      backdrop.style.top = "0";
      backdrop.style.left = "0";
      backdrop.style.zIndex = String(Z.frozen - 1);
      frag.appendChild(backdrop);
    }

    // Data cells.
    for (const i of rowOrder) {
      const r = G.visRows[i];
      for (let j = 0; j < G.visCols.length; j++) {
        const c = G.visCols[j];
        const merge = G.coverMap.get(key(r, c));
        if (merge && (merge.r1 !== r || merge.c1 !== c)) continue;
        const r2 = merge ? Math.min(merge.r2, G.maxRow) : r;
        const c2 = merge ? Math.min(merge.c2, G.maxCol) : c;
        buildItem(frag, r, c, r2, c2, Boolean(merge), i, j, colBase, gutter, rowHeaderW, colHeaderH, freezeRows, freezeCols, rowTopOf, colLeftOf);
      }
    }

    // Rebuild root.
    rootEl = doc.createElement("div");
    rootEl.className = "xlsx-sheet" + (settings.sheetBackground === "white" ? " xlsx-white" : "");
    scrollEl = doc.createElement("div");
    scrollEl.className = "xlsx-sheet-scroll";
    gridEl = doc.createElement("div");
    gridEl.className = "xlsx-sheet-grid";
    gridEl.style.position = "relative";
    gridEl.style.gridTemplateColumns = templateCols.join(" ");
    gridEl.style.gridTemplateRows = templateRows.join(" ");
    if (settings.scale && settings.scale !== 1) gridEl.style.zoom = String(settings.scale);
    gridEl.appendChild(frag);
    scrollEl.appendChild(gridEl);
    rootEl.appendChild(scrollEl);

    if (G.rowCapped) {
      const note = doc.createElement("div");
      note.className = "xlsx-cap-note";
      note.textContent = "Showing rows 1 to " + G.renderRowEnd + " of " + G.maxRow
        + ". Change the row limit in the plugin settings to see more.";
      rootEl.appendChild(note);
    }

    container.textContent = "";
    container.appendChild(rootEl);

    scrollEl.addEventListener("scroll", onScroll, { passive: true });
    scrollEl.addEventListener("click", onGridClick);
    gridEl.addEventListener("contextmenu", onGridContext);
    gridEl.addEventListener("mousedown", onGridMouseDown);
    gridEl.addEventListener("dblclick", onGridDoubleClick);
    if (!scrollEl.hasAttribute("tabindex")) scrollEl.setAttribute("tabindex", "0");
    scrollEl.addEventListener("keydown", onGridKeyDown);
    restoreSelection();

    if (prev) {
      scrollEl.scrollTop = prev.top;
      scrollEl.scrollLeft = prev.left;
    }
  }

  // Scrolling only rebuilds the grid when the visible row range changes, and
  // never more than once a frame.
  function onScroll() {
    if (destroyed || scrollPending) return;
    scrollPending = true;
    const raf = typeof requestAnimationFrame === "function" ? requestAnimationFrame : (cb) => setTimeout(cb, 16);
    raf(() => {
      scrollPending = false;
      if (destroyed || !scrollEl || !G || !G.computeWindow) return;
      const next = G.computeWindow(scrollEl.scrollTop, scrollEl.clientHeight);
      if (next.start !== windowStart || next.end !== windowEnd) render(true);
    });
  }

  function buildItem(frag, r, c, r2, c2, isMerge, i, j, colBase, gutter, rowHeaderW, colHeaderH, freezeRows, freezeCols, rowTopOf, colLeftOf) {
    const el = doc.createElement("div");
    // The cell's own address, so pointer handling and tests can find it.
    el.dataset.r = String(r);
    el.dataset.c = String(c);
    el.className = "xlsx-cell";
    let spanRows = 0;
    for (let rr = r; rr <= r2; rr++) if (G.visRowIndex.has(rr)) spanRows++;
    let spanCols = 0;
    for (let cc = c; cc <= c2; cc++) if (G.visColIndex.has(cc)) spanCols++;
    if (!spanRows || !spanCols) return;
    el.style.gridRow = (i + 2) + " / span " + spanRows;
    el.style.gridColumn = (colBase + j + 1) + " / span " + spanCols;
    if (isMerge) el.classList.add("xlsx-merged");

    const style = resolvedStyle(r, c);
    const disp = displayOf(r, c);
    const extra = disp.extra || null;
    const value = cellValue(r, c);
    const hasValue = value != null && disp.text !== "";

    let fillCss = styles.resolveFill(style.fill);
    if (extra && extra.fill) fillCss = extra.fill;
    const isGradient = style.fill && style.fill.type === "gradient" && !(extra && extra.fill);

    const stickyTop = r <= freezeRows ? rowTopOf(r) : null;
    const stickyLeft = c <= freezeCols ? colLeftOf(c) : null;
    const isSticky = stickyTop != null || stickyLeft != null;
    if (stickyTop != null) el.style.top = stickyTop + "px";
    if (stickyLeft != null) el.style.left = stickyLeft + "px";
    // Plain cells stay static and unstacked on purpose: thousands of stacking
    // contexts make scrolling stutter.

    if (fillCss) el.style.backgroundColor = fillCss;
    else if (isGradient) el.style.backgroundImage = styles.gradientCss(style.fill);
    else if (isSticky) el.style.backgroundColor = "var(--xlsx-sheet-bg)";

    // Text layout.
    el.style.display = "flex";
    el.style.flexDirection = "column";
    const vAlign = style.align.v === "center" ? "center" : style.align.v === "top" ? "flex-start" : "flex-end";
    el.style.justifyContent = vAlign;
    let hAlign = style.align.h;
    if (!hAlign) {
      const t = cellValueType(r, c);
      hAlign = (typeof value === "number") ? "right" : (typeof value === "boolean" || t === "e") ? "center" : "left";
    }
    el.style.alignItems = hAlign === "center" ? "center" : hAlign === "right" ? "flex-end" : "flex-start";
    const indent = style.align.indent || 0;
    el.style.padding = PAD_Y + "px " + (PAD_X + indent * 9) + "px " + PAD_Y + "px " + (hAlign === "right" && indent ? (PAD_X + indent * 9) : PAD_X) + "px";

    // Borders. Left and top are drawn by every item; right and bottom only at
    // the edge of the used range, so shared lines are drawn once. The frozen
    // boundary keeps its line while the sheet scrolls.
    const freezeBottom = freezeRows > 0 && r2 === freezeRows;
    const freezeRight = freezeCols > 0 && c2 === freezeCols;
    const afterFreezeRow = freezeRows > 0 && r === freezeRows + 1;
    const afterFreezeCol = freezeCols > 0 && c === freezeCols + 1;
    if (!afterFreezeCol) applyBorder(el, "left", pickEdge(r, c, r2, c2, "left"), true);
    if (!afterFreezeRow) applyBorder(el, "top", pickEdge(r, c, r2, c2, "top"), true);
    if (freezeRight) applyBorder(el, "right", pickEdge(r, c, r2, c2, "right"), true);
    if (freezeBottom) applyBorder(el, "bottom", pickEdge(r, c, r2, c2, "bottom"), true);
    if (!freezeRight && c2 >= G.usedC2) applyBorder(el, "right", pickEdge(r, c, r2, c2, "right"), true);
    if (!freezeBottom && r2 >= G.usedR2) applyBorder(el, "bottom", pickEdge(r, c, r2, c2, "bottom"), true);

    // Conditional data bar.
    if (extra && extra.bar) {
      const bar = doc.createElement("div");
      bar.className = "xlsx-databar";
      bar.style.width = Math.round(extra.bar.t * 100) + "%";
      bar.style.backgroundColor = extra.bar.color;
      el.appendChild(bar);
    }

    // Conditional icon set: a glyph before the value.
    if (extra && extra.icon && extra.icon.glyph) {
      const icon = doc.createElement("div");
      icon.className = "xlsx-cf-icon";
      icon.textContent = extra.icon.glyph;
      icon.style.color = extra.icon.color;
      el.appendChild(icon);
    }

    let overflows = false;
    let hasFilterButton = false;
    if (hasValue) {
      const span = doc.createElement("div");
      span.className = "xlsx-cell-text";
      span.textContent = disp.text;
      span.style.fontFamily = style.font.family;
      span.style.fontSize = style.font.sizePx + "px";
      if (style.font.bold || (extra && extra.bold)) span.style.fontWeight = "700";
      if (style.font.italic || (extra && extra.italic)) span.style.fontStyle = "italic";
      if (style.font.underline) {
        span.style.textDecoration = "underline";
        if (style.font.underline === "double") span.style.textDecorationStyle = "double";
      }
      if (style.font.strike) span.style.textDecoration = "line-through";

      let color = disp.color || styles.resolveFontColor(style.font.color, null);
      const autoFont = !disp.color && isAutoColor(style.font.color) && !extra;
      if (autoFont && fillCss) color = autoTextColorForFill(fillCss);
      else if (autoFont && settings.sheetBackground === "theme") color = null;
      if (color) span.style.color = color;

      const wrap = style.align.wrap;
      if (wrap || isMerge) {
        span.classList.add("xlsx-wrap");
        span.style.whiteSpace = "pre-wrap";
        span.style.overflowWrap = "break-word";
        span.style.maxWidth = "100%";
        // A wrapped cell wraps at the cell edge and each line follows the
        // horizontal alignment, so centred and right-aligned headers line up
        // the way Excel draws them.
        span.style.width = "100%";
        span.style.textAlign = hAlign;
        el.style.alignItems = "stretch";
      } else {
        span.style.whiteSpace = "pre";
        // Text overflows into empty neighbours the way Excel does, and is
        // clipped at the first non empty cell. max-content width is needed
        // because flex items would otherwise shrink to the cell width.
        el.style.overflow = "visible";
        overflows = true;
        span.style.width = "max-content";
        const limit = findBlocker(r, c);
        if (limit != null) {
          const own = G.colLeftData[G.visColIndex.get(c)];
          span.style.maxWidth = Math.max(0, limit - own - PAD_X) + "px";
          span.style.overflow = "hidden";
        } else {
          span.style.overflow = "visible";
        }
      }

      // Numbers that do not fit show hashes, like Excel.
      if (typeof value === "number" && !wrap && !isMerge) {
        const avail = G.colWidths[c] - PAD_X * 2 - indent * 9;
        const est = measureText(disp.text, style);
        const hashW = measureText("#", style);
        if (est != null && hashW != null && hashW > 0 && est > avail) {
          span.textContent = "#".repeat(Math.max(1, Math.floor(avail / hashW)));
        }
      }

      el.appendChild(span);
    }
    // Only content cells need clipping; empty cells skip the extra paint pass.
    if (!overflows && hasValue) el.style.overflow = "hidden";

    // Positioning is only applied where it is needed (sticky freeze, text
    // overflow, data bars, filter buttons). Everything else stays static so
    // scrolling does not maintain thousands of stacking contexts.
    if (searchQuery) {
      if (isSearchHit(r, c)) el.classList.add("xlsx-search-hit");
      const current = searchHits[searchIndex];
      if (current && current.r === r && current.c === c) el.classList.add("xlsx-search-current");
    }

    if (isSticky) {
      el.style.position = "sticky";
      el.style.zIndex = String(overflows ? Z.overflowFrozen : Z.frozen);
    } else if (overflows) {
      el.style.position = "relative";
      el.style.zIndex = String(Z.overflow);
    } else if ((extra && extra.bar) || hasFilterButton) {
      el.style.position = "relative";
    }

    // Hyperlinks.
    const link = sheet.hyperlinks.get(key(r, c));
    if (link) {
      el.classList.add("xlsx-link");
      el.dataset.linkKey = key(r, c);
      // The theme link colour is only safe on cells without a fill.
      if (!fillCss && !disp.color) el.classList.add("xlsx-link-default");
    }

    // Filter buttons on the header row of the autofilter range.
    if (G.filterRange && r === G.filterRange.r1 && c >= G.filterRange.c1 && c <= G.filterRange.c2) {
      const colId = c - G.filterRange.c1;
      const def = sheet.autoFilter.columns.get(colId);
      if (!def || def.showButton !== false) {
        hasFilterButton = true;
        const btn = doc.createElement("div");
        btn.className = "xlsx-filter-btn";
        btn.title = "Filter column " + colToLetter(c);
        btn.textContent = "v";
        btn.addEventListener("click", (ev) => {
          ev.stopPropagation();
          openFilterPopup(colId, btn);
        });
        el.appendChild(btn);
        if (filterState.has(colId) && filterState.get(colId) != null) el.classList.add("xlsx-filtered");
      }
    }

    frag.appendChild(el);
    itemEls.push({ el, r, c, text: disp.text });
    resolveItemEls.set(key(r, c), el);
  }

  function findBlocker(r, c) {
    for (let cc = c + 1; cc <= G.usedC2; cc++) {
      if (!G.visColIndex.has(cc)) continue;
      const merge = G.coverMap.get(key(r, cc));
      if (merge) return G.colLeftData[G.visColIndex.get(cc)];
      if (displayOf(r, cc).text !== "") return G.colLeftData[G.visColIndex.get(cc)];
    }
    return null;
  }

  function toggleOutline(groupKey) {
    outlineCollapsed.set(groupKey, !outlineCollapsed.get(groupKey));
    render(true);
  }

  // ---------- selection and clipboard ----------

  function isVisibleRow(r) {
    return G.visRowIndex.has(r);
  }

  function isVisibleCol(c) {
    return G.visColIndex.has(c);
  }

  // A merged cell selects as its whole rectangle, the way Excel does.
  function expandToMerge(r, c) {
    const merge = G.coverMap.get(key(r, c));
    if (!merge) return { r1: r, c1: c, r2: r, c2: c };
    return { r1: merge.r1, c1: merge.c1, r2: merge.r2, c2: merge.c2 };
  }

  function normalisedRange() {
    if (anchor.r == null || focus.r == null) return null;
    const a = expandToMerge(anchor.r, anchor.c);
    const b = expandToMerge(focus.r, focus.c);
    return {
      r1: Math.min(a.r1, b.r1),
      c1: Math.min(a.c1, b.c1),
      r2: Math.max(a.r2, b.r2),
      c2: Math.max(a.c2, b.c2),
    };
  }

  function clearSelectionClasses() {
    for (const el of itemEls) {
      el.el.classList.remove("xlsx-selected");
      el.el.classList.remove("xlsx-in-range");
    }
    selectedEl = null;
  }

  function paintSelection() {
    const range = normalisedRange();
    if (!range) return;
    for (const item of itemEls) {
      const inside = item.r >= range.r1 && item.r <= range.r2 && item.c >= range.c1 && item.c <= range.c2;
      item.el.classList.toggle("xlsx-in-range", inside);
      item.el.classList.toggle("xlsx-selected", item.r === anchor.r && item.c === anchor.c);
    }
    const anchorItem = itemEls.find((item) => item.r === anchor.r && item.c === anchor.c);
    selectedEl = anchorItem ? anchorItem.el : null;
    const active = itemEls.find((item) => item.r === focus.r && item.c === focus.c) || anchorItem;
    if (active) {
      selection.ref = colToLetter(active.c) + active.r;
      selection.r = active.r;
      selection.c = active.c;
      selection.text = active.text;
    }
    callbacks.onSelect({
      ref: selection.ref,
      r: selection.r,
      c: selection.c,
      text: selection.text,
      range: rangeLabel(range),
      count: countRangeCells(range),
    });
  }

  function rangeLabel(range) {
    if (!range) return "";
    const a = colToLetter(range.c1) + range.r1;
    const b = colToLetter(range.c2) + range.r2;
    return a === b ? a : a + ":" + b;
  }

  function countRangeCells(range) {
    let n = 0;
    for (const item of itemEls) {
      if (item.r >= range.r1 && item.r <= range.r2 && item.c >= range.c1 && item.c <= range.c2) n++;
    }
    return n;
  }

  function setSelection(r, c, extend) {
    if (!extend || anchor.r == null) {
      anchor.r = r;
      anchor.c = c;
    }
    focus.r = r;
    focus.c = c;
    paintSelection();
  }

  function moveFocus(dr, dc, extend) {
    if (focus.r == null) {
      setSelection(G.visRows[0], G.visCols[0], false);
      return;
    }
    const rowIndex = G.visRowIndex.get(focus.r);
    const colIndex = G.visColIndex.get(focus.c);
    const nextRow = G.visRows[Math.max(0, Math.min(G.visRows.length - 1, rowIndex + dr))];
    const nextCol = G.visCols[Math.max(0, Math.min(G.visCols.length - 1, colIndex + dc))];
    setSelection(nextRow, nextCol, extend);
    ensureVisible(nextCol, nextRow);
  }

  // Keeps the focused cell inside the scroller.
  function ensureVisible(col, row) {
    if (!scrollEl) return;
    const left = G.colLeftData[G.visColIndex.get(col)];
    const width = G.colWidths[col];
    const top = G.rowTopData[G.visRowIndex.get(row)];
    const height = G.rowHeights[row];
    if (left < scrollEl.scrollLeft) scrollEl.scrollLeft = left;
    else if (left + width > scrollEl.scrollLeft + scrollEl.clientWidth) {
      scrollEl.scrollLeft = left + width - scrollEl.clientWidth;
    }
    if (top < scrollEl.scrollTop) scrollEl.scrollTop = top;
    else if (top + height > scrollEl.scrollTop + scrollEl.clientHeight) {
      scrollEl.scrollTop = top + height - scrollEl.clientHeight;
    }
  }

  function selectAllVisble() {
    if (!G.visRows.length || !G.visCols.length) return;
    anchor.r = G.visRows[0];
    anchor.c = G.visCols[0];
    focus.r = G.visRows[G.visRows.length - 1];
    focus.c = G.visCols[G.visCols.length - 1];
    paintSelection();
  }

  // The displayed text of one cell, merges and blanks included.
  function textForCopy(r, c) {
    const merge = G.coverMap.get(key(r, c));
    if (merge && (merge.r1 !== r || merge.c1 !== c)) return "";
    return displayOf(r, c).text;
  }

  // A block of cells as tab separated rows, the way Excel puts a range on the
  // clipboard. Hidden and filtered rows and columns are skipped, exactly as
  // Excel skips them.
  function selectionTsv(range) {
    const lines = [];
    for (let r = range ? range.r1 : 0; r <= (range ? range.r2 : 0); r++) {
      if (!isVisibleRow(r)) continue;
      const cells = [];
      for (let c = range.c1; c <= range.c2; c++) {
        if (!isVisibleCol(c)) continue;
        cells.push(tsvField(textForCopy(r, c)));
      }
      lines.push(cells.join("\t"));
    }
    return lines.join("\n");
  }

  // A field that holds a tab, a newline or a quote is wrapped in quotes with
  // its quotes doubled, which is what Excel writes. Without
  // this a wrapped cell would split its row in the receiving application.
  function tsvField(text) {
    const value = text == null ? "" : String(text);
    if (/[\t\n\r"]/.test(value)) return '"' + value.replace(/"/g, '""') + '"';
    return value;
  }

  // The block as an HTML table. Rich copy carries the styling so a paste into
  // Word or PowerPoint keeps the fills and fonts; the plain copy leaves every
  // style off.
  function selectionHtml(range, withStyles) {
    const rows = [];
    for (let r = range.r1; r <= range.r2; r++) {
      if (!isVisibleRow(r)) continue;
      const cells = [];
      for (let c = range.c1; c <= range.c2; c++) {
        if (!isVisibleCol(c)) continue;
        const text = textForCopy(r, c);
        if (!withStyles) {
          cells.push("<td>" + escapeHtml(text) + "</td>");
          continue;
        }
        const style = styles.resolveXf(styleIndexFor(r, c));
        const css = [];
        const fill = styles.resolveFill(style.fill);
        if (fill) css.push("background-color:" + fill);
        css.push("font-family:" + style.font.family);
        css.push("font-size:" + style.font.sizePt + "pt");
        if (style.font.bold) css.push("font-weight:700");
        if (style.font.italic) css.push("font-style:italic");
        if (style.font.underline) css.push("text-decoration:underline");
        if (style.font.strike) css.push("text-decoration:line-through");
        const color = styles.resolveFontColor(style.font.color, null);
        if (color) css.push("color:" + color);
        const align = style.align.h || (typeof sheet.cells.get(key(r, c)) === "object" ? null : null);
        if (align) css.push("text-align:" + align);
        const merge = G.coverMap.get(key(r, c));
        let span = "";
        if (merge && merge.r1 === r && merge.c1 === c) {
          const colspan = merge.c2 - merge.c1 + 1;
          const rowspan = merge.r2 - merge.r1 + 1;
          if (colspan > 1) span += ' colspan="' + colspan + '"';
          if (rowspan > 1) span += ' rowspan="' + rowspan + '"';
        }
        cells.push("<td" + span + ' style="' + css.join(";") + '">' + escapeHtml(text) + "</td>");
      }
      rows.push("<tr>" + cells.join("") + "</tr>");
    }
    return '<table border="1" cellspacing="0" cellpadding="2">' + rows.join("") + "</table>";
  }

  function escapeHtml(text) {
    return String(text == null ? "" : text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  // Copy modes: "rich" carries styling and structure, "plain" is values only.
  async function copySelection(mode) {
    const range = normalisedRange();
    if (!range) {
      const text = selection.text || "";
      if (text) await writeClipboard(text, null, null);
      return false;
    }
    if (range.r1 === range.r2 && range.c1 === range.c2) {
      const text = textForCopy(range.r1, range.c1);
      await writeClipboard(text, null, null);
      return true;
    }
    const tsv = selectionTsv(range);
    if (mode === "plain") {
      await writeClipboard(tsv, null, null);
      return true;
    }
    const html = selectionHtml(range, true);
    await writeClipboard(tsv, html, "text/html");
    return true;
  }

  // The Electron clipboard takes several flavours at once when the async API
  // is available; otherwise the text flavour is still Excel compatible.
  async function writeClipboard(plain, html, htmlType) {
    try {
      if (html && typeof ClipboardItem === "function" && navigator.clipboard && navigator.clipboard.write) {
        const item = new ClipboardItem({
          "text/plain": new Blob([plain], { type: "text/plain" }),
          [htmlType || "text/html"]: new Blob([html], { type: htmlType || "text/html" }),
        });
        await navigator.clipboard.write([item]);
        return true;
      }
    } catch (err) {
      // Fall through to the plain text write.
    }
    try {
      await navigator.clipboard.writeText(plain);
      return true;
    } catch (err) {
      return false;
    }
  }

  function onGridKeyDown(ev) {
    if (ev.defaultPrevented) return;
    const tag = (ev.target && ev.target.tagName || "").toLowerCase();
    if (tag === "input" || tag === "textarea") return;
    const meta = ev.ctrlKey || ev.metaKey;
    if (meta && (ev.key === "c" || ev.key === "C")) {
      ev.preventDefault();
      // Shift means values only, without colours or fonts.
      copySelection(ev.shiftKey ? "plain" : "rich");
      return;
    }
    if (meta && (ev.key === "a" || ev.key === "A")) {
      ev.preventDefault();
      selectAllVisble();
      return;
    }
    switch (ev.key) {
      case "ArrowUp":
        ev.preventDefault();
        moveFocus(-1, 0, ev.shiftKey);
        break;
      case "ArrowDown":
        ev.preventDefault();
        moveFocus(1, 0, ev.shiftKey);
        break;
      case "ArrowLeft":
        ev.preventDefault();
        moveFocus(0, -1, ev.shiftKey);
        break;
      case "ArrowRight":
        ev.preventDefault();
        moveFocus(0, 1, ev.shiftKey);
        break;
      case "PageUp":
        ev.preventDefault();
        moveFocus(-20, 0, ev.shiftKey);
        break;
      case "PageDown":
        ev.preventDefault();
        moveFocus(20, 0, ev.shiftKey);
        break;
      case "Home":
        ev.preventDefault();
        moveFocus(0, -G.visCols.length, ev.shiftKey);
        break;
      case "End":
        ev.preventDefault();
        moveFocus(0, G.visCols.length, ev.shiftKey);
        break;
      case "Escape":
        clearSelectionClasses();
        anchor.r = anchor.c = focus.r = focus.c = null;
        break;
      default:
        break;
    }
  }

  // ---------- interactions ----------

  function onGridClick(ev) {
    const target = ev.target;
    const linkEl = closestWithClass(target, "xlsx-link");
    if (linkEl) {
      const entry = sheet.hyperlinks.get(linkEl.dataset.linkKey);
      if (entry) {
        ev.preventDefault();
        if (entry.location) {
          const info = splitSheetRef(entry.location);
          callbacks.onNavigate(info.sheet, info.ref);
        } else if (entry.target) {
          callbacks.onOpenExternal(entry.target);
        }
      }
    }
    const cellEl = closestWithClass(target, "xlsx-cell");
    if (!cellEl) return;
    const item = itemForElement(cellEl);
    if (!item) return;
    setSelection(item.r, item.c, ev.shiftKey);
    if (scrollEl) scrollEl.focus();
  }

  // Dragging across cells extends the selection, the way a spreadsheet does.
  // The move and release listeners sit on the document, so the drag survives
  // the pointer leaving the cell it started on.
  function onGridMouseDown(ev) {
    if (ev.button !== undefined && ev.button !== 0) return;
    const cellEl = closestWithClass(ev.target, "xlsx-cell");
    if (!cellEl) return;
    const item = itemForElement(cellEl);
    if (!item) return;
    if (ev.shiftKey) return;
    // Stop the browser starting a text selection across the grid.
    if (ev.preventDefault) ev.preventDefault();
    dragging = true;
    setSelection(item.r, item.c, false);
    doc.addEventListener("mousemove", onGridMouseMove, true);
    doc.addEventListener("mouseup", onGridMouseUp, true);
  }

  function onGridMouseMove(ev) {
    if (!dragging) return;
    const cellEl = closestWithClass(ev.target, "xlsx-cell");
    if (!cellEl) return;
    const item = itemForElement(cellEl);
    if (!item) return;
    if (item.r === focus.r && item.c === focus.c) return;
    setSelection(item.r, item.c, true);
  }

  function onGridMouseUp() {
    if (!dragging) return;
    dragging = false;
    doc.removeEventListener("mousemove", onGridMouseMove, true);
    doc.removeEventListener("mouseup", onGridMouseUp, true);
  }

  // ---------- column and row sizing ----------

  // Live resize: while the pointer is down the grid template is rewritten in
  // place, which is cheap, and a full render happens once on release.
  function startColumnResize(ev, col, visIndex) {
    if (ev.button !== undefined && ev.button !== 0) return;
    ev.preventDefault();
    ev.stopPropagation();
    const startX = ev.clientX != null ? ev.clientX : 0;
    const startWidth = G.colWidths[col];
    let width = startWidth;
    const onMove = (moveEvent) => {
      const x = moveEvent.clientX != null ? moveEvent.clientX : startX;
      width = Math.min(RESIZE_MAX_COL, Math.max(RESIZE_MIN_COL, Math.round(startWidth + (x - startX))));
      G.colWidths[col] = width;
      sizeOverrides.cols.set(col, width);
      applyColumnTemplate();
    };
    const onUp = () => {
      doc.removeEventListener("mousemove", onMove, true);
      doc.removeEventListener("mouseup", onUp, true);
      render(true);
      notifySizesChanged();
    };
    doc.addEventListener("mousemove", onMove, true);
    doc.addEventListener("mouseup", onUp, true);
    void visIndex;
  }

  function startRowResize(ev, row, visIndex) {
    if (ev.button !== undefined && ev.button !== 0) return;
    ev.preventDefault();
    ev.stopPropagation();
    const startY = ev.clientY != null ? ev.clientY : 0;
    const startHeight = G.rowHeights[row];
    let height = startHeight;
    const onMove = (moveEvent) => {
      const y = moveEvent.clientY != null ? moveEvent.clientY : startY;
      height = Math.min(RESIZE_MAX_ROW, Math.max(RESIZE_MIN_ROW, Math.round(startHeight + (y - startY))));
      G.rowHeights[row] = height;
      sizeOverrides.rows.set(row, height);
      applyRowTemplate();
    };
    const onUp = () => {
      doc.removeEventListener("mousemove", onMove, true);
      doc.removeEventListener("mouseup", onUp, true);
      render(true);
      notifySizesChanged();
    };
    doc.addEventListener("mousemove", onMove, true);
    doc.addEventListener("mouseup", onUp, true);
    void visIndex;
  }

  function columnTemplate() {
    const templateCols = [];
    if (G.hasRowOutlines || G.hasColOutlines) templateCols.push(GUTTER_WIDTH + "px");
    if (settings.showHeaders !== false) templateCols.push(ROW_HEADER_WIDTH + "px");
    for (const c of G.visCols) templateCols.push(G.colWidths[c] + "px");
    templateCols.push("minmax(40px, 1fr)");
    return templateCols.join(" ");
  }

  function rowTemplate() {
    const templateRows = [];
    if (settings.showHeaders !== false) templateRows.push(COL_HEADER_HEIGHT + "px");
    for (const r of G.visRows) templateRows.push(G.rowHeights[r] + "px");
    return templateRows.join(" ");
  }

  function applyColumnTemplate() {
    if (gridEl) {
      gridEl.style.gridTemplateColumns = columnTemplate();
      G.colLeftData = recomputeColLeft();
    }
  }

  function applyRowTemplate() {
    if (gridEl) {
      gridEl.style.gridTemplateRows = rowTemplate();
      G.rowTopData = recomputeRowTop();
      updateRowHeaderOffsets();
    }
  }

  function recomputeColLeft() {
    const data = new Array(G.visCols.length + 1).fill(0);
    for (let j = 0; j < G.visCols.length; j++) data[j + 1] = data[j] + G.colWidths[G.visCols[j]];
    return data;
  }

  function recomputeRowTop() {
    const data = new Array(G.visRows.length + 1).fill(0);
    for (let i = 0; i < G.visRows.length; i++) data[i + 1] = data[i] + G.rowHeights[G.visRows[i]];
    return data;
  }

  // Frozen row numbers are placed by hand, so they follow a live resize.
  function updateRowHeaderOffsets() {
    const colHeaderH = settings.showHeaders !== false ? COL_HEADER_HEIGHT : 0;
    for (const el of itemEls) {
      if (el.el.className.indexOf("xlsx-row-head") === -1) continue;
      const r = el.r;
      if (!G.visRowIndex.has(r)) continue;
      if (G.freezeRows && r <= G.freezeRows) {
        el.el.style.top = (colHeaderH + G.rowTopData[G.visRowIndex.get(r)]) + "px";
      }
    }
  }

  // Auto-fit measures the widest value in the column, the way double clicking a
  // column edge does in a spreadsheet. Only the rendered window is measured,
  // which is a good enough sample and keeps it instant.
  function autoFitColumn(col) {
    let widest = 0;
    for (const item of itemEls) {
      if (item.c !== col) continue;
      const style = resolvedStyle(item.r, item.c);
      const text = displayOf(item.r, item.c).text;
      if (!text) continue;
      const width = measureText(text, style);
      if (width != null && width > widest) widest = width;
    }
    const header = colToLetter(col);
    const headerWidth = measureText(header, styles.resolveXf(0));
    if (headerWidth != null && headerWidth > widest) widest = headerWidth;
    const target = Math.min(RESIZE_MAX_COL, Math.max(RESIZE_MIN_COL, Math.ceil(widest + PAD_X * 2 + 6)));
    sizeOverrides.cols.set(col, target);
    if (G) G.colWidths[col] = target;
    render(true);
    notifySizesChanged();
  }

  function autoFitRow(row) {
    let tallest = 0;
    for (const item of itemEls) {
      if (item.r !== row) continue;
      const style = resolvedStyle(item.r, item.c);
      const text = displayOf(item.r, item.c).text;
      if (!text) continue;
      const width = G.colWidths[item.c] - PAD_X * 2;
      const full = measureText(text, style);
      if (full == null || width <= 0) continue;
      const lines = Math.max(1, Math.ceil(full / width));
      const height = lines * (style.font.sizePx * 1.25) + PAD_Y * 2 + 4;
      if (height > tallest) tallest = height;
    }
    const target = Math.min(RESIZE_MAX_ROW, Math.max(RESIZE_MIN_ROW, Math.ceil(tallest)));
    sizeOverrides.rows.set(row, target);
    if (G) G.rowHeights[row] = target;
    render(true);
    notifySizesChanged();
  }

  function notifySizesChanged() {
    if (callbacks.onSizesChanged) callbacks.onSizesChanged(getSizes());
  }

  // The sizes to remember for this sheet.
  function getSizes() {
    const cols = {};
    const rows = {};
    for (const [c, px] of sizeOverrides.cols) cols[c] = px;
    for (const [r, px] of sizeOverrides.rows) rows[r] = px;
    return { cols, rows };
  }

  function resetSizes() {
    sizeOverrides.cols.clear();
    sizeOverrides.rows.clear();
    render(true);
    notifySizesChanged();
  }

  // A long value is clipped by the cell next to it, exactly as Excel clips it.
  // Double clicking shows the whole thing, read only.
  function onGridDoubleClick(ev) {
    const cellEl = closestWithClass(ev.target, "xlsx-cell");
    if (!cellEl) return;
    const item = itemForElement(cellEl);
    if (!item) return;
    ev.preventDefault();
    showCellText(item.r, item.c, cellEl);
  }

  function showCellText(r, c, cellEl) {
    closeTextPopup();
    const range = expandToMerge(r, c);
    const text = displayOf(range.r1, range.c1).text;
    const ref = colToLetter(range.c1) + range.r1 +
      (range.r2 > range.r1 || range.c2 > range.c1
        ? ":" + colToLetter(range.c2) + range.r2
        : "");
    const popup = doc.createElement("div");
    popup.className = "xlsx-cell-popup";
    const head = doc.createElement("div");
    head.className = "xlsx-cell-popup-head";
    head.textContent = ref;
    popup.appendChild(head);
    const body = doc.createElement("div");
    body.className = "xlsx-cell-popup-body";
    body.textContent = text || "(empty)";
    popup.appendChild(body);
    const foot = doc.createElement("div");
    foot.className = "xlsx-cell-popup-foot";
    const copyBtn = doc.createElement("button");
    copyBtn.className = "xlsx-filter-btn-plain";
    copyBtn.textContent = "Copy";
    copyBtn.addEventListener("click", () => writeClipboard(text, null, null));
    const closeBtn = doc.createElement("button");
    closeBtn.className = "xlsx-filter-btn-plain";
    closeBtn.textContent = "Close";
    closeBtn.addEventListener("click", () => closeTextPopup());
    foot.appendChild(copyBtn);
    foot.appendChild(closeBtn);
    popup.appendChild(foot);
    doc.body.appendChild(popup);
    const rect = cellEl.getBoundingClientRect ? cellEl.getBoundingClientRect() : { left: 40, bottom: 40, top: 40 };
    const width = Math.min(520, Math.max(240, text.length * 7 + 40));
    let left = rect.left;
    if (left + width > doc.documentElement.clientWidth - 8) {
      left = Math.max(8, doc.documentElement.clientWidth - width - 8);
    }
    popup.style.width = width + "px";
    popup.style.left = left + "px";
    const below = (rect.bottom || rect.top || 0) + 4;
    popup.style.top = Math.min(below, Math.max(8, doc.documentElement.clientHeight - 200)) + "px";
    textPopupEl = popup;
    const onDown = (downEvent) => {
      if (popup.contains(downEvent.target)) return;
      closeTextPopup();
    };
    const onKey = (keyEvent) => {
      if (keyEvent.key === "Escape") closeTextPopup();
    };
    popup._onDown = onDown;
    popup._onKey = onKey;
    doc.addEventListener("mousedown", onDown, true);
    doc.addEventListener("keydown", onKey, true);
  }

  function closeTextPopup() {
    if (textPopupEl) {
      if (textPopupEl._onDown) doc.removeEventListener("mousedown", textPopupEl._onDown, true);
      if (textPopupEl._onKey) doc.removeEventListener("keydown", textPopupEl._onKey, true);
      if (textPopupEl.parentNode) textPopupEl.parentNode.removeChild(textPopupEl);
      textPopupEl = null;
    }
  }

  function getSelection() {
    return {
      ref: selection.ref,
      r: selection.r,
      c: selection.c,
      text: selection.text,
      range: rangeLabel(normalisedRange()),
      count: normalisedRange() ? countRangeCells(normalisedRange()) : 0,
    };
  }

  function restoreSelection() {
    if (!selection.ref || !resolveItemEls) return;
    if (anchor.r == null) {
      anchor.r = selection.r;
      anchor.c = selection.c;
      focus.r = selection.r;
      focus.c = selection.c;
    }
    const el = resolveItemEls.get(key(selection.r, selection.c));
    if (el) {
      selectedEl = el;
      el.classList.add("xlsx-selected");
    }
    // A rendered grid only holds the rows near the viewport, so the whole
    // selected block is repainted from the selection state after each render.
    if (anchor.r != null && focus.r != null) paintRangeClasses();
  }

  function paintRangeClasses() {
    const range = normalisedRange();
    if (!range) return;
    for (const item of itemEls) {
      const inside = item.r >= range.r1 && item.r <= range.r2 && item.c >= range.c1 && item.c <= range.c2;
      item.el.classList.toggle("xlsx-in-range", inside);
    }
  }

  function onGridContext(ev) {
    const target = ev.target;
    const cellEl = closestWithClass(target, "xlsx-cell");
    if (!cellEl || !callbacks.showContextMenu) return;
    ev.preventDefault();
    const item = itemForElement(cellEl);
    if (!item) return;
    const rangeLabelText = rangeLabel(normalisedRange()) || (colToLetter(item.c) + item.r);
    callbacks.showContextMenu(ev, [
      { title: "Copy " + rangeLabelText + " with formatting (Ctrl+C)", icon: "copy", action: () => copySelection("rich") },
      { title: "Copy " + rangeLabelText + " as plain values (Ctrl+Shift+C)", icon: "clipboard-type", action: () => copySelection("plain") },
      { title: "Copy cell reference", icon: "hash", action: () => navigator.clipboard.writeText(rangeLabelText) },
      { title: "Select all visible cells", icon: "square-dashed", action: () => selectAllVisble() },
    ]);
  }

  function itemForElement(el) {
    for (const item of itemEls) if (item.el === el) return item;
    return null;
  }

  function closestWithClass(node, className) {
    let cur = node;
    while (cur && cur.classList) {
      if (cur.classList.contains(className)) return cur;
      cur = cur.parentNode;
    }
    return null;
  }

  // Scrolls so the cell sits below the frozen band and right of the frozen
  // columns. The offsets come from the grid's own tables rather than
  // offsetTop, which a CSS grid with sticky cells does not report usefully.
  function scrollToCell(ref) {
    const range = parseRange(ref || "A1");
    if (!range || !G) return;
    const i = G.visRowIndex.get(range.r1);
    const j = G.visColIndex.get(range.c1);
    if (i == null || j == null) return;
    const frozenHeight = G.freezeRows > 0 ? G.rowTopData[Math.min(G.freezeRows, G.visRows.length)] : 0;
    const top = Math.max(0, G.rowTopData[i] - frozenHeight - 40);
    const left = Math.max(0, G.colLeftData[j] - 40);
    doScroll(top, left);
    const el = resolveItemEls.get(key(range.r1, range.c1));
    if (el) {
      el.classList.add("xlsx-flash");
      setTimeout(() => el.classList.remove("xlsx-flash"), 900);
    }
  }

  // scrollTo is missing in some DOM shims (and older webviews), so fall back
  // to setting scrollTop and scrollLeft.
  function doScroll(top, left) {
    if (!scrollEl) return;
    // Clamp to the scrollable range only when the element reports one; a test
    // DOM has no layout and reports zero.
    const maxTop = scrollEl.scrollHeight > 0 ? Math.max(0, scrollEl.scrollHeight - scrollEl.clientHeight) : Infinity;
    const maxLeft = scrollEl.scrollWidth > 0 ? Math.max(0, scrollEl.scrollWidth - scrollEl.clientWidth) : Infinity;
    const t = Math.min(Math.max(0, top), maxTop);
    const l = Math.min(Math.max(0, left), maxLeft);
    if (typeof scrollEl.scrollTo === "function") {
      scrollEl.scrollTo({ top: t, left: l });
    } else {
      scrollEl.scrollTop = t;
      scrollEl.scrollLeft = l;
    }
  }

  // ---------- filters ----------

  function filterColumns() {
    if (!G || !G.filterRange) return [];
    const out = [];
    const range = G.filterRange;
    for (let c = range.c1; c <= range.c2; c++) {
      const colId = c - range.c1;
      const def = sheet.autoFilter.columns.get(colId);
      if (def && def.showButton === false) continue;
      const values = new Map();
      let blanks = false;
      for (let r = range.r1 + 1; r <= range.r2; r++) {
        const text = displayTextRaw(r, c);
        if (text === "") blanks = true;
        values.set(text, (values.get(text) || 0) + 1);
      }
      const headerText = displayTextRaw(range.r1, c);
      out.push({
        colId,
        col: c,
        letter: colToLetter(c),
        headerText,
        values: Array.from(values.entries())
          .map(([value, count]) => ({ value, count }))
          .sort((a, b) => naturalCompare(a.value, b.value)),
        blanks,
        showButton: !(def && def.showButton === false),
      });
    }
    return out;
  }

  function openFilterPopup(colId, anchor) {
    closePopup();
    const range = G.filterRange;
    if (!range) return;
    const col = range.c1 + colId;
    const values = [];
    let blanks = false;
    for (let r = range.r1 + 1; r <= range.r2; r++) {
      const text = displayTextRaw(r, col);
      if (text === "") { blanks = true; continue; }
      if (values.indexOf(text) === -1) values.push(text);
    }
    values.sort(naturalCompare);
    const allValues = blanks ? values.concat([""]) : values;

    let current = filterState.get(colId);
    if (!(current instanceof Set)) current = current && current.custom ? null : current;

    const popup = doc.createElement("div");
    popup.className = "xlsx-filter-popup";

    const head = doc.createElement("div");
    head.className = "xlsx-filter-popup-head";
    const headerText = displayTextRaw(range.r1, col);
    head.textContent = "Column " + colToLetter(col) + (headerText ? " (" + headerText + ")" : "");
    popup.appendChild(head);

    const search = doc.createElement("input");
    search.type = "text";
    search.placeholder = "Search values";
    search.className = "xlsx-filter-search";
    popup.appendChild(search);

    const allRow = doc.createElement("label");
    allRow.className = "xlsx-filter-item xlsx-filter-all";
    const allBox = doc.createElement("input");
    allBox.type = "checkbox";
    allBox.checked = current == null;
    allRow.appendChild(allBox);
    allRow.appendChild(doc.createTextNode("Select all"));
    popup.appendChild(allRow);

    const list = doc.createElement("div");
    list.className = "xlsx-filter-list";
    popup.appendChild(list);

    const rows = new Map();
    const makeItem = (value, label) => {
      const row = doc.createElement("label");
      row.className = "xlsx-filter-item";
      const box = doc.createElement("input");
      box.type = "checkbox";
      box.checked = current == null || current.has(value);
      box.addEventListener("change", () => {
        const state = filterState.get(colId);
        let set;
        if (state instanceof Set) set = new Set(state);
        else if (state && state.custom) set = new Set(allValues);
        else set = new Set(allValues);
        if (box.checked) set.add(value);
        else set.delete(value);
        let next = set;
        if (set.size === allValues.length) next = null;
        filterState.set(colId, next);
        allBox.checked = next == null;
        callbacks.onFilterChange();
        render(true);
      });
      row.appendChild(box);
      row.appendChild(doc.createTextNode(label));
      rows.set(value, { row, box, label });
      list.appendChild(row);
      return row;
    };
    for (const value of values) makeItem(value, value);
    if (blanks) makeItem("", "(Blanks)");

    allBox.addEventListener("change", () => {
      if (allBox.checked) filterState.set(colId, null);
      else filterState.set(colId, new Set());
      for (const entry of rows.values()) entry.box.checked = allBox.checked;
      callbacks.onFilterChange();
      render(true);
    });

    search.addEventListener("input", () => {
      const q = search.value.toLowerCase();
      for (const entry of rows.values()) {
        const show = entry.label.toLowerCase().indexOf(q) !== -1;
        entry.row.style.display = show ? "" : "none";
      }
    });

    const footer = doc.createElement("div");
    footer.className = "xlsx-filter-popup-foot";

    const clearBtn = doc.createElement("button");
    clearBtn.className = "xlsx-filter-btn-plain";
    clearBtn.textContent = "Clear filter";
    clearBtn.addEventListener("click", () => {
      filterState.delete(colId);
      callbacks.onFilterChange();
      closePopup();
      render(true);
    });
    footer.appendChild(clearBtn);

    const closeBtn = doc.createElement("button");
    closeBtn.className = "xlsx-filter-btn-plain xlsx-filter-close";
    closeBtn.textContent = "Done";
    closeBtn.addEventListener("click", () => closePopup());
    footer.appendChild(closeBtn);
    popup.appendChild(footer);

    doc.body.appendChild(popup);
    const rect = anchor.getBoundingClientRect();
    const width = 230;
    let left = rect.right - width;
    if (left < 8) left = 8;
    if (left + width > doc.documentElement.clientWidth - 8) left = doc.documentElement.clientWidth - width - 8;
    popup.style.width = width + "px";
    popup.style.left = left + "px";
    const below = rect.bottom + 4;
    const maxHeight = Math.min(320, doc.documentElement.clientHeight - below - 16);
    popup.style.top = below + "px";
    popup.style.maxHeight = maxHeight + "px";
    popupEl = popup;

    const onDocDown = (ev) => {
      if (popup.contains(ev.target)) return;
      closePopup();
    };
    const onKey = (ev) => {
      if (ev.key === "Escape") closePopup();
    };
    popup._onDocDown = onDocDown;
    popup._onKey = onKey;
    doc.addEventListener("mousedown", onDocDown, true);
    doc.addEventListener("keydown", onKey, true);
  }

  function closePopup() {
    if (!popupEl) return;
    if (popupEl._onDocDown) doc.removeEventListener("mousedown", popupEl._onDocDown, true);
    if (popupEl._onKey) doc.removeEventListener("keydown", popupEl._onKey, true);
    if (popupEl.parentNode) popupEl.parentNode.removeChild(popupEl);
    popupEl = null;
  }

  // ---------- public API ----------

  function getState() {
    const filters = [];
    if (G && G.filterRange) {
      for (const colInfo of filterColumns()) {
        const sel = filterState.get(colInfo.colId);
        let selected = null;
        if (sel instanceof Set) selected = Array.from(sel);
        else if (sel && sel.custom) selected = "(custom)";
        filters.push({
          colId: colInfo.colId,
          letter: colInfo.letter,
          headerText: colInfo.headerText,
          values: colInfo.values,
          blanks: colInfo.blanks,
          totalCount: colInfo.values.length + (colInfo.blanks ? 1 : 0),
          selected,
          all: sel == null,
        });
      }
    }
    return {
      sheetName: sheet.name,
      totalRows: G ? G.maxRow : 0,
      visibleRows: G ? G.visRows.length : 0,
      totalCols: G ? G.maxCol : 0,
      visibleCols: G ? G.visCols.length : 0,
      rowCapped: G ? G.rowCapped : false,
      frozenRows: sheet.freeze ? sheet.freeze.rows : 0,
      frozenCols: sheet.freeze ? sheet.freeze.cols : 0,
      filters,
    };
  }

  // Search reads the sheet model, not the rendered cells, because only the
  // rows near the viewport exist in the DOM. Matches are sorted in reading
  // order and the view scrolls to each one as it is reached.
  function search(query) {
    searchQuery = String(query == null ? "" : query).toLowerCase().trim();
    searchHits = [];
    searchIndex = -1;
    if (searchQuery) {
      for (const [cellKey, cell] of sheet.cells) {
        if (cell == null || cell.v == null) continue;
        const at = cellKey.indexOf(":");
        const r = Number(cellKey.slice(0, at));
        const c = Number(cellKey.slice(at + 1));
        if (!G.visRowIndex.has(r) || !G.visColIndex.has(c)) continue;
        const text = displayOf(r, c).text;
        if (text && text.toLowerCase().indexOf(searchQuery) !== -1) searchHits.push({ r, c });
      }
      searchHits.sort((a, b) => (a.r - b.r) || (a.c - b.c));
    }
    render(true);
    return searchHits.length;
  }

  let searchIndex = -1;
  let searchQuery = "";
  let searchHits = [];

  function searchNext(dir) {
    if (!searchHits.length) return null;
    searchIndex = (searchIndex + dir + searchHits.length * 2) % searchHits.length;
    const hit = searchHits[searchIndex];
    scrollToCell(colToLetter(hit.c) + hit.r);
    render(true);
    return { index: searchIndex + 1, count: searchHits.length };
  }

  function isSearchHit(r, c) {
    return searchQuery !== "" && displayOf(r, c).text.toLowerCase().indexOf(searchQuery) !== -1;
  }

  function getScrollPosition() {
    return scrollEl ? { top: scrollEl.scrollTop, left: scrollEl.scrollLeft } : { top: 0, left: 0 };
  }

  function setScrollPosition(pos) {
    if (!scrollEl || !pos) return;
    scrollEl.scrollTop = pos.top || 0;
    scrollEl.scrollLeft = pos.left || 0;
  }

  function setSettings(patch) {
    Object.assign(settings, patch);
  }

  function destroy() {
    destroyed = true;
    closePopup();
    closeTextPopup();
    dragging = false;
    doc.removeEventListener("mousemove", onGridMouseMove, true);
    doc.removeEventListener("mouseup", onGridMouseUp, true);
    if (scrollEl) scrollEl.removeEventListener("scroll", onScroll);
    if (selectedEl) selectedEl.classList.remove("xlsx-selected");
    selectedEl = null;
    if (container) container.textContent = "";
    rootEl = scrollEl = gridEl = null;
  }

  render(false);

  return {
    render,
    destroy,
    getState,
    search,
    searchNext,
    scrollToCell,
    setSettings,
    getScrollPosition,
    setScrollPosition,
    getSelection,
    copySelection,
    selectAll: selectAllVisble,
    getSizes,
    resetSizes,
    autoFitColumn,
    autoFitRow,
    showCellText,
    getSelectionRange: () => normalisedRange(),
    getFilterValues: filterColumns,
    setFilter(colId, selection) {
      filterState.set(colId, selection);
      render(true);
    },
    clearFilters() {
      filterState.clear();
      render(true);
    },
  };
}

module.exports = { createRenderer };
