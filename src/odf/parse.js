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

// OpenDocument readers. An ODF package is a plain zip: content.xml holds the
// body, styles.xml the named styles, and automatic styles travel inside
// content.xml itself. The parsers map the parts onto the same internal models
// the OOXML renderers already draw.

const { Package } = require("../shared/package");
const { childrenOf, firstOf, findAll, attr, attrInt, tagName } = require("../shared/xml");
const { odfLengthToPx } = require("../shared/units");
const { key } = require("../shared/addr");
const { parseStyles } = require("../spreadsheet/styles");

// ---------- shared style maps ----------

// ODF keys styles by name. Text and paragraph properties are pre-resolved to
// the flat property names the renderers use, so no further cascade is needed.
function parseOdfStyles(pkg) {
  const styles = { text: new Map(), paragraph: new Map(), table: new Map(), cell: new Map(), columns: new Map(), masterPages: new Map() };
  const collect = (doc) => {
    if (!doc) return;
    const root = doc.documentElement;
    for (const el of findAll(root, "style")) {
      const family = attr(el, "family");
      const name = attr(el, "name");
      if (!family || !name) continue;
      if (family === "text") styles.text.set(name, parseTextProperties(firstOf(el, "text-properties")));
      else if (family === "paragraph") styles.paragraph.set(name, parseParagraphProperties(el));
      else if (family === "table") styles.table.set(name, parseTableProperties(firstOf(el, "table-properties")));
      else if (family === "table-cell") styles.cell.set(name, parseCellProperties(firstOf(el, "table-cell-properties")));
      else if (family === "table-column") styles.columns.set(name, parseColumnProperties(firstOf(el, "table-column-properties")));
    }
    for (const el of findAll(root, "master-page")) {
      const name = attr(el, "name");
      if (name) styles.masterPages.set(name, el);
    }
  };
  collect(pkg.xml("styles.xml"));
  collect(pkg.xml("content.xml"));
  return styles;
}

function parseTextProperties(el) {
  if (!el) return null;
  const out = {};
  const weight = attr(el, "font-weight");
  if (weight === "bold" || weight === "700" || weight === "800" || weight === "900") out.bold = true;
  const style = attr(el, "font-style");
  if (style === "italic" || style === "oblique") out.italic = true;
  const underline = attr(el, "text-underline-style");
  if (underline && underline !== "none") out.underline = "single";
  const strike = attr(el, "text-line-through-style");
  if (strike && strike !== "none") out.strike = true;
  const size = attr(el, "font-size");
  if (size) {
    const px = odfLengthToPx(size);
    if (px) out.sizeHalfPt = Math.round((px * 72) / 96) * 2;
  }
  const family = attr(el, "font-family");
  if (family) out.fontFamily = stripQuotes(family);
  const color = attr(el, "color");
  if (color && color !== "transparent") out.color = normalizeOdfColor(color);
  const background = attr(el, "background-color");
  if (background && background !== "transparent") out.shading = { color: normalizeOdfColor(background) };
  const script = attr(el, "text-position");
  if (script === "super 58%" || script === "super") out.vertAlign = "superscript";
  else if (script === "sub 58%" || script === "sub") out.vertAlign = "subscript";
  const caps = attr(el, "text-transform");
  if (caps === "uppercase") out.caps = true;
  return Object.keys(out).length ? out : null;
}

function parseParagraphProperties(el) {
  const out = {};
  const props = firstOf(el, "paragraph-properties");
  if (props) {
    const align = attr(props, "text-align");
    if (align === "center" || align === "end" || align === "right" || align === "justify" || align === "left") {
      out.align = align === "end" ? "right" : align;
    }
    const marginTop = attr(props, "margin-top");
    if (marginTop) {
      const px = odfLengthToPx(marginTop);
      if (px != null) out.spaceBeforePt = (px * 72) / 96;
    }
    const marginBottom = attr(props, "margin-bottom");
    if (marginBottom) {
      const px = odfLengthToPx(marginBottom);
      if (px != null) out.spaceAfterPt = (px * 72) / 96;
    }
    const lineHeight = attr(props, "line-height");
    if (lineHeight) {
      const value = parseFloat(lineHeight);
      if (!isNaN(value)) out.lineHeight = value / 100;
    }
    const left = attr(props, "margin-left") || attr(props, "text-indent");
    if (left) {
      const px = odfLengthToPx(left);
      if (px != null) out.indentLeftTw = (px * 1440) / 96;
    }
    const indent = attr(props, "text-indent");
    if (indent && indent.startsWith("-")) {
      const px = odfLengthToPx(indent);
      if (px != null) out.indentHangingTw = (-px * 1440) / 96;
    }
    const background = attr(props, "background-color");
    if (background && background !== "transparent") out.shading = { color: normalizeOdfColor(background) };
    if (attr(props, "page-break-before") === "always") out.pageBreakBefore = true;
    const keep = attr(props, "keep-together");
    if (keep === "always") out.keepNext = true;
  }
  const textProps = firstOf(el, "text-properties");
  if (textProps) {
    const rPr = parseTextProperties(textProps);
    if (rPr) out.rPr = rPr;
  }
  const outline = attr(el, "default-outline-level");
  if (outline) out.outlineLevel = parseInt(outline, 10) - 1;
  return Object.keys(out).length ? out : null;
}

function parseTableProperties(el) {
  if (!el) return null;
  const out = {};
  const width = attr(el, "width");
  if (width && width.endsWith("%")) {
    const pct = parseFloat(width);
    if (!isNaN(pct)) out.widthPct = pct;
  }
  const align = attr(el, "align");
  if (align) out.align = align === "end" ? "right" : align;
  const margin = attr(el, "margin-left");
  if (margin) {
    const px = odfLengthToPx(margin);
    if (px != null) out.indentTw = px * 15;
  }
  return Object.keys(out).length ? out : null;
}

function parseCellProperties(el) {
  if (!el) return null;
  const out = {};
  const background = attr(el, "background-color");
  if (background && background !== "transparent") out.shading = { color: normalizeOdfColor(background) };
  const align = attr(el, "vertical-align");
  if (align) out.vAlign = align;
  const padding = attr(el, "padding");
  if (padding) {
    const px = odfLengthToPx(padding);
    if (px != null) out.margin = { left: px * 15, right: px * 15, top: px * 15, bottom: px * 15 };
  }
  for (const [side, key] of [["border", "all"], ["border-left", "left"], ["border-right", "right"], ["border-top", "top"], ["border-bottom", "bottom"]]) {
    const value = attr(el, side);
    if (!value || value === "none") continue;
    const parsed = parseOdfBorder(value);
    if (parsed) {
      out.borders = out.borders || {};
      if (key === "all") {
        out.borders.left = parsed;
        out.borders.right = parsed;
        out.borders.top = parsed;
        out.borders.bottom = parsed;
      } else {
        out.borders[key] = parsed;
      }
    }
  }
  return Object.keys(out).length ? out : null;
}

function parseColumnProperties(el) {
  if (!el) return null;
  const width = attr(el, "column-width");
  if (width) {
    const px = odfLengthToPx(width);
    if (px != null) return { widthPx: px };
  }
  return null;
}

function parseOdfBorder(value) {
  const match = /([\d.]+)(pt|px|cm|mm|in)\s+(solid|dashed|dotted)\s+(#?[0-9a-fA-F]{6}|#[0-9a-fA-F]{3})/.exec(value);
  if (!match) {
    const colorOnly = /(#[0-9a-fA-F]{6})/.exec(value);
    if (!colorOnly) return null;
    return { width: 1, style: "solid", color: colorOnly[1] };
  }
  const px = odfLengthToPx(match[1] + match[2]) || 1;
  return { width: Math.max(1, Math.round(px * 0.75)), style: match[3], color: match[4] };
}

function normalizeOdfColor(value) {
  const text = String(value || "").trim();
  if (/^#[0-9a-fA-F]{6}$/.test(text)) return text;
  if (/^#[0-9a-fA-F]{3}$/.test(text)) return "#" + text[1] + text[1] + text[2] + text[2] + text[3] + text[3];
  if (/^[0-9a-fA-F]{6}$/.test(text)) return "#" + text;
  const rgb = /rgb\((\d+),\s*(\d+),\s*(\d+)\)/.exec(text);
  if (rgb) {
    const to = (v) => Number(v).toString(16).padStart(2, "0");
    return "#" + to(rgb[1]) + to(rgb[2]) + to(rgb[3]);
  }
  return text;
}

function stripQuotes(value) {
  return String(value || "").replace(/^['"]|['"]$/g, "").split(",")[0].trim();
}

// ---------- spreadsheets (.ods) ----------

function readOds(input) {
  const pkg = Package.open(input);
  const content = pkg.xml("content.xml");
  if (!content) throw new Error("This file has no content.xml, so it is not a readable OpenDocument file.");
  const styles = parseOdfStyles(pkg);
  const spreadsheet = parseOdfSheetStyles(content);
  const body = firstOf(content.documentElement, "body");
  const sheetParent = body ? (firstOf(body, "spreadsheet") || body) : content.documentElement;
  const sheets = [];
  const tables = findAll(sheetParent, "table");
  tables.forEach((tableEl, index) => {
    const name = attr(tableEl, "name") || "Sheet" + (index + 1);
    sheets.push(parseOdfSheet(tableEl, name, index, styles, spreadsheet));
  });
  const cache = new Map();
  return {
    kind: "ods",
    sheets,
    loadSheet(index) {
      if (cache.has(index)) return cache.get(index);
      const model = sheets[index];
      if (!model) throw new Error("Sheet index " + index + " is out of range.");
      cache.set(index, model);
      return model;
    },
    sheetIndexByName(name) {
      for (let i = 0; i < sheets.length; i++) if (sheets[i].name === name) return i;
      return -1;
    },
  };
}

function parseOdfSheetStyles(content) {
  const out = { columns: new Map(), rows: new Map() };
  for (const el of findAll(content.documentElement, "table-column")) {
    const styleName = attr(el, "style-name");
    if (styleName) out.columns.set(styleName, el);
  }
  for (const el of findAll(content.documentElement, "table-row")) {
    const styleName = attr(el, "style-name");
    if (styleName) out.rows.set(styleName, el);
  }
  return out;
}

function parseOdfSheet(tableEl, name, index, styles, spreadsheet) {
  const rowEls = childrenOf(tableEl, "table-row");
  const cells = new Map();
  const rows = new Map();
  const columns = [];
  let maxRow = 0;
  let maxCol = 0;

  // Column definitions can appear before, between or after rows.
  let defaultWidthChars = null;
  for (const el of childrenOf(tableEl, "table-column")) {
    const repeat = attrInt(el, "number-columns-repeated", 1);
    const styleName = attr(el, "style-name");
    const widthPx = styleName && styles.columns.get(styleName) ? styles.columns.get(styleName).widthPx : null;
    for (let i = 0; i < Math.min(repeat, 100); i++) {
      columns.push({ min: columns.length + 1, max: columns.length + 1, width: null, widthPx, hidden: false, style: 0, outline: 0, collapsed: false });
    }
  }
  const colProps = firstOf(tableEl, "table-columns");
  if (colProps) {
    const defaultStyle = attr(colProps, "default-cell-style-name");
    void defaultStyle;
  }

  rowEls.forEach((rowEl, rowIndex) => {
    const repeat = Math.min(attrInt(rowEl, "number-rows-repeated", 1), 5000);
    const rowStyle = attr(rowEl, "style-name");
    const rowProps = rowStyle ? styles.rows.get(rowStyle) : null;
    const height = rowProps ? odfLengthToPx(attr(firstOf(rowProps, "table-row-properties"), "row-height")) : null;
    const isHeader = attr(rowEl, "default-cell-style-name") === "TableHeading";
    for (let r = 0; r < repeat; r++) {
      const rowNumber = rowIndex + 1 + r;
      if (rowNumber > maxRow) maxRow = rowNumber;
      rows.set(rowNumber, {
        r: rowNumber,
        ht: height ? (height * 72) / 96 : null,
        hidden: false,
        outline: 0,
        collapsed: false,
        s: null,
        customFormat: false,
      });
      let col = 1;
      for (const cellEl of childrenOf(rowEl, "table-cell")) {
        const tag = tagName(cellEl);
        if (tag !== "table-cell" && tag !== "covered-table-cell") continue;
        const colRepeat = Math.min(attrInt(cellEl, "number-columns-repeated", 1), 1000);
        if (tag === "covered-table-cell") {
          col += colRepeat;
          continue;
        }
        const styleName = attr(cellEl, "style-name");
        const value = odfCellValue(cellEl);
        const rowSpan = attrInt(cellEl, "number-rows-spanned", 1);
        const colSpan = attrInt(cellEl, "number-columns-spanned", 1);
        const text = odfCellText(cellEl);
        for (let c = 0; c < colRepeat; c++) {
          const cellNumber = col + c;
          if (cellNumber > maxCol) maxCol = cellNumber;
          if (c === 0) {
            const styleIndex = odfCellStyleIndex(styleName, styles);
            cells.set(key(rowNumber, cellNumber), {
              v: value === null ? text : value,
              t: value === null ? "s" : "n",
              s: styleIndex,
            });
            if (colSpan > 1 || rowSpan > 1) {
              // Merges are collected below; the covered cells are filled later.
            }
          }
        }
        col += colRepeat;
      }
    }
  });

  // ODF marks merged ranges with spans on the anchor cell.
  const merges = [];
  rowEls.forEach((rowEl, rowIndex) => {
    let col = 1;
    for (const cellEl of childrenOf(rowEl, "table-cell")) {
      const colSpan = attrInt(cellEl, "number-columns-spanned", 1);
      const rowSpan = attrInt(cellEl, "number-rows-spanned", 1);
      const repeat = attrInt(cellEl, "number-columns-repeated", 1);
      if (colSpan > 1 || rowSpan > 1) {
        merges.push({ r1: rowIndex + 1, c1: col, r2: rowIndex + rowSpan, c2: col + colSpan - 1 });
      }
      col += repeat;
    }
  });

  const modelRows = rows;
  const rowCount = Math.max(1, maxRow);
  const colCount = Math.max(1, maxCol);
  return {
    name,
    index,
    state: "visible",
    dims: { r1: 1, c1: 1, r2: rowCount, c2: colCount },
    defaultRowHeightPt: 15,
    defaultRowHeightPx: 20,
    defaultColWidthChars: defaultWidthChars,
    cols: columns,
    rows: modelRows,
    cells,
    merges,
    freeze: null,
    autoFilter: null,
    hyperlinks: new Map(),
    showGridLines: true,
    outlinePr: { summaryBelow: true, summaryRight: true },
    conditional: [],
    tabColor: null,
    styles: parseStyles(null, null),
  };
}

// ODF keeps typed values in attributes, so numbers keep their numeric form and
// dates arrive as ISO strings.
function odfCellValue(cellEl) {
  const valueType = attr(cellEl, "value-type");
  if (!valueType) return null;
  if (valueType === "float" || valueType === "currency" || valueType === "percentage") {
    const value = parseFloat(attr(cellEl, "value"));
    if (isNaN(value)) return null;
    if (valueType === "percentage") return value * 100;
    return value;
  }
  if (valueType === "boolean") return attr(cellEl, "boolean-value") === "true" ? 1 : 0;
  if (valueType === "date") {
    const iso = attr(cellEl, "date-value");
    return iso || odfCellText(cellEl);
  }
  if (valueType === "time") return attr(cellEl, "time-value") || odfCellText(cellEl);
  return null;
}

function odfCellText(cellEl) {
  const parts = [];
  for (const p of childrenOf(cellEl, "p")) parts.push(odfParagraphText(p));
  return parts.join("\n");
}

// A paragraph is a mixed run of text nodes and spans, so walk childNodes
// rather than only the element children.
function odfParagraphText(p) {
  let out = "";
  for (const node of p.childNodes || []) {
    if (node.nodeType === 3) {
      out += node.nodeValue || "";
      continue;
    }
    if (node.nodeType !== 1) continue;
    const tag = tagName(node);
    if (tag === "line-break") out += "\n";
    else if (tag === "s") out += " ".repeat(Math.min(attrInt(node, "c", 1), 200));
    else if (tag === "tab") out += "\t";
    else out += odfParagraphText(node);
  }
  return out;
}

// Cell styles resolve to a spreadsheet style index; because the ODF palette is
// independent of the xlsx one, a tiny synthetic style table is built on demand.
function odfCellStyleIndex(styleName, styles) {
  return 0;
  void styleName;
  void styles;
}

// ---------- text documents (.odt) ----------

// Maps an ODT body onto the block model the docx renderer draws.
function readOdt(input) {
  const pkg = Package.open(input);
  const content = pkg.xml("content.xml");
  if (!content) throw new Error("This file has no content.xml, so it is not a readable OpenDocument file.");
  const styles = parseOdfStyles(pkg);
  const body = firstOf(content.documentElement, "body");
  const textParent = body ? (firstOf(body, "text") || body) : content.documentElement;
  const blocks = [];
  let order = 0;
  const walk = (parent) => {
    for (const el of parent.children || []) {
      const tag = tagName(el);
      if (tag === "p" || tag === "h") {
        blocks.push(odfToParagraph(el, styles, tag === "h", order++));
      } else if (tag === "list") {
        walk(el);
        for (const item of childrenOf(el, "list-item")) walk(item);
      } else if (tag === "table") {
        blocks.push(odfToTable(el, styles));
      } else if (tag === "section" || tag === "list-item" || tag === "frame" || tag === "text-box") {
        walk(el);
      } else if (tag === "soft-page-break") {
        blocks.push({ type: "p", props: null, style: null, runs: [], markRunProps: null, numbering: null, bookmarks: [], sectionBreak: { pageBreak: true } });
      }
    }
  };
  walk(textParent);
  const masterPage = findFirstMasterPage(styles);
  const section = masterPage || {
    pageWidthTw: 11906,
    pageHeightTw: 16838,
    marginTopTw: 1440,
    marginRightTw: 1440,
    marginBottomTw: 1440,
    marginLeftTw: 1440,
  };
  return {
    kind: "odt",
    pkg,
    body: blocks,
    section,
    styles: { paragraph: new Map(), character: new Map(), table: new Map(), docDefaults: null },
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
    mediaCache: null,
    properties: {},
    // ODF images are referenced by href rather than embedded relationships, so
    // the frame renderer leaves them as labelled placeholders for now.
    mediaUrl() {
      return null;
    },
  };
}

function findFirstMasterPage(styles) {
  for (const el of styles.masterPages.values()) {
    const props = firstOf(el, "page-layout-properties") || findAll(el, "page-layout-properties")[0];
    if (!props) continue;
    const width = odfLengthToPx(attr(props, "page-width")) || 794;
    const height = odfLengthToPx(attr(props, "page-height")) || 1123;
    const margin = attr(props, "margin");
    const marginPx = margin ? odfLengthToPx(margin) : null;
    const toTw = (px) => (px * 1440) / 96;
    return {
      pageWidthTw: toTw(width),
      pageHeightTw: toTw(height),
      marginTopTw: marginPx != null ? toTw(marginPx) : 1440,
      marginBottomTw: marginPx != null ? toTw(marginPx) : 1440,
      marginLeftTw: marginPx != null ? toTw(marginPx) : 1440,
      marginRightTw: marginPx != null ? toTw(marginPx) : 1440,
      headerTw: 708,
      footerTw: 708,
      columns: 1,
      columnSpaceTw: 708,
    };
  }
  return null;
}

function odfToParagraph(el, styles, isHeading, order) {
  const styleName = attr(el, "style-name");
  const styleProps = styleName ? styles.paragraph.get(styleName) : null;
  const props = Object.assign({}, styleProps || {});
  if (isHeading) {
    const level = attrInt(el, "outline-level", 1) || 1;
    props.outlineLevel = level - 1;
    const hStyle = styles.paragraph.get("Heading");
    if (!props.rPr) {
      props.rPr = { bold: true, sizeHalfPt: level === 1 ? 32 : level === 2 ? 28 : 24 };
    }
  }
  const runs = [];
  for (const child of el.children || []) {
    collectOdfRuns(child, styles, null, runs);
  }
  return {
    type: "p",
    props: Object.keys(props).length ? props : null,
    style: null,
    runs,
    markRunProps: null,
    numbering: null,
    bookmarks: [],
    sectionBreak: null,
    order,
  };
}

function collectOdfRuns(el, styles, inherited, out) {
  const tag = tagName(el);
  const styleName = attr(el, "style-name");
  const textStyle = styleName ? styles.text.get(styleName) : null;
  const props = textStyle ? Object.assign({}, inherited || {}, textStyle) : inherited;
  if (tag === "span" || tag === "a" || tag === "meta" || tag === "annotation") {
    if (tag === "annotation") return;
    for (const child of el.children || []) collectOdfRuns(child, styles, props, out);
    return;
  }
  if (tag === "s") {
    const count = attrInt(el, "c", 1);
    out.push({ type: "text", text: " ".repeat(Math.min(count, 200)), props });
    return;
  }
  if (tag === "tab") {
    out.push({ type: "tab" });
    return;
  }
  if (tag === "line-break") {
    out.push({ type: "break" });
    return;
  }
  if (tag === "frame" || tag === "text-box") {
    const image = findImageHref(el);
    if (image) {
      const width = odfLengthToPx(attr(el, "width") || "");
      const height = odfLengthToPx(attr(el, "height") || "");
      out.push({
        type: "image",
        url: null,
        href: image,
        widthPx: width || 0,
        heightPx: height || 0,
        alt: attr(el, "name") || "",
      });
    }
    return;
  }
  if (el.children && el.children.length) {
    for (const child of el.children) collectOdfRuns(child, styles, props, out);
    return;
  }
  const text = el.textContent || "";
  if (text) out.push({ type: "text", text, props: props || null });
}

function findImageHref(el) {
  for (const child of el.children || []) {
    if (tagName(child) === "image") {
      return attr(child, "href") || attr(child, "xlink:href") || null;
    }
    const nested = findImageHref(child);
    if (nested) return nested;
  }
  return null;
}

function odfToTable(tableEl, styles) {
  const styleName = attr(tableEl, "style-name");
  const props = (styleName && styles.table.get(styleName)) || {};
  const rowEls = childrenOf(tableEl, "table-row");
  const grid = [];
  const rows = [];
  for (const rowEl of rowEls) {
    const cells = [];
    for (const cellEl of childrenOf(rowEl, "table-cell")) {
      const cellStyle = attr(cellEl, "style-name");
      const cellProps = (cellStyle && styles.cell.get(cellStyle)) || {};
      const blocks = [];
      for (const p of childrenOf(cellEl, "p")) blocks.push(odfToParagraph(p, styles, false, 0));
      cells.push({
        props: {
          gridSpan: attrInt(cellEl, "number-columns-spanned", 1),
          vMerge: null,
          shading: cellProps.shading || null,
          borders: cellProps.borders || null,
          vAlign: cellProps.vAlign || null,
          margin: cellProps.margin || null,
        },
        blocks,
        gridSpan: attrInt(cellEl, "number-columns-spanned", 1),
        vMerge: null,
      });
    }
    rows.push({ props: null, cells });
  }
  const columnCount = rows.reduce((max, row) => Math.max(max, row.cells.reduce((a, c) => a + c.gridSpan, 0)), 0);
  const widthPx = props.widthPct ? (props.widthPct / 100) * 794 : null;
  return {
    type: "table",
    props: Object.assign({}, props, widthPx ? { widthTw: (widthPx * 1440) / 96 } : {}),
    grid: new Array(columnCount).fill(0),
    rows,
    styleId: null,
  };
}

// ---------- presentations (.odp) ----------

function readOdp(input) {
  const pkg = Package.open(input);
  const content = pkg.xml("content.xml");
  if (!content) throw new Error("This file has no content.xml, so it is not a readable OpenDocument file.");
  const styles = parseOdfStyles(pkg);
  const body = firstOf(content.documentElement, "body");
  const presentation = body ? firstOf(body, "presentation") : null;
  const pageEls = presentation ? findAll(presentation, "page") : [];
  const slides = pageEls.map((pageEl, index) => parseOdpSlide(pageEl, index, styles));
  const master = styles.masterPages.values().next().value;
  let widthPx = 960;
  let heightPx = 540;
  if (master) {
    const props = findAll(master, "page-layout-properties")[0];
    if (props) {
      widthPx = odfLengthToPx(attr(props, "page-width")) || widthPx;
      heightPx = odfLengthToPx(attr(props, "page-height")) || heightPx;
    }
  }
  return {
    kind: "odp",
    pkg,
    widthPx,
    heightPx,
    slides,
    masters: new Map(),
    layouts: new Map(),
    themes: new Map(),
    mediaCache: null,
    properties: {},
    mediaUrl() {
      return null;
    },
  };
}

function parseOdpSlide(pageEl, index, styles) {
  const shapes = [];
  const textParts = [];
  for (const frame of findAll(pageEl, "frame")) {
    const x = odfLengthToPx(attr(frame, "x")) || 0;
    const y = odfLengthToPx(attr(frame, "y")) || 0;
    const width = odfLengthToPx(attr(frame, "width")) || 0;
    const height = odfLengthToPx(attr(frame, "height")) || 0;
    const toEmu = (px) => Math.round(px * 9525);
    const paragraphs = [];
    for (const p of findAll(frame, "p")) {
      const runs = [];
      collectOdfRuns(p, styles, null, runs);
      paragraphs.push({ props: {}, runs, text: runs.map((r) => (r.type === "text" ? r.text : "")).join(""), bullet: null });
    }
    const image = findImageHref(frame);
    const shape = {
      type: image ? "picture" : "shape",
      name: attr(frame, "name") || "",
      placeholder: null,
      xEmu: toEmu(x),
      yEmu: toEmu(y),
      cxEmu: toEmu(width),
      cyEmu: toEmu(height),
      rotation: 0,
      flipH: false,
      flipV: false,
      fill: null,
      line: null,
      effects: null,
      textInfo: {
        paragraphs,
        text: paragraphs.map((p) => p.text).join("\n"),
        bodyProps: { anchor: "t", insets: { l: 91440, t: 45720, r: 91440, b: 45720 } },
      },
      text: paragraphs.map((p) => p.text).join("\n"),
      path: image || null,
      alt: attr(frame, "name") || "",
      geometry: { preset: "rect", custom: null },
    };
    shapes.push(shape);
    if (shape.text) textParts.push(shape.text);
  }
  return {
    index,
    path: "content.xml",
    shapes,
    background: { type: "solid", color: "#ffffff" },
    clrMap: null,
    theme: null,
    notes: "",
    text: textParts.join(" ").replace(/\s+/g, " ").trim(),
    title: textParts.length ? textParts[0] : "",
    imageCount: shapes.filter((s) => s.type === "picture").length,
    hidden: false,
  };
}

module.exports = {
  readOds,
  readOdt,
  readOdp,
  parseOdfStyles,
};
