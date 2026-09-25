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

// Reads word/styles.xml into the style cascade the renderer needs: document
// defaults, paragraph styles with their basedOn chains, character styles and
// the table style fragments.

const { childrenOf, firstOf, attr, attrBool, attrInt, flag, tagName } = require("../shared/xml");
const { colorFromContainer, themeColor } = require("../shared/color");

// Word's built-in style names, used when a document relies on the defaults that
// styles.xml only declares by id.
const BUILTIN_HEADING = /^heading\s*([1-9])$/i;
const BUILTIN_NAMES = {
  normal: "Normal",
  "default paragraph font": "DefaultParagraphFont",
};

function parseStyles(doc, theme) {
  const styles = {
    docDefaults: null,
    paragraph: new Map(),
    character: new Map(),
    table: new Map(),
    numbering: new Map(),
    theme,
    defaultParagraphStyleId: "Normal",
    defaultCharacterStyleId: "DefaultParagraphFont",
    defaultTableStyleId: "TableNormal",
    latent: new Map(),
  };
  if (!doc) return styles;
  const root = doc.documentElement;
  const dd = firstOf(root, "docDefaults");
  if (dd) {
    styles.docDefaults = {
      rPr: parseRunProps(firstOf(dd, "rPrDefault") ? firstOf(firstOf(dd, "rPrDefault"), "rPr") : null, theme),
      pPr: parseParaProps(firstOf(dd, "pPrDefault") ? firstOf(firstOf(dd, "pPrDefault"), "pPr") : null, theme),
    };
  }

  for (const el of childrenOf(root, "style")) {
    const type = attr(el, "type") || "paragraph";
    const id = attr(el, "styleId");
    if (!id) continue;
    const entry = {
      id,
      type,
      name: attr(firstOf(el, "name"), "val") || id,
      basedOn: attr(firstOf(el, "basedOn"), "val"),
      link: attr(firstOf(el, "link"), "val"),
      next: attr(firstOf(el, "next"), "val"),
      uiPriority: attrInt(firstOf(el, "uiPriority"), "val", 0),
      hidden: flag(firstOf(el, "semiHidden"), false),
      custom: attrBool(el, "customStyle", false),
      default: flag(firstOf(el, "qFormat"), false),
      pPr: null,
      rPr: null,
      tblPr: null,
    };
    const pPrEl = firstOf(el, "pPr");
    if (pPrEl) entry.pPr = parseParaProps(pPrEl, theme);
    const rPrEl = firstOf(el, "rPr");
    if (rPrEl) entry.rPr = parseRunProps(rPrEl, theme);
    const tblPrEl = firstOf(el, "tblPr");
    if (tblPrEl) entry.tblPr = parseTableProps(tblPrEl, theme);
    if (type === "paragraph") styles.paragraph.set(id, entry);
    else if (type === "character") styles.character.set(id, entry);
    else if (type === "table") styles.table.set(id, entry);
    else if (type === "numbering") styles.numbering.set(id, entry);
  }

  const defaults = firstOf(root, "docDefaults");
  void defaults;
  const settingsDefaults = findStyleDefault(root, "paragraph");
  if (settingsDefaults) styles.defaultParagraphStyleId = settingsDefaults;
  const charDefault = findStyleDefault(root, "character");
  if (charDefault) styles.defaultCharacterStyleId = charDefault;
  const tableDefault = findStyleDefault(root, "table");
  if (tableDefault) styles.defaultTableStyleId = tableDefault;

  for (const el of childrenOf(root, "latentStyles")) {
    const name = attr(el, "name");
    if (name) styles.latent.set(name.toLowerCase(), el);
  }
  return styles;
}

function findStyleDefault(root, type) {
  for (const el of childrenOf(root, "style")) {
    if ((attr(el, "type") || "paragraph") !== type) continue;
    if (attrBool(el, "default", false)) return attr(el, "styleId");
  }
  return null;
}

// ---------- property parsing ----------

function parseParaProps(el, theme) {
  if (!el) return null;
  const out = {};
  const pStyle = firstOf(el, "pStyle");
  if (pStyle) out.styleId = attr(pStyle, "val");
  const jc = firstOf(el, "jc");
  if (jc) out.align = jcToCss(attr(jc, "val"));
  const numPr = firstOf(el, "numPr");
  if (numPr) {
    out.numId = attrInt(firstOf(numPr, "numId"), "val", null);
    out.ilvl = attrInt(firstOf(numPr, "ilvl"), "val", 0);
  }
  const spacing = firstOf(el, "spacing");
  if (spacing) {
    const before = attrInt(spacing, "before", null);
    const after = attrInt(spacing, "after", null);
    const line = attrInt(spacing, "line", null);
    const rule = attr(spacing, "lineRule") || "auto";
    if (before !== null && before >= 0) out.spaceBeforePt = before / 20;
    if (after !== null && after >= 0) out.spaceAfterPt = after / 20;
    if (line !== null && line >= 0) {
      out.lineRule = rule;
      if (rule === "auto") out.lineHeight = line / 240;
      else out.lineHeightPt = line / 20;
    }
    const beforeAuto = flag(firstOf(spacing, "beforeAutospacing"), false);
    const afterAuto = flag(firstOf(spacing, "afterAutospacing"), false);
    if (beforeAuto) out.spaceBeforeAuto = true;
    if (afterAuto) out.spaceAfterAuto = true;
    out.spacingContextual = true;
  }
  const ind = firstOf(el, "ind");
  if (ind) {
    const left = attrInt(ind, "left", null);
    const start = attrInt(ind, "start", null);
    const right = attrInt(ind, "right", null);
    const end = attrInt(ind, "end", null);
    const hanging = attrInt(ind, "hanging", null);
    const firstLine = attrInt(ind, "firstLine", null);
    if (left !== null || start !== null) out.indentLeftTw = left !== null ? left : start;
    if (right !== null || end !== null) out.indentRightTw = right !== null ? right : end;
    if (hanging !== null && hanging !== 0) out.indentHangingTw = hanging;
    else if (firstLine !== null && firstLine !== 0) out.indentFirstLineTw = firstLine;
  }
  const keepNext = firstOf(el, "keepNext");
  if (keepNext) out.keepNext = flag(keepNext, true);
  const pageBreakBefore = firstOf(el, "pageBreakBefore");
  if (pageBreakBefore) out.pageBreakBefore = flag(pageBreakBefore, true);
  const pBdr = firstOf(el, "pBdr");
  if (pBdr) out.borders = parseBorders(pBdr, theme);
  const shd = firstOf(el, "shd");
  if (shd) out.shading = parseShading(shd, theme);
  const sectPr = firstOf(el, "sectPr");
  if (sectPr) out.sectPr = parseSectionProps(sectPr, theme);
  const outlineLvl = firstOf(el, "outlineLvl");
  if (outlineLvl) out.outlineLevel = attrInt(outlineLvl, "val", 0);
  const rPr = firstOf(el, "rPr");
  if (rPr) out.rPr = parseRunProps(rPr, theme);
  const tabs = firstOf(el, "tabs");
  if (tabs) out.tabs = parseTabs(tabs);
  const contextualSpacing = firstOf(el, "contextualSpacing");
  if (contextualSpacing) out.contextualSpacing = flag(contextualSpacing, true);
  const textAlignment = firstOf(el, "textAlignment");
  if (textAlignment) out.textAlignment = attr(textAlignment, "val");
  const snapToGrid = firstOf(el, "snapToGrid");
  if (snapToGrid) out.snapToGrid = flag(snapToGrid, true);
  const wid = firstOf(el, "widowControl");
  if (wid) out.widowControl = flag(wid, true);
  return out;
}

function parseTabs(el) {
  const out = [];
  for (const tab of childrenOf(el, "tab")) {
    const pos = attrInt(tab, "pos", null);
    if (pos === null) continue;
    out.push({ posTw: pos, align: attr(tab, "val") || "left", leader: attr(tab, "leader") || "none" });
  }
  return out;
}

function jcToCss(val) {
  switch (val) {
    case "center": return "center";
    case "right":
    case "end": return "right";
    case "both":
    case "distribute": return "justify";
    case "start": return "left";
    default: return "left";
  }
}

function parseRunProps(el, theme) {
  if (!el) return null;
  const out = {};
  const rFonts = firstOf(el, "rFonts");
  if (rFonts) {
    const ascii = attr(rFonts, "ascii") || attr(rFonts, "hAnsi") || attr(rFonts, "cs");
    const asciiTheme = attr(rFonts, "asciiTheme") || attr(rFonts, "hAnsiTheme");
    if (ascii) out.fontFamily = ascii;
    else out.fontFamilyTheme = asciiTheme;
    const eastAsia = attr(rFonts, "eastAsia");
    if (eastAsia) out.fontFamilyEastAsia = eastAsia;
    out.fontHint = attr(rFonts, "hint");
  }
  const b = firstOf(el, "b");
  if (b) out.bold = flag(b, true);
  const bCs = firstOf(el, "bCs");
  if (bCs && out.bold === undefined) out.bold = flag(bCs, true);
  const i = firstOf(el, "i");
  if (i) out.italic = flag(i, true);
  const iCs = firstOf(el, "iCs");
  if (iCs && out.italic === undefined) out.italic = flag(iCs, true);
  const u = firstOf(el, "u");
  if (u) out.underline = attr(u, "val") === "none" ? false : (attr(u, "val") || "single");
  const strike = firstOf(el, "strike");
  if (strike) out.strike = flag(strike, true);
  const dstrike = firstOf(el, "dstrike");
  if (dstrike) out.strike = flag(dstrike, true);
  const caps = firstOf(el, "caps");
  if (caps) out.caps = flag(caps, true);
  const smallCaps = firstOf(el, "smallCaps");
  if (smallCaps) out.smallCaps = flag(smallCaps, true);
  const sz = firstOf(el, "sz");
  if (sz) {
    const v = attrInt(sz, "val", null);
    if (v !== null && v > 0) out.sizeHalfPt = v;
  }
  const color = firstOf(el, "color");
  if (color) {
    const val = attr(color, "val");
    if (val === "auto") out.colorAuto = true;
    else if (val) out.color = "#" + String(val).replace(/^#/, "");
    const themeName = attr(color, "themeColor");
    const tint = attrInt(color, "themeTint", null);
    const shade = attrInt(color, "themeShade", null);
    if (themeName && theme) {
      const base = themeColor(theme, themeName, null);
      if (base) out.color = base;
      out.themeColorName = themeName;
      if (tint !== null) out.themeTintHex = tint;
      if (shade !== null) out.themeShadeHex = shade;
    }
  }
  const highlight = firstOf(el, "highlight");
  if (highlight) out.highlight = HIGHLIGHT_COLORS[attr(highlight, "val")] || null;
  const shd = firstOf(el, "shd");
  if (shd) out.shading = parseShading(shd, theme);
  const vertAlign = firstOf(el, "vertAlign");
  if (vertAlign) out.vertAlign = attr(vertAlign, "val");
  const spacing = firstOf(el, "spacing");
  if (spacing) {
    const v = attrInt(spacing, "val", null);
    if (v !== null) out.letterSpacingPt = v / 20;
  }
  const position = firstOf(el, "position");
  if (position) {
    const v = attrInt(position, "val", null);
    if (v !== null) out.positionHalfPt = v;
  }
  const em = firstOf(el, "em");
  if (em) out.emphasis = attr(em, "val");
  const vanish = firstOf(el, "vanish");
  if (vanish) out.hidden = flag(vanish, true);
  const rStyle = firstOf(el, "rStyle");
  if (rStyle) out.styleId = attr(rStyle, "val");
  return out;
}

const HIGHLIGHT_COLORS = {
  black: "#000000",
  blue: "#0000ff",
  cyan: "#00ffff",
  green: "#00ff00",
  magenta: "#ff00ff",
  red: "#ff0000",
  yellow: "#ffff00",
  white: "#ffffff",
  darkBlue: "#000080",
  darkCyan: "#008080",
  darkGreen: "#008000",
  darkMagenta: "#800080",
  darkRed: "#800000",
  darkYellow: "#808000",
  darkGray: "#808080",
  lightGray: "#c0c0c0",
  none: null,
};

function parseShading(el, theme) {
  if (!el) return null;
  const val = attr(el, "val") || "clear";
  if (val === "clear" || val === "nil") {
    const fill = attr(el, "fill");
    if (!fill || fill === "auto") return null;
    return { type: "solid", color: "#" + String(fill).replace(/^#/, "") };
  }
  const fill = attr(el, "fill");
  const color = attr(el, "color");
  const base = fill && fill !== "auto" ? "#" + fill.replace(/^#/, "") : null;
  const patternColor = color && color !== "auto" ? "#" + color.replace(/^#/, "") : null;
  if (!base && !patternColor) return null;
  return { type: "pattern", color: patternColor || base, base: base || "#ffffff", pattern: val };
}

function parseBorders(el, theme) {
  const out = {};
  for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"]) {
    const b = firstOf(el, side);
    if (!b) continue;
    const val = attr(b, "val") || "single";
    if (val === "none" || val === "nil") continue;
    const sz = attrInt(b, "sz", 4);
    const color = attr(b, "color");
    out[side] = {
      width: Math.max(1, Math.round((sz / 8) * 1.5)),
      style: BORDER_STYLES[val] || "solid",
      color: color && color !== "auto" ? "#" + color.replace(/^#/, "") : "#000000",
      space: attrInt(b, "space", 0),
    };
  }
  return Object.keys(out).length ? out : null;
}

const BORDER_STYLES = {
  single: "solid",
  thick: "solid",
  double: "double",
  dotted: "dotted",
  dashed: "dashed",
  dotDash: "dashed",
  dotDotDash: "dashed",
  wave: "solid",
  doubleWave: "solid",
  dashSmallGap: "dashed",
  threeDEmboss: "solid",
  threeDEngrave: "solid",
  inset: "solid",
  outset: "solid",
};

function parseTableProps(el, theme) {
  if (!el) return null;
  const out = {};
  const style = firstOf(el, "tblStyle");
  if (style) out.styleId = attr(style, "val");
  const w = firstOf(el, "tblW");
  if (w) {
    const type = attr(w, "type") || "dxa";
    const v = attrInt(w, "w", 0);
    if (type === "dxa") out.widthTw = v;
    else if (type === "pct") out.widthPct = v / 50;
    out.widthType = type;
  }
  const jc = firstOf(el, "jc");
  if (jc) out.align = attr(jc, "val");
  const layout = firstOf(el, "tblLayout");
  if (layout) out.layout = attr(layout, "type");
  const borders = firstOf(el, "tblBorders");
  if (borders) out.borders = parseBorders(borders, theme);
  const shd = firstOf(el, "shd");
  if (shd) out.shading = parseShading(shd, theme);
  const cellMar = firstOf(el, "tblCellMar");
  if (cellMar) {
    out.cellMargin = {};
    for (const side of ["top", "left", "bottom", "right"]) {
      const m = firstOf(cellMar, side);
      if (m) out.cellMargin[side] = attrInt(m, "w", 0);
    }
  }
  const look = firstOf(el, "tblLook");
  if (look) {
    out.look = {
      firstRow: attrBool(look, "firstRow", false),
      lastRow: attrBool(look, "lastRow", false),
      firstColumn: attrBool(look, "firstColumn", false),
      lastColumn: attrBool(look, "lastColumn", false),
      noHBand: attrBool(look, "noHBand", false),
      noVBand: attrBool(look, "noVBand", false),
    };
  }
  const indent = firstOf(el, "tblInd");
  if (indent) out.indentTw = attrInt(indent, "w", null);
  void theme;
  return out;
}

function parseRowProps(el) {
  if (!el) return null;
  const out = {};
  const trHeight = firstOf(el, "trHeight");
  if (trHeight) {
    out.heightTw = attrInt(trHeight, "val", null);
    out.heightRule = attr(trHeight, "hRule") || "auto";
  }
  if (firstOf(el, "tblHeader")) out.header = true;
  if (firstOf(el, "cantSplit")) out.cantSplit = true;
  return out;
}

function parseCellProps(el, theme) {
  if (!el) return null;
  const out = {};
  const tcW = firstOf(el, "tcW");
  if (tcW) {
    const type = attr(tcW, "type") || "dxa";
    const v = attrInt(tcW, "w", 0);
    if (type === "dxa") out.widthTw = v;
    else if (type === "pct") out.widthPct = v / 50;
  }
  const gridSpan = firstOf(el, "gridSpan");
  if (gridSpan) out.gridSpan = attrInt(gridSpan, "val", 1);
  const vMerge = firstOf(el, "vMerge");
  if (vMerge) out.vMerge = attr(vMerge, "val") || "continue";
  const shd = firstOf(el, "shd");
  if (shd) out.shading = parseShading(shd, theme);
  const borders = firstOf(el, "tcBorders");
  if (borders) out.borders = parseBorders(borders, theme);
  const valign = firstOf(el, "vAlign");
  if (valign) out.vAlign = attr(valign, "val");
  const tcMar = firstOf(el, "tcMar");
  if (tcMar) {
    out.margin = {};
    for (const side of ["top", "left", "bottom", "right"]) {
      const m = firstOf(tcMar, side);
      if (m) out.margin[side] = attrInt(m, "w", 0);
    }
  }
  const span = firstOf(el, "textDirection");
  if (span) out.textDirection = attr(span, "val");
  out.hideMark = Boolean(firstOf(el, "hideMark"));
  return out;
}

function parseSectionProps(el, theme) {
  if (!el) return null;
  const out = {};
  const pgSz = firstOf(el, "pgSz");
  if (pgSz) {
    out.pageWidthTw = attrInt(pgSz, "w", 11906);
    out.pageHeightTw = attrInt(pgSz, "h", 16838);
    out.orientation = attr(pgSz, "orient") || "portrait";
  }
  const pgMar = firstOf(el, "pgMar");
  if (pgMar) {
    out.marginTopTw = attrInt(pgMar, "top", 1440);
    out.marginRightTw = attrInt(pgMar, "right", 1440);
    out.marginBottomTw = attrInt(pgMar, "bottom", 1440);
    out.marginLeftTw = attrInt(pgMar, "left", 1440);
    out.headerTw = attrInt(pgMar, "header", 708);
    out.footerTw = attrInt(pgMar, "footer", 708);
    out.gutterTw = attrInt(pgMar, "gutter", 0);
  }
  const cols = firstOf(el, "cols");
  if (cols) {
    out.columns = attrInt(cols, "num", 1);
    out.columnSpaceTw = attrInt(cols, "space", 708);
  }
  const headerRef = firstOf(el, "headerReference");
  if (headerRef) out.headerRefs = collectRefs(el, "headerReference");
  const footerRef = firstOf(el, "footerReference");
  if (footerRef) out.footerRefs = collectRefs(el, "footerReference");
  const type = firstOf(el, "type");
  if (type) out.type = attr(type, "val");
  const pgNumType = firstOf(el, "pgNumType");
  if (pgNumType) out.pageNumberFormat = attr(pgNumType, "fmt");
  const titlePg = firstOf(el, "titlePg");
  if (titlePg) out.titlePage = flag(titlePg, true);
  const docGrid = firstOf(el, "docGrid");
  if (docGrid) {
    out.docGrid = attrInt(docGrid, "linePitch", null);
    // The type decides whether the grid is actually applied: only "lines" and
    // "linesAndChars" snap to it. Word writes a linePitch into almost every
    // document with the default type, where the pitch is informational.
    out.docGridType = attr(docGrid, "type") || "default";
  }
  void theme;
  return out;
}

function collectRefs(parent, tag) {
  const out = {};
  for (const el of childrenOf(parent, tag)) {
    const type = attr(el, "type") || "default";
    const rid = attr(el, "id") || attr(el, "r:id");
    if (rid) out[type] = rid;
  }
  return out;
}

// ---------- numbering ----------

// Flattens numbering.xml into numId -> abstract definition with levels.
function parseNumbering(doc, theme) {
  const out = { numIdToAbstract: new Map(), abstract: new Map(), counters: new Map() };
  if (!doc) return out;
  const root = doc.documentElement;
  const abstractIds = new Map();
  for (const el of childrenOf(root, "abstractNum")) {
    const id = attrInt(el, "abstractNumId", null);
    if (id === null) continue;
    const levels = [];
    for (const lvl of childrenOf(el, "lvl")) {
      const ilvl = attrInt(lvl, "ilvl", 0);
      const start = attrInt(firstOf(lvl, "start"), "val", 1);
      const numFmt = attr(firstOf(lvl, "numFmt"), "val") || "decimal";
      const lvlText = attr(firstOf(lvl, "lvlText"), "val") || "";
      const lvlJc = attr(firstOf(lvl, "lvlJc"), "val") || "left";
      const suff = attr(firstOf(lvl, "suff"), "val") || "tab";
      const pPr = firstOf(lvl, "pPr");
      const rPr = firstOf(lvl, "rPr");
      const picBullet = firstOf(lvl, "lvlPicBulletId");
      levels[ilvl] = {
        start,
        numFmt,
        lvlText,
        lvlJc,
        suff,
        pPr: pPr ? parseParaProps(pPr, theme) : null,
        rPr: rPr ? parseRunProps(rPr, theme) : null,
        isLgl: flag(firstOf(lvl, "isLgl"), false),
        tplc: attr(lvl, "tplc"),
        picBulletId: picBullet ? attrInt(picBullet, "val", null) : null,
      };
    }
    const styleLink = firstOf(el, "numStyleLink");
    const styleLinkId = firstOf(el, "styleLink");
    abstractIds.set(id, {
      id,
      levels,
      multiLevelType: attr(firstOf(el, "multiLevelType"), "val"),
      numStyleLink: styleLink ? attr(styleLink, "val") : null,
      styleLink: styleLinkId ? attr(styleLinkId, "val") : null,
    });
  }
  for (const el of childrenOf(root, "num")) {
    const numId = attrInt(el, "numId", null);
    const abstractEl = firstOf(el, "abstractNumId");
    const abstractId = abstractEl ? attrInt(abstractEl, "val", null) : null;
    if (numId === null || abstractId === null) continue;
    const overrides = new Map();
    for (const o of childrenOf(el, "lvlOverride")) {
      const ilvl = attrInt(o, "ilvl", 0);
      const startOverride = firstOf(o, "startOverride");
      const lvlEl = firstOf(o, "lvl");
      overrides.set(ilvl, {
        startOverride: startOverride ? attrInt(startOverride, "val", null) : null,
        level: lvlEl ? {
          numFmt: attr(firstOf(lvlEl, "numFmt"), "val"),
          lvlText: attr(firstOf(lvlEl, "lvlText"), "val"),
          start: attrInt(firstOf(lvlEl, "start"), "val", 1),
        } : null,
      });
    }
    out.numIdToAbstract.set(numId, { abstractId, overrides });
  }
  out.abstract = abstractIds;
  return out;
}

module.exports = {
  parseStyles,
  parseParaProps,
  parseRunProps,
  parseTableProps,
  parseRowProps,
  parseCellProps,
  parseSectionProps,
  parseShading,
  parseBorders,
  parseNumbering,
  jcToCss,
  HIGHLIGHT_COLORS,
  BORDER_STYLES,
  BUILTIN_HEADING,
  BUILTIN_NAMES,
};
