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

const { attr, attrInt, attrBool, childrenOf, firstOf, parseXml } = require("../shared/xml");
const { resolveColor } = require("../shared/color");
const { getFormatCode } = require("./numfmt");

// Maps a font name to a CSS stack. Common Office fonts get metric compatible
// fallbacks so line breaks look close to Excel on Linux.
const FONT_FALLBACKS = {
  "Arial": "Arial, Helvetica, Liberation Sans, sans-serif",
  "Arial Narrow": "\"Arial Narrow\", \"Liberation Sans Narrow\", Arial, sans-serif",
  "Aptos": "Aptos, Calibri, Carlito, sans-serif",
  "Aptos Narrow": "\"Aptos Narrow\", \"Arial Narrow\", Arial, sans-serif",
  "Calibri": "Calibri, Carlito, \"Segoe UI\", sans-serif",
  "Cambria": "Cambria, \"Liberation Serif\", serif",
  "Consolas": "Consolas, \"Liberation Mono\", monospace",
  "Courier New": "\"Courier New\", \"Liberation Mono\", monospace",
  "Georgia": "Georgia, \"Liberation Serif\", serif",
  "Segoe UI": "\"Segoe UI\", \"Noto Sans\", sans-serif",
  "Times New Roman": "\"Times New Roman\", \"Liberation Serif\", serif",
  "Verdana": "Verdana, \"DejaVu Sans\", sans-serif",
};

function fontFamilyCss(name) {
  if (!name) return "sans-serif";
  if (FONT_FALLBACKS[name]) return FONT_FALLBACKS[name];
  if (/^[A-Za-z0-9 ]+$/.test(name)) return "\"" + name + "\", sans-serif";
  return name;
}

function parseColorEl(el) {
  if (!el) return null;
  const rgb = attr(el, "rgb");
  const theme = attr(el, "theme");
  const indexed = attr(el, "indexed");
  const auto = attr(el, "auto");
  const tint = attr(el, "tint");
  const color = {};
  if (rgb != null) color.rgb = rgb;
  if (theme != null) color.theme = parseInt(theme, 10);
  if (indexed != null) color.indexed = parseInt(indexed, 10);
  if (auto != null) color.auto = auto === "1" || auto === "true";
  if (tint != null) color.tint = parseFloat(tint);
  if (Object.keys(color).length === 0) return null;
  return color;
}

function parseFont(el) {
  const font = { name: "Calibri", size: 11, bold: false, italic: false, underline: false, strike: false, color: null, vertAlign: null };
  const name = firstOf(el, "name");
  if (name) font.name = attr(name, "val") || font.name;
  const sz = firstOf(el, "sz");
  if (sz) font.size = parseFloat(attr(sz, "val")) || font.size;
  if (firstOf(el, "b")) font.bold = true;
  if (firstOf(el, "i")) font.italic = true;
  const u = firstOf(el, "u");
  if (u) font.underline = attr(u, "val") || "single";
  if (firstOf(el, "strike")) font.strike = true;
  const color = firstOf(el, "color");
  if (color) font.color = parseColorEl(color);
  const vert = firstOf(el, "vertAlign");
  if (vert) font.vertAlign = attr(vert, "val");
  return font;
}

function parseFill(el) {
  const pattern = firstOf(el, "patternFill");
  if (pattern) {
    const patternType = attr(pattern, "patternType");
    if (!patternType || patternType === "none" || patternType === "gray125") return null;
    const fg = parseColorEl(firstOf(pattern, "fgColor"));
    if (patternType === "solid") {
      if (!fg) return null;
      return { type: "solid", color: fg };
    }
    if (!fg) return null;
    const alpha = patternType === "darkGray" ? 0.5
      : patternType === "mediumGray" ? 0.35
      : patternType === "lightGray" ? 0.2
      : patternType === "gray0625" ? 0.08
      : 0.3;
    return { type: "pattern", color: fg, alpha };
  }
  const gradient = firstOf(el, "gradientFill");
  if (gradient) {
    const degree = parseFloat(attr(gradient, "degree"));
    const stops = [];
    for (const stop of childrenOf(gradient, "stop")) {
      stops.push({
        position: parseFloat(attr(stop, "position")) || 0,
        color: parseColorEl(firstOf(stop, "color")),
      });
    }
    if (stops.length > 0) {
      return { type: "gradient", degree: isNaN(degree) ? 0 : degree, stops };
    }
  }
  return null;
}

function parseBorderEdge(el) {
  if (!el) return null;
  const style = attr(el, "style");
  if (!style) return null;
  return { style, color: parseColorEl(firstOf(el, "color")) };
}

function parseBorder(el) {
  return {
    left: parseBorderEdge(firstOf(el, "left")),
    right: parseBorderEdge(firstOf(el, "right")),
    top: parseBorderEdge(firstOf(el, "top")),
    bottom: parseBorderEdge(firstOf(el, "bottom")),
    diagonalUp: (firstOf(el, "diagonal") && attr(el, "diagonalUp") === "1") || false,
    diagonal: parseBorderEdge(firstOf(el, "diagonal")),
  };
}

const STYLE_WEIGHTS = {
  hair: 1, dotted: 2, dashed: 3, dashDot: 4, dashDotDot: 5,
  thin: 6, mediumDashed: 7, mediumDashDot: 8, mediumDashDotDot: 9,
  slantDashDot: 10, medium: 11, thick: 12, double: 13,
};

const BORDER_CSS = {
  hair: { width: 1, css: "solid" },
  dotted: { width: 1, css: "dotted" },
  dashed: { width: 1, css: "dashed" },
  dashDot: { width: 1, css: "dashed" },
  dashDotDot: { width: 1, css: "dashed" },
  slantDashDot: { width: 1, css: "dashed" },
  thin: { width: 1, css: "solid" },
  mediumDashed: { width: 2, css: "dashed" },
  mediumDashDot: { width: 2, css: "dashed" },
  mediumDashDotDot: { width: 2, css: "dashed" },
  medium: { width: 2, css: "solid" },
  thick: { width: 3, css: "solid" },
  double: { width: 3, css: "double" },
};

function borderSpec(edge, theme) {
  if (!edge || !edge.style) return null;
  const css = BORDER_CSS[edge.style] || { width: 1, css: "solid" };
  const color = resolveColor(edge.color, theme, "#000000") || "#000000";
  return {
    style: edge.style,
    weight: STYLE_WEIGHTS[edge.style] || 1,
    width: css.width,
    css: css.css,
    color,
  };
}

// Picks the stronger of two border specs, the way Excel resolves shared edges.
function strongerBorder(a, b) {
  if (!a) return b;
  if (!b) return a;
  if (a.weight !== b.weight) return a.weight > b.weight ? a : b;
  return a;
}

function parseAlignment(el) {
  if (!el) return {};
  const rot = attrInt(el, "textRotation", 0);
  return {
    h: attr(el, "horizontal"),
    v: attr(el, "vertical"),
    wrap: attr(el, "wrapText") === "1",
    shrink: attr(el, "shrinkToFit") === "1",
    indent: attrInt(el, "indent", 0),
    rot,
  };
}

function parseXf(el) {
  return {
    numFmtId: attrInt(el, "numFmtId", 0),
    fontId: attrInt(el, "fontId", 0),
    fillId: attrInt(el, "fillId", 0),
    borderId: attrInt(el, "borderId", 0),
    xfId: attr(el, "xfId"),
    alignment: parseAlignment(firstOf(el, "alignment")),
  };
}

function parseDxf(el) {
  const font = firstOf(el, "font");
  const fill = firstOf(el, "fill");
  const border = firstOf(el, "border");
  const numFmt = firstOf(el, "numFmt");
  return {
    font: font ? parseFont(font) : null,
    fill: fill ? parseFill(fill) : null,
    border: border ? parseBorder(border) : null,
    numFmt: numFmt ? { id: attrInt(numFmt, "numFmtId", 0), code: attr(numFmt, "formatCode") || null } : null,
  };
}

function fillToCss(fill, theme) {
  if (!fill) return null;
  if (fill.type === "solid") return resolveColor(fill.color, theme, null);
  if (fill.type === "pattern") {
    const base = resolveColor(fill.color, theme, null);
    if (!base) return null;
    if (base.startsWith("#")) {
      const r = parseInt(base.slice(1, 3), 16);
      const g = parseInt(base.slice(3, 5), 16);
      const b = parseInt(base.slice(5, 7), 16);
      return "rgba(" + r + ", " + g + ", " + b + ", " + fill.alpha + ")";
    }
    return base;
  }
  return null;
}

function gradientToCss(fill, theme) {
  const stops = [];
  for (const stop of fill.stops) {
    const color = resolveColor(stop.color, theme, "#FFFFFF");
    stops.push(color + " " + Math.round(stop.position * 100) + "%");
  }
  return "linear-gradient(" + (fill.degree || 90) + "deg, " + stops.join(", ") + ")";
}

function parseStyles(xmlText, theme) {
  const model = {
    numFmts: new Map(),
    fonts: [],
    fills: [],
    borders: [],
    xfs: [],
    dxfs: [],
    theme,
  };
  if (!xmlText) {
    model.xfs.push(parseXf({ getAttribute: () => null, children: [] }));
    return buildApi(model);
  }
  const doc = parseXml(xmlText);
  const root = doc.documentElement;

  const numFmts = firstOf(root, "numFmts");
  if (numFmts) {
    for (const nf of childrenOf(numFmts, "numFmt")) {
      model.numFmts.set(attrInt(nf, "numFmtId", 0), attr(nf, "formatCode") || "");
    }
  }
  const fonts = firstOf(root, "fonts");
  if (fonts) for (const f of childrenOf(fonts, "font")) model.fonts.push(parseFont(f));
  const fills = firstOf(root, "fills");
  if (fills) for (const f of childrenOf(fills, "fill")) model.fills.push(parseFill(f));
  const borders = firstOf(root, "borders");
  if (borders) for (const b of childrenOf(borders, "border")) model.borders.push(parseBorder(b));
  const cellXfs = firstOf(root, "cellXfs");
  if (cellXfs) for (const xf of childrenOf(cellXfs, "xf")) model.xfs.push(parseXf(xf));
  const dxfs = firstOf(root, "dxfs");
  if (dxfs) for (const dxf of childrenOf(dxfs, "dxf")) model.dxfs.push(parseDxf(dxf));

  return buildApi(model);
}

function buildApi(model) {
  const cache = new Map();
  const theme = model.theme;

  model.resolveXf = function (idx) {
    const index = idx || 0;
    if (cache.has(index)) return cache.get(index);
    const xf = model.xfs[index] || model.xfs[0] || { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, alignment: {} };
    const font = model.fonts[xf.fontId] || model.fonts[0] || { name: "Calibri", size: 11, color: null };
    const fill = model.fills[xf.fillId] || null;
    const border = model.borders[xf.borderId] || {};
    const resolved = {
      xfIndex: index,
      font: {
        name: font.name,
        family: fontFamilyCss(font.name),
        sizePt: font.size,
        sizePx: Math.round((font.size * 96) / 72 * 100) / 100,
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        strike: font.strike,
        color: font.color,
      },
      fill: fill,
      borders: {
        left: borderSpec(border.left, theme),
        right: borderSpec(border.right, theme),
        top: borderSpec(border.top, theme),
        bottom: borderSpec(border.bottom, theme),
      },
      align: xf.alignment || {},
      numFmtId: xf.numFmtId,
      fmt: getFormatCode(xf.numFmtId, model.numFmts),
    };
    cache.set(index, resolved);
    return resolved;
  };

  model.resolveFill = function (fill) {
    if (!fill) return null;
    return fill.type === "gradient" ? gradientToCss(fill, theme) : fillToCss(fill, theme);
  };
  model.gradientCss = function (fill) {
    return fill && fill.type === "gradient" ? gradientToCss(fill, theme) : null;
  };
  model.resolveFontColor = function (color, fallback) {
    return resolveColor(color, theme, fallback === undefined ? null : fallback);
  };
  model.dxfStyle = function (idx) {
    const dxf = model.dxfs[idx];
    if (!dxf) return null;
    return {
      font: dxf.font ? {
        bold: dxf.font.bold,
        italic: dxf.font.italic,
        strike: dxf.font.strike,
        underline: dxf.font.underline,
        color: dxf.font.color ? resolveColor(dxf.font.color, theme, null) : null,
      } : null,
      fill: dxf.fill ? gradientOrSolid(dxf.fill, theme) : null,
      border: dxf.border,
    };
  };
  function gradientOrSolid(fill) {
    if (fill.type === "gradient") return null;
    return fillToCss(fill, theme);
  }
  return model;
}

module.exports = {
  parseStyles,
  parseColorEl,
  fontFamilyCss,
  strongerBorder,
  borderSpec,
};
