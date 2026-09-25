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

// Colour maths shared by every renderer: CSS strings, OOXML theme colours with
// their tint and shade modifiers, and the DrawingML colour space transforms.

const { attr, firstOf, tagName } = require("./xml");

const INDEXED_COLORS = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
];

// Theme colour slot names in the order a:clrScheme and w:themeColor use them.
const THEME_SLOTS = [
  "dk1", "lt1", "dk2", "lt2",
  "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
  "hlink", "folHlink",
];

// OOXML colour names used by w:themeColor to index the same scheme.
const THEME_NAME_TO_INDEX = {
  dark1: 0, light1: 1, dark2: 2, light2: 3,
  accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9,
  hyperlink: 10, followedhyperlink: 11,
  background1: 1, text1: 0, background2: 3, text2: 2,
};

function parseHex(hex) {
  let h = String(hex).replace("#", "").trim();
  if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  if (h.length === 6) h = "FF" + h;
  if (h.length !== 8) return null;
  const n = parseInt(h, 16);
  if (isNaN(n)) return null;
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255, a: (n >> 24) & 255 };
}

function rgbToHex(r, g, b) {
  const to = (v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0");
  return "#" + to(r) + to(g) + to(b);
}

function toCss(c) {
  if (!c) return null;
  if (c.a >= 255) return rgbToHex(c.r, c.g, c.b);
  return "rgba(" + c.r + ", " + c.g + ", " + c.b + ", " + Math.round((c.a / 255) * 1000) / 1000 + ")";
}

function rgbToHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return { h: h / 6, s, l };
}

function hueToRgb(p, q, t) {
  if (t < 0) t += 1;
  if (t > 1) t -= 1;
  if (t < 1 / 6) return p + (q - p) * 6 * t;
  if (t < 1 / 2) return q;
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
  return p;
}

function hslToRgb(h, s, l) {
  if (s === 0) {
    const v = l * 255;
    return { r: v, g: v, b: v };
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return {
    r: hueToRgb(p, q, h + 1 / 3) * 255,
    g: hueToRgb(p, q, h) * 255,
    b: hueToRgb(p, q, h - 1 / 3) * 255,
  };
}

// Excel tint: positive tint moves the colour toward white, negative toward black.
function applyTint(hex, tint) {
  if (!tint) return hex;
  const rgb = parseHex(hex);
  if (!rgb) return hex;
  const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
  let l = hsl.l;
  l = tint < 0 ? l * (1 + tint) : l + (1 - l) * tint;
  const out = hslToRgb(hsl.h, hsl.s, Math.max(0, Math.min(1, l)));
  return rgbToHex(out.r, out.g, out.b);
}

function mixColors(a, b, t) {
  const pa = typeof a === "string" ? parseHex(a) : a;
  const pb = typeof b === "string" ? parseHex(b) : b;
  if (!pa || !pb) return typeof a === "string" ? a : null;
  const mix = (x, y) => Math.round(x + (y - x) * t);
  return rgbToHex(mix(pa.r, pb.r), mix(pa.g, pb.g), mix(pa.b, pb.b));
}

// DrawingML luminance modulation, used by a:lumMod and a:lumOff.
function applyLum(hex, mod, off) {
  const rgb = parseHex(hex);
  if (!rgb) return hex;
  const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
  let l = hsl.l * (mod == null ? 1 : mod) + (off == null ? 0 : off);
  l = Math.max(0, Math.min(1, l));
  const out = hslToRgb(hsl.h, hsl.s, l);
  return rgbToHex(out.r, out.g, out.b);
}

// DrawingML saturation modulation, used by a:satMod.
function applySat(hex, mult, off) {
  const rgb = parseHex(hex);
  if (!rgb) return hex;
  const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
  let s = hsl.s * (mult == null ? 1 : mult) + (off == null ? 0 : off);
  s = Math.max(0, Math.min(1, s));
  const out = hslToRgb(hsl.h, s, hsl.l);
  return rgbToHex(out.r, out.g, out.b);
}

function applyHue(hex, offset) {
  const rgb = parseHex(hex);
  if (!rgb) return hex;
  const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
  let h = (hsl.h + offset) % 1;
  if (h < 0) h += 1;
  const out = hslToRgb(h, hsl.s, hsl.l);
  return rgbToHex(out.r, out.g, out.b);
}

// ---------- theme parsing ----------

// Parses theme1.xml (or its text) into slot colours plus the minor and major
// latin fonts. Accepts either a parsed document or the raw XML string.
function parseTheme(input) {
  const theme = { colors: [], colorMap: {}, fonts: { minor: null, major: null } };
  if (!input) return theme;
  let doc = input;
  if (typeof input === "string") {
    try {
      doc = parseXml(input);
    } catch (err) {
      return theme;
    }
  }
  const scheme = findDescendant(doc.documentElement, "clrScheme");
  const byName = {};
  if (scheme) {
    for (const el of scheme.children || []) {
      const name = tagName(el).replace(/^.*:/, "");
      const srgb = firstOf(el, "srgbClr");
      const sys = firstOf(el, "sysClr");
      let hex = null;
      if (srgb) hex = attr(srgb, "val");
      else if (sys) hex = attr(sys, "lastClr") || "000000";
      if (hex) byName[name] = "#" + String(hex).replace(/^#/, "");
    }
  }
  theme.byName = byName;
  theme.colors = THEME_SLOTS.map((name) => byName[name] || null);
  for (const key of Object.keys(THEME_NAME_TO_INDEX)) {
    theme.colorMap[key] = byName[key] || null;
  }
  const fontScheme = findDescendant(doc.documentElement, "fontScheme");
  if (fontScheme) {
    for (const kind of ["minor", "major"]) {
      const font = firstOf(fontScheme, kind);
      const latin = font ? firstOf(font, "latin") : null;
      if (latin && attr(latin, "typeface")) theme.fonts[kind] = attr(latin, "typeface");
    }
  }
  return theme;
}

function findDescendant(el, tag) {
  if (!el) return null;
  const list = el.children || [];
  for (let i = 0; i < list.length; i++) {
    const child = list[i];
    if (tagName(child) === tag) return child;
    const found = findDescendant(child, tag);
    if (found) return found;
  }
  return null;
}

// Resolves a theme slot index or w:themeColor name to a hex colour.
function themeColor(theme, slot, tint) {
  if (!theme) return null;
  let hex = null;
  if (typeof slot === "number") hex = theme.colors ? theme.colors[slot] : null;
  else if (typeof slot === "string") {
    const key = String(slot).toLowerCase().replace(/[^a-z0-9]/g, "");
    if (theme.byName && theme.byName[key]) hex = theme.byName[key];
    else if (theme.colorMap && theme.colorMap[key]) hex = theme.colorMap[key];
    else if (THEME_NAME_TO_INDEX[key] != null && theme.colors) hex = theme.colors[THEME_NAME_TO_INDEX[key]];
  }
  if (!hex) return null;
  if (tint) hex = applyTint(hex, tint);
  return hex;
}

// ---------- colour element parsing ----------

// Reads the child element of a fill or colour container: <a:srgbClr> or
// <a:schemeClr> with optional transforms. Returns a CSS colour or null.
function colorFromContainer(container, theme, fallback) {
  const fb = fallback === undefined ? null : fallback;
  if (!container) return fb;
  for (const child of container.children || []) {
    const name = tagName(child);
    if (name === "srgbClr") {
      let hex = "#" + String(attr(child, "val") || "").replace(/^#/, "");
      return applyColorTransforms(hex, child);
    }
    if (name === "schemeClr") {
      const hex = themeColor(theme, attr(child, "val"), null);
      if (!hex) return fb;
      return applyColorTransforms(hex, child);
    }
    if (name === "sysClr") {
      const hex = "#" + String(attr(child, "lastClr") || "000000").replace(/^#/, "");
      return applyColorTransforms(hex, child);
    }
    if (name === "prstClr") {
      const hex = PRESET_COLORS[String(attr(child, "val") || "").toLowerCase()];
      if (!hex) return fb;
      return applyColorTransforms(hex, child);
    }
    if (name === "scrgbClr") return fb;
    if (name === "hslClr") {
      const h = parseFloat(attr(child, "hue") || "0") / 60000;
      const s = parseFloat(attr(child, "sat") || "0") / 100000;
      const l = parseFloat(attr(child, "lum") || "0") / 100000;
      const rgb = hslToRgb(h / 360, s, l);
      return applyColorTransforms(rgbToHex(rgb.r, rgb.g, rgb.b), child);
    }
  }
  return fb;
}

function applyColorTransforms(hex, el) {
  let out = hex;
  const list = el.children || [];
  const transforms = [];
  for (let i = 0; i < list.length; i++) {
    const t = list[i];
    const name = tagName(t);
    const val = attr(t, "val");
    const n = val === null ? null : parseFloat(val);
    if (name === "alpha") {
      const a = n == null ? 1 : n / 100000;
      out = cssWithAlpha(out, a);
    } else if (name === "lumMod") {
      transforms.push(["lumMod", n == null ? 1 : n / 100000]);
    } else if (name === "lumOff") {
      transforms.push(["lumOff", n == null ? 0 : n / 100000]);
    } else if (name === "satMod") {
      transforms.push(["satMod", n == null ? 1 : n / 100000]);
    } else if (name === "satOff") {
      transforms.push(["satOff", n == null ? 0 : n / 100000]);
    } else if (name === "hueMod") {
      transforms.push(["hueMod", n == null ? 1 : n / 100000]);
    } else if (name === "hueOff") {
      transforms.push(["hueOff", n == null ? 0 : n / 60000]);
    } else if (name === "shade") {
      transforms.push(["shade", n == null ? 1 : n / 100000]);
    } else if (name === "tint") {
      transforms.push(["tint", n == null ? 1 : n / 100000]);
    } else if (name === "inv") {
      transforms.push(["inv", 1]);
    } else if (name === "comp") {
      transforms.push(["comp", 1]);
    } else if (name === "gray") {
      transforms.push(["gray", 1]);
    }
  }
  if (!transforms.length) return out;
  let rgb = parseHex(out);
  if (!rgb) return out;
  let alpha = rgb.a;
  let hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
  let { h, s, l } = hsl;
  for (const [kind, v] of transforms) {
    switch (kind) {
      case "lumMod": l *= v; break;
      case "lumOff": l += v; break;
      case "satMod": s *= v; break;
      case "satOff": s += v; break;
      case "hueMod": h *= v; break;
      case "hueOff": h += v / 360; break;
      case "shade": l *= v; break;
      case "tint": l = l + (1 - l) * v; break;
      case "inv": { h = (h + 0.5) % 1; break; }
      case "comp": { h = (h + 0.5) % 1; s = Math.min(1, s * 1); break; }
      case "gray": { s = 0; break; }
      default: break;
    }
  }
  h = ((h % 1) + 1) % 1;
  s = Math.max(0, Math.min(1, s));
  l = Math.max(0, Math.min(1, l));
  const mixed = hslToRgb(h, s, l);
  if (alpha < 255) {
    return "rgba(" + Math.round(mixed.r) + ", " + Math.round(mixed.g) + ", " + Math.round(mixed.b) + ", " +
      Math.round((alpha / 255) * 1000) / 1000 + ")";
  }
  return rgbToHex(mixed.r, mixed.g, mixed.b);
}

function cssWithAlpha(css, alpha) {
  const rgb = parseHex(css) || parseRgbCss(css);
  if (!rgb) return css;
  return "rgba(" + rgb.r + ", " + rgb.g + ", " + rgb.b + ", " + Math.round(alpha * 1000) / 1000 + ")";
}

function parseRgbCss(css) {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(String(css));
  if (!m) return null;
  return { r: Number(m[1]), g: Number(m[2]), b: Number(m[3]), a: 255 };
}

const PRESET_COLORS = {
  black: "#000000", white: "#ffffff", red: "#ff0000", green: "#008000",
  blue: "#0000ff", yellow: "#ffff00", cyan: "#00ffff", magenta: "#ff00ff",
  gray: "#808080", grey: "#808080", orange: "#ffa500", purple: "#800080",
  brown: "#a52a2a", pink: "#ffc0cb", lime: "#00ff00", navy: "#000080",
  silver: "#c0c0c0", maroon: "#800000", olive: "#808000", teal: "#008080",
  aqua: "#00ffff", gold: "#ffd700", indigo: "#4b0082", violet: "#ee82ee",
  darkblue: "#00008b", darkgreen: "#006400", darkred: "#8b0000",
  lightblue: "#add8e6", lightgreen: "#90ee90", lightgrey: "#d3d3d3",
  lightgray: "#d3d3d3", darkgrey: "#a9a9a9", darkgray: "#a9a9a9",
};

// ---------- spreadsheet colour references ----------

// A spreadsheet colour is { rgb, theme, tint, indexed, auto } as parsed from
// styles.xml. Returns a CSS colour or the fallback.
function resolveColor(color, theme, fallback) {
  const fb = fallback === undefined ? null : fallback;
  if (!color) return fb;
  if (color.auto) return fb;
  if (color.rgb) {
    const c = parseHex(color.rgb);
    if (!c) return fb;
    if (c.a >= 255) return rgbToHex(c.r, c.g, c.b);
    return "rgba(" + c.r + ", " + c.g + ", " + c.b + ", " + Math.round((c.a / 255) * 1000) / 1000 + ")";
  }
  if (color.theme != null) {
    const base = themeColor(theme, color.theme, null);
    if (!base) return fb;
    return color.tint ? applyTint(base, color.tint) : base;
  }
  if (color.indexed != null) {
    if (color.indexed === 64) return "#000000";
    if (color.indexed === 65) return "#FFFFFF";
    const hex = INDEXED_COLORS[color.indexed];
    if (!hex) return fb;
    let out = hex;
    if (color.tint) out = applyTint(out, color.tint);
    const parsed = parseHex(out);
    return parsed ? rgbToHex(parsed.r, parsed.g, parsed.b) : fb;
  }
  return fb;
}

// True when the colour means "automatic" (no explicit colour given).
function isAutoColor(color) {
  if (!color) return true;
  if (color.auto) return true;
  if (color.theme === 1 && !color.tint) return true;
  return false;
}

// Auto text colour on a coloured fill: black on light fills, white on dark
// ones, so a dark theme cannot turn black text white on a pastel cell.
function autoTextColorForFill(fillCss) {
  const rgb = parseHex(fillCss) || parseRgbCss(fillCss);
  if (!rgb) return null;
  const lum = (0.299 * rgb.r + 0.587 * rgb.g + 0.114 * rgb.b) / 255;
  return lum > 0.62 ? "#000000" : "#ffffff";
}

function isDarkColor(css) {
  const rgb = parseHex(css) || parseRgbCss(css);
  if (!rgb) return false;
  return (0.299 * rgb.r + 0.587 * rgb.g + 0.114 * rgb.b) / 255 < 0.5;
}

module.exports = {
  INDEXED_COLORS,
  THEME_SLOTS,
  THEME_NAME_TO_INDEX,
  PRESET_COLORS,
  parseHex,
  rgbToHex,
  toCss,
  applyTint,
  applyLum,
  applySat,
  applyHue,
  mixColors,
  resolveColor,
  isAutoColor,
  parseTheme,
  themeColor,
  colorFromContainer,
  applyColorTransforms,
  autoTextColorForFill,
  isDarkColor,
};
