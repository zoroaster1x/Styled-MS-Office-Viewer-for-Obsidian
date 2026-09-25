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

// Conditional formatting: the rules, their thresholds and what they draw.
//
// Kept apart from the renderer because the interesting part is arithmetic, not
// DOM: a top ten rule needs the values of its range, a data bar needs the range
// minimum and maximum, and an icon set needs the bucket boundaries. The range
// is scanned once per conditional block and the numbers are kept, so a hundred
// cells in the same range do not each rescan it.
//
// The values come in through a reader callback, which is what makes this
// testable without a workbook: test/cf.mjs hands it an array.

const MAX_RANGE_CELLS = 50000;

// Three traffic lights, three arrows, three symbols, three flags, four and
// five ratings, and the star and triangle sets. Straight text glyphs keep the
// plugin free of icon assets.
const ICON_SETS = {
  "3TrafficLights1": { glyphs: ["\u25CF", "\u25CF", "\u25CF"], colors: ["#63BE7B", "#FFEB84", "#F8696B"] },
  "3TrafficLights2": { glyphs: ["\u25CF", "\u25CF", "\u25CF"], colors: ["#63BE7B", "#FFEB84", "#F8696B"] },
  "3Arrows": { glyphs: ["\u25B2", "\u25B6", "\u25BC"], colors: ["#63BE7B", "#FFEB84", "#F8696B"] },
  "3ArrowsGray": { glyphs: ["\u25B2", "\u25B6", "\u25BC"], colors: ["#808080", "#A6A6A6", "#BFBFBF"] },
  "3Symbols": { glyphs: ["\u2714", "!", "\u2716"], colors: ["#63BE7B", "#FFEB84", "#F8696B"] },
  "3Symbols2": { glyphs: ["\u2714", "!", "\u2716"], colors: ["#63BE7B", "#FFEB84", "#F8696B"] },
  "3Flags": { glyphs: ["\u2691", "\u2691", "\u2691"], colors: ["#63BE7B", "#FFEB84", "#F8696B"] },
  "4Rating": { glyphs: ["\u25CF", "\u25CF", "\u25CF", "\u25CF"], colors: ["#63BE7B", "#B1D580", "#FFEB84", "#F8696B"] },
  "5Rating": { glyphs: ["\u25CF", "\u25CF", "\u25CF", "\u25CF", "\u25CF"], colors: ["#63BE7B", "#B1D580", "#FFEB84", "#F4B183", "#F8696B"] },
  "4TrafficLights": { glyphs: ["\u25CF", "\u25CF", "\u25CF", "\u25CF"], colors: ["#63BE7B", "#B1D580", "#FFEB84", "#F8696B"] },
  "5Boxes": { glyphs: ["\u25A0", "\u25A0", "\u25A0", "\u25A0", "\u25A0"], colors: ["#63BE7B", "#B1D580", "#FFEB84", "#F4B183", "#F8696B"] },
};

function iconSetSpec(name) {
  return ICON_SETS[name] || ICON_SETS["3TrafficLights1"];
}

// The numbers behind one conditional block, scanned once. min and max come from
// the values, not from the first cell the renderer happens to ask about, which
// is what makes a data bar or a colour scale honest.
function blockStats(block, readValue) {
  if (block.__stats) return block.__stats;
  const values = [];
  let scanned = 0;
  for (const range of block.ranges || []) {
    for (let r = range.r1; r <= range.r2 && scanned < MAX_RANGE_CELLS; r++) {
      for (let c = range.c1; c <= range.c2 && scanned < MAX_RANGE_CELLS; c++) {
        scanned++;
        const value = readValue(r, c);
        if (typeof value === "number" && isFinite(value)) values.push(value);
      }
    }
  }
  values.sort((a, b) => a - b);
  let sum = 0;
  for (const value of values) sum += value;
  const stats = {
    values,
    count: values.length,
    min: values.length ? values[0] : 0,
    max: values.length ? values[values.length - 1] : 0,
    average: values.length ? sum / values.length : 0,
  };
  block.__stats = stats;
  return stats;
}

// A cfvo entry is a boundary: a literal number, a percentage or percentile of
// the range, or the range's own minimum, maximum or average.
function thresholdOf(cfvo, stats, value, kind) {
  if (!cfvo) return null;
  const raw = cfvo.val != null ? parseFloat(cfvo.val) : NaN;
  switch (cfvo.type) {
    case "num":
      return isNaN(raw) ? null : raw;
    case "percent": {
      const span = stats.max - stats.min;
      if (!span) return stats.min;
      return isNaN(raw) ? null : stats.min + (raw / 100) * span;
    }
    case "percentile": {
      if (!stats.count || isNaN(raw)) return null;
      const at = Math.min(stats.count - 1, Math.max(0, Math.round((raw / 100) * (stats.count - 1))));
      return stats.values[at];
    }
    case "min":
      return stats.min;
    case "max":
      return stats.max;
    case "formula":
      return isNaN(raw) ? null : raw;
    default:
      // "autoMin" and "autoMax" in the x14 extension, and anything unknown.
      if (kind === "max") return stats.max;
      return stats.min;
  }
  void value;
}

// Where a value sits between the first and last boundary, 0 to 1.
function barPosition(values, value, stats) {
  if (!values || values.length < 2) return null;
  const low = thresholdOf(values[0], stats, value, "min");
  const high = thresholdOf(values[values.length - 1], stats, value, "max");
  if (low == null || high == null || high === low) return 0;
  return Math.max(0, Math.min(1, (value - low) / (high - low)));
}

// The colour a three colour scale picks, interpolating the middle stop.
function scaleColour(scale, value, stats, resolve) {
  if (!scale || !scale.colors || !scale.colors.length) return null;
  const stops = scale.values && scale.values.length === scale.colors.length
    ? scale.values
    : scale.colors.map((color, index) => ({
      type: index === 0 ? "min" : index === scale.colors.length - 1 ? "max" : "percentile",
      val: String((index / Math.max(1, scale.colors.length - 1)) * 100),
    }));
  const points = stops.map((stop) => thresholdOf(stop, stats, value, "min"));
  const first = points[0];
  const last = points[points.length - 1];
  if (first == null || last == null) return resolve(scale.colors[0]);
  if (last === first) return resolve(scale.colors[scale.colors.length - 1]);
  const position = Math.max(0, Math.min(1, (value - first) / (last - first)));
  const scaled = position * (points.length - 1);
  const index = Math.min(points.length - 2, Math.floor(scaled));
  const local = scaled - index;
  const from = resolve(scale.colors[index]);
  const to = resolve(scale.colors[index + 1]);
  return mixColour(from, to, local);
}

// Blends two #rrggbb colours in sRGB, which is what Excel's scales look like
// on screen for the stops Office writes.
function mixColour(from, to, t) {
  const a = parseHexColour(from);
  const b = parseHexColour(to);
  if (!a || !b) return from;
  const mix = (x, y) => Math.round(x + (y - x) * t);
  return "#" + [mix(a[0], b[0]), mix(a[1], b[1]), mix(a[2], b[2])]
    .map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, "0"))
    .join("");
}

function parseHexColour(value) {
  if (typeof value !== "string") return null;
  const match = /^#?([0-9a-f]{6})$/i.exec(value.trim());
  if (!match) return null;
  const n = parseInt(match[1], 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

// Which icon bucket a value falls in. The first cfvo is the floor and each
// later one is a boundary; reverse flips the set as Excel does.
function iconFor(iconSet, value, stats) {
  const spec = iconSetSpec(iconSet.set);
  const count = spec.glyphs.length;
  const values = iconSet.values || [];
  let index = 0;
  for (let i = 1; i < count && i < values.length; i++) {
    const boundary = thresholdOf(values[i], stats, value, "min");
    if (boundary == null) continue;
    if (value >= boundary) index = i;
  }
  if (iconSet.reverse) index = count - 1 - index;
  return { glyph: spec.glyphs[index], color: spec.colors[index] };
}

// Does this rule fire for the value? Statistics come from the block.
function ruleMatches(rule, value, text, stats) {
  switch (rule.type) {
    case "cellIs": {
      const first = parseFloat(rule.formulas && rule.formulas[0]);
      if (isNaN(first)) return false;
      switch (rule.operator) {
        case "greaterThan": return value > first;
        case "greaterThanOrEqual": return value >= first;
        case "lessThan": return value < first;
        case "lessThanOrEqual": return value <= first;
        case "equal": return value === first;
        case "notEqual": return value !== first;
        case "between": {
          const second = parseFloat(rule.formulas[1]);
          return value >= Math.min(first, second) && value <= Math.max(first, second);
        }
        case "notBetween": {
          const second = parseFloat(rule.formulas[1]);
          return value < Math.min(first, second) || value > Math.max(first, second);
        }
        default: return false;
      }
    }
    case "containsText":
      return rule.text != null && String(text).toLowerCase().indexOf(String(rule.text).toLowerCase()) !== -1;
    case "notContainsText":
      return rule.text != null && String(text).toLowerCase().indexOf(String(rule.text).toLowerCase()) === -1;
    case "beginsWith":
      return rule.text != null && String(text).toLowerCase().startsWith(String(rule.text).toLowerCase());
    case "endsWith":
      return rule.text != null && String(text).toLowerCase().endsWith(String(rule.text).toLowerCase());
    case "containsBlanks":
      return String(text) === "";
    case "notContainsBlanks":
      return String(text) !== "";
    case "top10": {
      if (!stats.count) return false;
      let wanted = rule.rank && rule.rank > 0 ? rule.rank : 10;
      if (rule.percent) wanted = Math.max(1, Math.ceil((wanted / 100) * stats.count));
      wanted = Math.min(wanted, stats.count);
      if (rule.bottom) return value <= stats.values[wanted - 1];
      return value >= stats.values[stats.count - wanted];
    }
    case "aboveAverage":
      return stats.count > 0 && value > stats.average;
    case "belowAverage":
      return stats.count > 0 && value < stats.average;
    case "duplicateValues":
      return stats.count > 0 && countValue(stats.values, value) > 1;
    case "uniqueValues":
      return stats.count > 0 && countValue(stats.values, value) === 1;
    case "colorScale":
    case "dataBar":
    case "iconSet":
      return true;
    default:
      return false;
  }
}

function countValue(sorted, value) {
  let count = 0;
  for (const entry of sorted) {
    if (entry === value) count++;
    else if (entry > value) break;
  }
  return count;
}

// What a rule draws: a fill, a font colour, a data bar or an icon. The caller
// resolves colours and differential styles.
function overrideFor(rules, value, text, stats, helpers) {
  for (const rule of rules || []) {
    if (!ruleMatches(rule, value, text, stats)) continue;
    if (rule.type === "colorScale" && rule.colorScale) {
      const color = scaleColour(rule.colorScale, value, stats, helpers.resolveColour);
      if (color) return { fill: color };
    } else if (rule.type === "dataBar" && rule.dataBar) {
      const t = barPosition(rule.dataBar.values, value, stats);
      return { bar: { t: t == null ? 0 : t, color: helpers.resolveColour(rule.dataBar.color, "#638EC6") } };
    } else if (rule.type === "iconSet" && rule.iconSet) {
      return { icon: iconFor(rule.iconSet, value, stats), hideValue: rule.iconSet.showValue === false };
    } else if (rule.dxfId != null) {
      const dxf = helpers.dxfStyle(rule.dxfId);
      if (dxf) {
        return {
          fill: dxf.fill,
          fontColor: dxf.font && dxf.font.color ? dxf.font.color : null,
          bold: dxf.font ? dxf.font.bold : false,
          italic: dxf.font ? dxf.font.italic : false,
        };
      }
    }
    if (rule.stopIfTrue) return null;
  }
  return null;
}

// The rule that fires for a cell, with the block statistics it needs.
function overrideForCell(sheet, row, col, value, text, helpers) {
  if (!sheet.conditional || !sheet.conditional.length) return null;
  for (const block of sheet.conditional) {
    let inside = false;
    for (const range of block.ranges || []) {
      if (row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2) {
        inside = true;
        break;
      }
    }
    if (!inside) continue;
    const stats = blockStats(block, (r, c) => helpers.readValue(r, c));
    const result = overrideFor(block.rules, value, text, stats, helpers);
    if (result) return result;
  }
  return null;
}

module.exports = {
  ICON_SETS,
  iconSetSpec,
  blockStats,
  thresholdOf,
  barPosition,
  scaleColour,
  iconFor,
  ruleMatches,
  overrideFor,
  overrideForCell,
};
