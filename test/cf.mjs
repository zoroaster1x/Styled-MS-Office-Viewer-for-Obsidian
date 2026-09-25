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

// Conditional formatting arithmetic: thresholds, data bars, colour scales,
// icon sets and the rule types that need the range's own statistics (top ten,
// above average, duplicates). No files, no DOM.
//
//   bun test/cf.mjs

import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const cf = require("../src/spreadsheet/conditional.js");

let pass = 0;
let fail = 0;
function check(name, got, want) {
  if (got === want) {
    pass++;
    return;
  }
  fail++;
  console.log("FAIL", name, "->", JSON.stringify(got), "wanted", JSON.stringify(want));
}

// A block of five values read through a callback, as the renderer does.
const values = new Map([["0:0", 10], ["1:0", 20], ["2:0", 30], ["3:0", 40], ["4:0", 50]]);
const readValue = (r, c) => (values.has(r + ":" + c) ? values.get(r + ":" + c) : null);
const block = { ranges: [{ r1: 0, r2: 4, c1: 0, c2: 0 }], rules: [] };
const stats = cf.blockStats(block, readValue);
check("range count", stats.count, 5);
check("range min", stats.min, 10);
check("range max", stats.max, 50);
check("range average", stats.average, 30);
check("stats are memoised", cf.blockStats(block, () => 999).count, 5);

// Thresholds.
check("num threshold", cf.thresholdOf({ type: "num", val: "25" }, stats), 25);
check("min threshold", cf.thresholdOf({ type: "min" }, stats), 10);
check("max threshold", cf.thresholdOf({ type: "max" }, stats), 50);
check("percent threshold", cf.thresholdOf({ type: "percent", val: "50" }, stats), 30);
check("percentile threshold", cf.thresholdOf({ type: "percentile", val: "50" }, stats), 30);
check("autoMin is the range", cf.thresholdOf({ type: "autoMin" }, stats), 10);
check("autoMax is the range", cf.thresholdOf({ type: "autoMax" }, stats, 0, "max"), 50);

// Data bars: 30 sits halfway between 10 and 50.
check("bar position", cf.barPosition([{ type: "min" }, { type: "max" }], 30, stats), 0.5);
check("bar position clamps", cf.barPosition([{ type: "min" }, { type: "num", val: "30" }], 50, stats), 1);
check("bar on a flat range", cf.barPosition([{ type: "min" }, { type: "max" }], 7, { min: 7, max: 7, count: 1, values: [7] }), 0);

// Colour scales: white to red to green, and a two stop scale.
const resolve = (colour) => (typeof colour === "string" ? colour : "#000000");
const three = { values: [{ type: "min" }, { type: "percent", val: "50" }, { type: "max" }], colors: ["#ffffff", "#ffeb84", "#63be7b"] };
check("scale at the low stop", cf.scaleColour(three, 10, stats, resolve), "#ffffff");
check("scale at the middle stop", cf.scaleColour(three, 30, stats, resolve), "#ffeb84");
check("scale at the high stop", cf.scaleColour(three, 50, stats, resolve), "#63be7b");
check("scale interpolates", cf.scaleColour(three, 20, stats, resolve), "#fff5c2");
const two = { values: [{ type: "min" }, { type: "max" }], colors: ["#000000", "#ffffff"] };
check("two stop scale midpoint", cf.scaleColour(two, 30, stats, resolve), "#808080");

// Icon sets: three traffic lights on thirds of the range.
const icons = { set: "3TrafficLights1", values: [{ type: "percent", val: "0" }, { type: "percent", val: "33" }, { type: "percent", val: "67" }], reverse: false, showValue: true };
check("icon low bucket", cf.iconFor(icons, 10, stats).color, "#63BE7B");
check("icon middle bucket", cf.iconFor(icons, 30, stats).color, "#FFEB84");
check("icon high bucket", cf.iconFor(icons, 50, stats).color, "#F8696B");
check("icon reverse", cf.iconFor({ set: "3TrafficLights1", values: icons.values, reverse: true }, 50, stats).color, "#63BE7B");
check("icon set falls back", cf.iconFor({ set: "nonexistent", values: [] }, 10, stats).glyph, "\u25CF");
check("four rating has four steps", cf.iconSetSpec("4Rating").glyphs.length, 4);

// Rules that need the range.
const topTwo = { type: "top10", rank: 2 };
check("top 2 catches 50", cf.ruleMatches(topTwo, 50, "", stats), true);
check("top 2 catches 40", cf.ruleMatches(topTwo, 40, "", stats), true);
check("top 2 rejects 30", cf.ruleMatches(topTwo, 30, "", stats), false);
check("bottom 2 catches 10", cf.ruleMatches({ type: "top10", rank: 2, bottom: true }, 10, "", stats), true);
check("top 40 percent of five is two values", cf.ruleMatches({ type: "top10", rank: 40, percent: true }, 30, "", stats), false);
check("top 60 percent of five is three values", cf.ruleMatches({ type: "top10", rank: 60, percent: true }, 30, "", stats), true);
check("bottom 20 percent of five is one value", cf.ruleMatches({ type: "top10", rank: 20, percent: true, bottom: true }, 20, "", stats), false);
check("above average rejects 30", cf.ruleMatches({ type: "aboveAverage" }, 30, "", stats), false);
check("above average catches 40", cf.ruleMatches({ type: "aboveAverage" }, 40, "", stats), true);
check("below average catches 20", cf.ruleMatches({ type: "belowAverage" }, 20, "", stats), true);
check("duplicate needs a twin", cf.ruleMatches({ type: "duplicateValues" }, 40, "", stats), false);
const twins = { values: [1, 2, 2, 3], count: 4, min: 1, max: 3, average: 2 };
check("duplicate finds a twin", cf.ruleMatches({ type: "duplicateValues" }, 2, "", twins), true);
check("unique finds the odd one", cf.ruleMatches({ type: "uniqueValues" }, 3, "", twins), true);
check("unique rejects a twin", cf.ruleMatches({ type: "uniqueValues" }, 2, "", twins), false);
check("cellIs greaterThan", cf.ruleMatches({ type: "cellIs", operator: "greaterThan", formulas: ["35"] }, 40, "", stats), true);
check("cellIs between", cf.ruleMatches({ type: "cellIs", operator: "between", formulas: ["20", "40"] }, 30, "", stats), true);
check("containsText", cf.ruleMatches({ type: "containsText", text: "ab" }, 0, "aBc", stats), true);
check("endsWith", cf.ruleMatches({ type: "endsWith", text: "bc" }, 0, "abc", stats), true);
check("blanks", cf.ruleMatches({ type: "containsBlanks" }, 0, "", stats), true);

// What a rule draws.
const helpers = {
  resolveColour: (colour, fallback) => (typeof colour === "string" ? colour : fallback),
  dxfStyle: (id) => (id === 3 ? { fill: "#ffeeaa", font: { color: "#aa0000", bold: true } } : null),
  readValue: (row, col) => (values.has(row + ":" + col) ? values.get(row + ":" + col) : null),
};
const bar = cf.overrideFor([{ type: "dataBar", dataBar: { color: "#638EC6", values: [{ type: "min" }, { type: "max" }] } }], 30, "", stats, helpers);
check("override returns a bar", bar.bar.t, 0.5);
check("bar colour", bar.bar.color, "#638EC6");
const icon = cf.overrideFor([{ type: "iconSet", iconSet: icons }], 50, "", stats, helpers);
check("override returns an icon", icon.icon.color, "#F8696B");
check("icon keeps the value by default", icon.hideValue, false);
const hidden = cf.overrideFor([{ type: "iconSet", iconSet: { set: "3Symbols", values: icons.values, showValue: false } }], 50, "", stats, helpers);
check("showValue=0 hides the value", hidden.hideValue, true);
const dxf = cf.overrideFor([{ type: "cellIs", operator: "greaterThan", formulas: ["35"], dxfId: 3 }], 40, "", stats, helpers);
check("dxf fill", dxf.fill, "#ffeeaa");
check("dxf font", dxf.fontColor, "#aa0000");
check("dxf bold", dxf.bold, true);
check("stopIfTrue stops later rules", cf.overrideFor([
  { type: "cellIs", operator: "greaterThan", formulas: ["35"], stopIfTrue: true },
  { type: "colorScale", colorScale: two },
], 40, "", stats, helpers), null);

// The cell level entry point: range matching plus reading values.
const sheet = { conditional: [{ ranges: [{ r1: 0, r2: 4, c1: 0, c2: 0 }], rules: [{ type: "colorScale", colorScale: two }] }] };
const result = cf.overrideForCell(sheet, 2, 0, 30, "", helpers);
check("cell entry point resolves through the range", result.fill, "#808080");
check("outside the range is untouched", cf.overrideForCell(sheet, 9, 9, 30, "", helpers), null);
check("no conditional blocks means nothing", cf.overrideForCell({ conditional: [] }, 0, 0, 1, "", helpers), null);

console.log(`cf: ${pass} pass, ${fail} fail`);
if (fail) process.exit(1);
