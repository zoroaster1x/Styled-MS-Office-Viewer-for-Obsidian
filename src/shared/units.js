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

// Unit conversions and small numeric helpers shared by the renderers.

// OOXML measures font sizes in half points, EMU in 1/914400 inch, and lengths
// in twentieths of a point.
const EMU_PER_INCH = 914400;
const EMU_PER_PX = 9525;
const PT_PER_INCH = 72;
const TWIP_PER_PT = 20;

function emuToPx(emu) {
  return Number(emu || 0) / EMU_PER_PX;
}

function pxToEmu(px) {
  return Math.round(Number(px || 0) * EMU_PER_PX);
}

// half points -> px
function halfPtToPx(halfPt) {
  return (Number(halfPt || 0) / 2) * (96 / 72);
}

// points -> px
function ptToPx(pt) {
  return Number(pt || 0) * (96 / 72);
}

// twentieths of a point -> px
function twipToPx(twip) {
  return Number(twip || 0) / TWIP_PER_PT * (96 / 72);
}

// ODF lengths: "1.25cm", "12pt", "0.5in", "20px".
function odfLengthToPx(text) {
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(String(text || "").trim());
  if (!m) return null;
  const n = parseFloat(m[1]);
  if (isNaN(n)) return null;
  switch (m[2]) {
    case "cm": return n * 96 / 2.54;
    case "mm": return n * 96 / 25.4;
    case "in": return n * 96;
    case "pt": return n * 96 / 72;
    case "pc": return n * 16;
    case "px":
    default: return n;
  }
}

function clamp(v, lo, hi) {
  return v < lo ? lo : v > hi ? hi : v;
}

function round2(v) {
  return Math.round(Number(v) * 100) / 100;
}

// Percentages written as "62.5%" or "62500" (hundredths of a percent in OOXML).
function pctToFraction(text) {
  const s = String(text == null ? "" : text).trim();
  if (!s) return null;
  if (s.endsWith("%")) {
    const n = parseFloat(s);
    return isNaN(n) ? null : n / 100;
  }
  const n = parseFloat(s);
  if (isNaN(n)) return null;
  return n > 1 ? n / 100000 : n;
}

// Physical page size in px at 96dpi, from an OOXML page size in twips.
function pageSizePx(wTwips, hTwips) {
  return { w: twipToPx(wTwips), h: twipToPx(hTwips) };
}

module.exports = {
  EMU_PER_INCH,
  EMU_PER_PX,
  emuToPx,
  pxToEmu,
  halfPtToPx,
  ptToPx,
  twipToPx,
  odfLengthToPx,
  clamp,
  round2,
  pctToFraction,
  pageSizePx,
};
