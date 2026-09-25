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

/*
 * Preset shape geometry, the a:prstGeom silhouettes: arrows, callouts, stars,
 * flowchart symbols, brackets and the rest. The formulas follow the avLst
 * defaults in ECMA-376 presetShapeDefinitions.xml. Points and path data are in
 * local pixels for the shape box, which is what a CSS clip path wants.
 *
 * Adapted from the preset geometry module of Office Attachment Viewer
 * (maic-for-office) by Aqu-Lab, Apache License 2.0, with thanks. Changes: the
 * TypeScript types are gone, the module is CommonJS, and the connector helpers
 * are kept while the canvas renderer that used them is not part of this plugin.
 */
"use strict";
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __defNormalProp = (obj, key, value) => key in obj ? __defProp(obj, key, { enumerable: true, configurable: true, writable: true, value }) : obj[key] = value;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);
var __publicField = (obj, key, value) => __defNormalProp(obj, typeof key !== "symbol" ? key + "" : key, value);
var preset_geometry_exports = {};
__export(preset_geometry_exports, {
  connectorBezier: () => connectorBezier,
  connectorPoints: () => connectorPoints,
  isConnectorPreset: () => isConnectorPreset,
  isPillPreset: () => isPillPreset,
  presetPath: () => presetPath,
  presetPolygon: () => presetPolygon
});
module.exports = __toCommonJS(preset_geometry_exports);
const CONNECTOR_RE = /^(line|straightConnector\d?|bentConnector\d|curvedConnector\d)$/;
function isConnectorPreset(preset) {
  return !!preset && CONNECTOR_RE.test(preset);
}
function connectorPoints(preset, w, h, flipH, flipV, adjust) {
  let pts;
  if (/^bentConnector|^curvedConnector/.test(preset)) {
    pts = bentConnectorPts(preset, w, h, adjust);
  } else {
    pts = [0, 0, w, h];
  }
  if (flipH) for (let i = 0; i < pts.length; i += 2) pts[i] = w - pts[i];
  if (flipV) for (let i = 1; i < pts.length; i += 2) pts[i] = h - pts[i];
  return pts;
}
function bentConnectorPts(preset, w, h, adjust) {
  const n = parseInt(preset.slice(-1), 10) || 2;
  const clamp = (v, lo = -2, hi = 3) => Math.min(Math.max(v, lo), hi);
  const adj = (name, dflt) => adjust?.[name] != null ? clamp(adjust[name] / 1e5) : dflt;
  if (n <= 2) {
    return [0, 0, w, 0, w, h];
  } else if (n === 3) {
    const a1 = adj("adj1", 0.5);
    const mx = w * a1;
    return [0, 0, mx, 0, mx, h, w, h];
  } else if (n === 4) {
    const a1 = adj("adj1", 0.5);
    const a2 = adj("adj2", 0.5);
    const mx = w * a1;
    const my = h * a2;
    return [0, 0, mx, 0, mx, my, w, my, w, h];
  } else {
    const a1 = adj("adj1", 0.333);
    const a2 = adj("adj2", 0.5);
    const a3 = adj("adj3", 0.667);
    const x1 = w * a1;
    const y1 = h * a2;
    const x2 = w * a3;
    return [0, 0, x1, 0, x1, y1, x2, y1, x2, h, w, h];
  }
}
function connectorBezier(pts) {
  const nPts = pts.length / 2;
  if (nPts < 3) return [];
  const bezier = [];
  for (let i = 1; i < nPts; i++) {
    const x0 = pts[(i - 1) * 2];
    const y0 = pts[(i - 1) * 2 + 1];
    const x1 = pts[i * 2];
    const y1 = pts[i * 2 + 1];
    const prevX = i > 1 ? pts[(i - 2) * 2] : x0;
    const prevY = i > 1 ? pts[(i - 2) * 2 + 1] : y0;
    const nextX = i < nPts - 1 ? pts[(i + 1) * 2] : x1;
    const nextY = i < nPts - 1 ? pts[(i + 1) * 2 + 1] : y1;
    const cp1x = x0 + (x1 - prevX) / 6;
    const cp1y = y0 + (y1 - prevY) / 6;
    const cp2x = x1 - (nextX - x0) / 6;
    const cp2y = y1 - (nextY - y0) / 6;
    bezier.push(cp1x, cp1y, cp2x, cp2y, x1, y1);
  }
  return bezier;
}
function presetPolygon(preset, w, h, adjust) {
  if (!preset || w <= 0 || h <= 0) return null;
  const ss = Math.min(w, h);
  const frac = (name, dflt) => Math.min(Math.max((adjust?.[name] ?? dflt) / 1e5, 0), 1);
  switch (preset) {
    case "triangle": {
      const apex = w * frac("adj", 5e4);
      return [apex, 0, w, h, 0, h];
    }
    case "rtTriangle":
      return [0, 0, w, h, 0, h];
    case "diamond":
    case "flowChartDecision":
      return [w / 2, 0, w, h / 2, w / 2, h, 0, h / 2];
    case "parallelogram": {
      const inset = ss * frac("adj", 25e3);
      return [inset, 0, w, 0, w - inset, h, 0, h];
    }
    case "trapezoid": {
      const inset = ss * frac("adj", 25e3);
      return [inset, 0, w - inset, 0, w, h, 0, h];
    }
    case "pentagon": {
      return [w / 2, 0, w, h * 0.382, w * 0.809, h, w * 0.191, h, 0, h * 0.382];
    }
    case "hexagon": {
      const inset = ss * frac("adj", 25e3);
      return [inset, 0, w - inset, 0, w, h / 2, w - inset, h, inset, h, 0, h / 2];
    }
    case "octagon": {
      const c = ss * frac("adj", 29289);
      return [c, 0, w - c, 0, w, c, w, h - c, w - c, h, c, h, 0, h - c, 0, c];
    }
    case "mathPlus": {
      const t = ss * frac("adj1", 23520);
      const dx = w * 73490 / 2e5;
      const dy = h * 73490 / 2e5;
      const hc = w / 2;
      const vc = h / 2;
      const x1 = hc - dx;
      const x2 = hc - t;
      const x3 = hc + t;
      const x4 = hc + dx;
      const y1 = vc - dy;
      const y2 = vc - t;
      const y3 = vc + t;
      const y4 = vc + dy;
      return [x1, y2, x2, y2, x2, y1, x3, y1, x3, y2, x4, y2, x4, y3, x3, y3, x3, y4, x2, y4, x2, y3, x1, y3];
    }
    case "mathNotEqual": {
      const a1 = Math.min(frac("adj1", 23520), 0.5);
      const crAng = Math.min(Math.max(adjust?.adj2 ?? 66e5, 42e5), 66e5);
      const a3 = Math.min(frac("adj3", 11760), 1 - 2 * a1);
      const dy1 = h * a1;
      const dy2 = h * a3 / 2;
      const dx1 = w * 73490 / 2e5;
      const hc = w / 2;
      const vc = h / 2;
      const hd2 = h / 2;
      const x1 = hc - dx1;
      const x8 = hc + dx1;
      const y2 = vc - dy2;
      const y3 = vc + dy2;
      const y1 = y2 - dy1;
      const y4 = y3 + dy1;
      const cadj2 = (crAng - 54e5) / 6e4;
      const xadj2 = hd2 * Math.tan(cadj2 * D2R);
      const len = Math.hypot(xadj2, hd2);
      const bhw = len * dy1 / hd2;
      const x7 = hc + xadj2 - bhw / 2;
      const x6 = x7 - xadj2 * y1 / hd2;
      const x5 = x7 - xadj2 * y2 / hd2;
      const x4 = x7 - xadj2 * y3 / hd2;
      const x3 = x7 - xadj2 * y4 / hd2;
      const rx6 = x6 + bhw;
      const rx5 = x5 + bhw;
      const rx4 = x4 + bhw;
      const rx3 = x3 + bhw;
      const rx7 = x7 + bhw;
      const dx7 = dy1 * hd2 / len;
      const dy3 = dy1 * xadj2 / len;
      const rx = cadj2 > 0 ? x7 + dx7 : rx7;
      const lx = cadj2 > 0 ? x7 : rx7 - dx7;
      const ry = cadj2 > 0 ? dy3 : 0;
      const ly = cadj2 > 0 ? 0 : -dy3;
      const dlx = w - rx;
      const drx = w - lx;
      const dly = h - ry;
      const dry = h - ly;
      return [
        x1,
        y1,
        x6,
        y1,
        lx,
        ly,
        rx,
        ry,
        rx6,
        y1,
        x8,
        y1,
        x8,
        y2,
        rx5,
        y2,
        rx4,
        y3,
        x8,
        y3,
        x8,
        y4,
        rx3,
        y4,
        drx,
        dry,
        dlx,
        dly,
        x3,
        y4,
        x1,
        y4,
        x1,
        y3,
        x4,
        y3,
        x5,
        y2,
        x1,
        y2
      ];
    }
    case "plus": {
      const a = ss * frac("adj", 25e3);
      const x1 = a;
      const x2 = w - a;
      const y1 = a;
      const y2 = h - a;
      return [
        x1,
        0,
        x2,
        0,
        x2,
        y1,
        w,
        y1,
        w,
        y2,
        x2,
        y2,
        x2,
        h,
        x1,
        h,
        x1,
        y2,
        0,
        y2,
        0,
        y1,
        x1,
        y1
      ];
    }
    case "rightArrow": {
      const thick = h * frac("adj1", 5e4);
      const head = Math.min(w, ss * frac("adj2", 5e4));
      const y1 = (h - thick) / 2;
      const y2 = (h + thick) / 2;
      const xh = w - head;
      return [0, y1, xh, y1, xh, 0, w, h / 2, xh, h, xh, y2, 0, y2];
    }
    case "notchedRightArrow": {
      const thick = h * frac("adj1", 5e4);
      const head = Math.min(w, ss * frac("adj2", 5e4));
      const y1 = (h - thick) / 2;
      const y2 = (h + thick) / 2;
      const xh = w - head;
      const notch = head * thick / h;
      return [0, y1, xh, y1, xh, 0, w, h / 2, xh, h, xh, y2, 0, y2, notch, h / 2];
    }
    case "leftArrow": {
      const thick = h * frac("adj1", 5e4);
      const head = Math.min(w, ss * frac("adj2", 5e4));
      const y1 = (h - thick) / 2;
      const y2 = (h + thick) / 2;
      return [w, y1, head, y1, head, 0, 0, h / 2, head, h, head, y2, w, y2];
    }
    case "upArrow": {
      const thick = w * frac("adj1", 5e4);
      const head = Math.min(h, ss * frac("adj2", 5e4));
      const x1 = (w - thick) / 2;
      const x2 = (w + thick) / 2;
      return [x1, h, x1, head, 0, head, w / 2, 0, w, head, x2, head, x2, h];
    }
    case "downArrow": {
      const thick = w * frac("adj1", 5e4);
      const head = Math.min(h, ss * frac("adj2", 5e4));
      const x1 = (w - thick) / 2;
      const x2 = (w + thick) / 2;
      const yh = h - head;
      return [x1, 0, x1, yh, 0, yh, w / 2, h, w, yh, x2, yh, x2, 0];
    }
    // Arrow callouts: a callout box on one side with an arrow (shaft + head) growing
    // out of the opposite side (ECMA guide formulas; adj4 = box extent)
    case "upArrowCallout":
    case "downArrowCallout": {
      const fracU = (name, dflt) => Math.max((adjust?.[name] ?? dflt) / 1e5, 0);
      const a2 = Math.min(fracU("adj2", 25e3), 0.5 * w / ss);
      const a1 = Math.min(fracU("adj1", 25e3), a2 * 2);
      const a3 = Math.min(fracU("adj3", 25e3), h / ss);
      const a4 = Math.min(fracU("adj4", 64977), 1 - a3 * ss / h);
      const headHalf = ss * a2;
      const shaftHalf = ss * a1 / 2;
      const x1 = w / 2 - headHalf;
      const x2 = w / 2 - shaftHalf;
      const x3 = w / 2 + shaftHalf;
      const x4 = w / 2 + headHalf;
      if (preset === "upArrowCallout") {
        const y12 = ss * a3;
        const y22 = h - h * a4;
        return [0, y22, x2, y22, x2, y12, x1, y12, w / 2, 0, x4, y12, x3, y12, x3, y22, w, y22, w, h, 0, h];
      }
      const y1 = h - ss * a3;
      const y2 = h * a4;
      return [0, y2, x2, y2, x2, y1, x1, y1, w / 2, h, x4, y1, x3, y1, x3, y2, w, y2, w, 0, 0, 0];
    }
    case "leftArrowCallout":
    case "rightArrowCallout": {
      const fracU = (name, dflt) => Math.max((adjust?.[name] ?? dflt) / 1e5, 0);
      const a2 = Math.min(fracU("adj2", 25e3), 0.5 * h / ss);
      const a1 = Math.min(fracU("adj1", 25e3), a2 * 2);
      const a3 = Math.min(fracU("adj3", 25e3), w / ss);
      const a4 = Math.min(fracU("adj4", 64977), 1 - a3 * ss / w);
      const headHalf = ss * a2;
      const shaftHalf = ss * a1 / 2;
      const y1 = h / 2 - headHalf;
      const y2 = h / 2 - shaftHalf;
      const y3 = h / 2 + shaftHalf;
      const y4 = h / 2 + headHalf;
      if (preset === "leftArrowCallout") {
        const x12 = ss * a3;
        const x22 = w - w * a4;
        return [x22, 0, x22, y2, x12, y2, x12, y1, 0, h / 2, x12, y4, x12, y3, x22, y3, x22, h, w, h, w, 0];
      }
      const x1 = w - ss * a3;
      const x2 = w * a4;
      return [x2, 0, x2, y2, x1, y2, x1, y1, w, h / 2, x1, y4, x1, y3, x2, y3, x2, h, 0, h, 0, 0];
    }
    case "leftRightArrow": {
      const thick = h * frac("adj1", 5e4);
      const head = Math.min(w / 2, ss * frac("adj2", 5e4));
      const y1 = (h - thick) / 2;
      const y2 = (h + thick) / 2;
      return [
        0,
        h / 2,
        head,
        0,
        head,
        y1,
        w - head,
        y1,
        w - head,
        0,
        w,
        h / 2,
        w - head,
        h,
        w - head,
        y2,
        head,
        y2,
        head,
        h
      ];
    }
    case "upDownArrow": {
      const thick = w * frac("adj1", 5e4);
      const head = Math.min(h / 2, ss * frac("adj2", 5e4));
      const x1 = (w - thick) / 2;
      const x2 = (w + thick) / 2;
      return [
        w / 2,
        0,
        w,
        head,
        x2,
        head,
        x2,
        h - head,
        w,
        h - head,
        w / 2,
        h,
        0,
        h - head,
        x1,
        h - head,
        x1,
        head,
        0,
        head
      ];
    }
    case "chevron": {
      const d = ss * frac("adj", 5e4);
      return [0, 0, w - d, 0, w, h / 2, w - d, h, 0, h, d, h / 2];
    }
    case "homePlate": {
      const d = ss * frac("adj", 5e4);
      return [0, 0, w - d, 0, w, h / 2, w - d, h, 0, h];
    }
    case "snip1Rect": {
      const a = ss * frac("adj", 16667);
      return [0, 0, w - a, 0, w, a, w, h, 0, h];
    }
    case "snip2SameRect": {
      const a1 = ss * frac("adj1", 16667);
      const a2 = ss * frac("adj2", 0);
      return [a1, 0, w - a1, 0, w, a1, w, h - a2, w - a2, h, a2, h, 0, h - a2, 0, a1];
    }
    case "snip2DiagRect": {
      const a1 = ss * frac("adj1", 0);
      const a2 = ss * frac("adj2", 16667);
      return [a1, 0, w - a2, 0, w, a2, w, h - a1, w - a1, h, a2, h, 0, h - a2, 0, a1];
    }
    case "halfFrame": {
      const y1 = ss * frac("adj1", 33333);
      const x1 = ss * frac("adj2", 33333);
      const x2 = Math.max(w - y1 * w / h, x1);
      const y2 = Math.max(h - x1 * h / w, y1);
      return [0, 0, w, 0, x2, y1, x1, y1, x1, y2, 0, h];
    }
    case "corner": {
      const y1 = ss * frac("adj1", 5e4);
      const x1 = ss * frac("adj2", 5e4);
      return [0, 0, x1, 0, x1, h - y1, w, h - y1, w, h, 0, h];
    }
    case "diagStripe": {
      const a = frac("adj", 5e4);
      return [0, h * a, w * a, 0, w, 0, 0, h];
    }
    case "lightningBolt": {
      const u = [
        8472,
        0,
        12860,
        6672,
        11050,
        6672,
        16577,
        12007,
        14767,
        12007,
        21600,
        21600,
        10800,
        14387,
        12377,
        14387,
        5333,
        6667,
        7778,
        6667
      ];
      return u.map((v, i) => v / 21600 * (i % 2 === 0 ? w : h));
    }
    case "flowChartPreparation":
      return [w * 0.2, 0, w * 0.8, 0, w, h / 2, w * 0.8, h, w * 0.2, h, 0, h / 2];
    case "flowChartManualInput":
      return [0, h / 5, w, 0, w, h, 0, h];
    case "flowChartManualOperation":
      return [0, 0, w, 0, w * 0.8, h, w * 0.2, h];
    case "flowChartOffpageConnector":
      return [0, 0, w, 0, w, h * 0.8, w / 2, h, 0, h * 0.8];
    case "flowChartExtract":
      return [w / 2, 0, w, h, 0, h];
    case "flowChartMerge":
      return [0, 0, w, 0, w / 2, h];
    case "flowChartCollate":
      return [0, 0, w, 0, w / 2, h / 2, w, h, 0, h, w / 2, h / 2];
    case "gear6": {
      const depth = Math.min(frac("adj1", 15e3) * 2, 0.6);
      return gearPoints(6, w, h, 1 - depth);
    }
    case "gear9": {
      const depth = Math.min(frac("adj1", 1e4) * 2, 0.6);
      return gearPoints(9, w, h, 1 - depth);
    }
    case "quadArrow": {
      const sw2 = ss * frac("adj1", 22500) / 2;
      const hw = ss * frac("adj2", 22500);
      const hl = ss * frac("adj3", 22500);
      const cx = w / 2;
      const cy = h / 2;
      return [
        cx,
        0,
        cx + hw,
        hl,
        cx + sw2,
        hl,
        cx + sw2,
        cy - sw2,
        w - hl,
        cy - sw2,
        w - hl,
        cy - hw,
        w,
        cy,
        w - hl,
        cy + hw,
        w - hl,
        cy + sw2,
        cx + sw2,
        cy + sw2,
        cx + sw2,
        h - hl,
        cx + hw,
        h - hl,
        cx,
        h,
        cx - hw,
        h - hl,
        cx - sw2,
        h - hl,
        cx - sw2,
        cy + sw2,
        hl,
        cy + sw2,
        hl,
        cy + hw,
        0,
        cy,
        hl,
        cy - hw,
        hl,
        cy - sw2,
        cx - sw2,
        cy - sw2,
        cx - sw2,
        hl,
        cx - hw,
        hl
      ];
    }
    case "bentArrow": {
      const t = ss * frac("adj1", 25e3);
      const hw = ss * frac("adj2", 25e3);
      const hl = ss * frac("adj3", 25e3);
      const yc = Math.max(hw, t / 2);
      return [
        0,
        h,
        0,
        yc - t / 2,
        w - hl,
        yc - t / 2,
        w - hl,
        yc - hw,
        w,
        yc,
        w - hl,
        yc + hw,
        w - hl,
        yc + t / 2,
        t,
        yc + t / 2,
        t,
        h
      ];
    }
    case "wedgeRectCallout": {
      const tipX = w / 2 + w * adjRaw(adjust, "adj1", -20833);
      const tipY = h / 2 + h * adjRaw(adjust, "adj2", 62500);
      return wedgeCalloutPolygon(w, h, tipX, tipY);
    }
    case "irregularSeal1":
      return sealPoints(IRREGULAR_SEAL_1, w, h);
    case "irregularSeal2":
      return sealPoints(IRREGULAR_SEAL_2, w, h);
    case "star4":
      return starPoints(4, w, h, frac("adj", 12500));
    case "star5":
      return starPoints(5, w, h, frac("adj", 19098));
    case "star6":
      return starPoints(6, w, h, frac("adj", 28868));
    case "star7":
      return starPoints(7, w, h, frac("adj", 34601));
    case "star8":
      return starPoints(8, w, h, frac("adj", 37500));
    case "star10":
      return starPoints(10, w, h, frac("adj", 42533));
    case "star12":
      return starPoints(12, w, h, frac("adj", 37500));
    case "star16":
      return starPoints(16, w, h, frac("adj", 37500));
    case "star24":
      return starPoints(24, w, h, frac("adj", 37500));
    case "star32":
      return starPoints(32, w, h, frac("adj", 37500));
    default:
      return null;
  }
}
function adjRaw(adjust, name, dflt) {
  const v = (adjust?.[name] ?? dflt) / 1e5;
  return Math.min(Math.max(v, -2), 2);
}
function wedgeCalloutPolygon(w, h, tipX, tipY) {
  const g = Math.min(w, h) * 0.1;
  const clamp = (v, lo, hi) => Math.min(Math.max(v, lo), hi);
  const nx = (tipX - w / 2) / w;
  const ny = (tipY - h / 2) / h;
  if (Math.abs(ny) >= Math.abs(nx)) {
    const bx = clamp(tipX, 2 * g, w - 2 * g);
    if (ny >= 0) return [0, 0, w, 0, w, h, bx + g, h, tipX, tipY, bx - g, h, 0, h];
    return [0, 0, bx - g, 0, tipX, tipY, bx + g, 0, w, 0, w, h, 0, h];
  }
  const by = clamp(tipY, 2 * g, h - 2 * g);
  if (nx >= 0) return [0, 0, w, 0, w, by - g, tipX, tipY, w, by + g, w, h, 0, h];
  return [0, 0, w, 0, w, h, 0, h, 0, by + g, tipX, tipY, 0, by - g];
}
function gearPoints(teeth, w, h, innerR) {
  const cx = w / 2;
  const cy = h / 2;
  const pitch = 360 / teeth;
  const tipHalf = pitch * 0.16;
  const rootHalf = pitch * 0.38;
  const pts = [];
  for (let i = 0; i < teeth; i++) {
    const c = -90 + i * pitch;
    for (const [off, r] of [
      [-rootHalf, innerR],
      [-tipHalf, 1],
      [tipHalf, 1],
      [rootHalf, innerR]
    ]) {
      const a = (c + off) * Math.PI / 180;
      pts.push(cx + Math.cos(a) * cx * r, cy + Math.sin(a) * cy * r);
    }
  }
  return pts;
}
const IRREGULAR_SEAL_1 = [
  10800,
  5800,
  14522,
  0,
  14155,
  5325,
  18380,
  4457,
  16702,
  7315,
  21097,
  8137,
  17607,
  10475,
  21600,
  13290,
  16837,
  12942,
  18145,
  18095,
  14020,
  14457,
  13247,
  19737,
  10532,
  14935,
  8485,
  21600,
  7715,
  15627,
  4762,
  17617,
  5667,
  13937,
  135,
  14587,
  3722,
  11775,
  0,
  8837,
  4627,
  7992,
  2777,
  4912,
  7772,
  6142,
  8485,
  1017
];
const IRREGULAR_SEAL_2 = [
  11462,
  4342,
  14790,
  0,
  14525,
  5777,
  18007,
  3172,
  16380,
  6532,
  21600,
  6645,
  16985,
  9402,
  18270,
  11290,
  16380,
  12310,
  18877,
  15632,
  14640,
  14350,
  14942,
  17370,
  12180,
  15935,
  11612,
  18842,
  9872,
  17370,
  8700,
  19712,
  7527,
  18125,
  4917,
  21600,
  4805,
  18240,
  1285,
  17825,
  3330,
  15370,
  0,
  12877,
  3935,
  11592,
  1172,
  8270,
  5372,
  7817,
  4502,
  3625,
  8550,
  6382,
  9722,
  1887
];
function sealPoints(coords21600, w, h) {
  const pts = [];
  for (let i = 0; i < coords21600.length; i += 2) {
    pts.push(coords21600[i] / 21600 * w, coords21600[i + 1] / 21600 * h);
  }
  return pts;
}
function starPoints(n, w, h, innerFrac) {
  const cx = w / 2;
  const cy = h / 2;
  const pts = [];
  for (let i = 0; i < n * 2; i++) {
    const ang = -Math.PI / 2 + i * Math.PI / n;
    const f = i % 2 === 0 ? 1 : innerFrac * 2;
    const fr = Math.min(f, 1);
    pts.push(cx + Math.cos(ang) * cx * fr, cy + Math.sin(ang) * cy * fr);
  }
  return pts;
}
function isPillPreset(preset) {
  return preset === "flowChartTerminator" || preset === "flowChartAlternateProcess";
}
const R2 = (v) => Math.round(v * 100) / 100;
const D2R = Math.PI / 180;
class PathB {
  constructor() {
    __publicField(this, "parts", []);
  }
  M(x, y) {
    this.parts.push(`M ${R2(x)} ${R2(y)}`);
    return this;
  }
  L(x, y) {
    this.parts.push(`L ${R2(x)} ${R2(y)}`);
    return this;
  }
  Q(x1, y1, x, y) {
    this.parts.push(`Q ${R2(x1)} ${R2(y1)} ${R2(x)} ${R2(y)}`);
    return this;
  }
  C(x1, y1, x2, y2, x, y) {
    this.parts.push(`C ${R2(x1)} ${R2(y1)} ${R2(x2)} ${R2(y2)} ${R2(x)} ${R2(y)}`);
    return this;
  }
  Z() {
    this.parts.push("Z");
    return this;
  }
  /** Parametric angles (degrees, y-down clockwise positive); move says whether to M/L to the arc start first */
  arc(cx, cy, rx, ry, startDeg, sweepDeg, move) {
    const st = startDeg * D2R;
    const sw = sweepDeg * D2R;
    const sx = cx + rx * Math.cos(st);
    const sy = cy + ry * Math.sin(st);
    if (move === "M") this.M(sx, sy);
    else if (move === "L") this.L(sx, sy);
    if (sw === 0) return this;
    const segs = Math.max(1, Math.ceil(Math.abs(sw) / (Math.PI / 2)));
    const da = sw / segs;
    const k = 4 / 3 * Math.tan(da / 4);
    for (let i = 0; i < segs; i++) {
      const a1 = st + i * da;
      const a2 = a1 + da;
      const x1 = cx + rx * Math.cos(a1);
      const y1 = cy + ry * Math.sin(a1);
      const x2 = cx + rx * Math.cos(a2);
      const y2 = cy + ry * Math.sin(a2);
      this.C(
        x1 - k * rx * Math.sin(a1),
        y1 + k * ry * Math.cos(a1),
        x2 + k * rx * Math.sin(a2),
        y2 - k * ry * Math.cos(a2),
        x2,
        y2
      );
    }
    return this;
  }
  d() {
    return this.parts.join(" ");
  }
}
function ellipseSub(b, cx, cy, rx, ry, ccw = false) {
  b.arc(cx, cy, rx, ry, 0, ccw ? -360 : 360, "M").Z();
}
function mixedCornerRect(w, h, sizes, kinds) {
  const [tl, tr, br, bl] = sizes;
  const b = new PathB();
  b.M(tl, 0).L(w - tr, 0);
  if (kinds[1] === "round") b.arc(w - tr, tr, tr, tr, 270, 90);
  else if (kinds[1] === "snip") b.L(w, tr);
  b.L(w, h - br);
  if (kinds[2] === "round") b.arc(w - br, h - br, br, br, 0, 90);
  else if (kinds[2] === "snip") b.L(w - br, h);
  b.L(bl, h);
  if (kinds[3] === "round") b.arc(bl, h - bl, bl, bl, 90, 90);
  else if (kinds[3] === "snip") b.L(0, h - bl);
  b.L(0, tl);
  if (kinds[0] === "round") b.arc(tl, tl, tl, tl, 180, 90);
  else if (kinds[0] === "snip") b.L(tl, 0);
  return b.Z().d();
}
function cloudBlob(w, h) {
  const b = new PathB();
  const u = [
    [0.2, 0.85],
    [0.06, 0.86],
    [0, 0.72],
    [0.02, 0.59],
    [0.03, 0.47],
    [0.11, 0.39],
    [0.2, 0.42],
    [0.19, 0.26],
    [0.29, 0.14],
    [0.4, 0.2],
    [0.45, 0.07],
    [0.6, 0.04],
    [0.67, 0.14],
    [0.76, 0.04],
    [0.91, 0.1],
    [0.92, 0.26],
    [0.99, 0.31],
    [1, 0.46],
    [0.97, 0.56],
    [1, 0.69],
    [0.94, 0.81],
    [0.85, 0.82],
    [0.83, 0.94],
    [0.72, 0.98],
    [0.65, 0.91],
    [0.58, 1],
    [0.45, 1],
    [0.39, 0.91],
    [0.33, 0.98],
    [0.23, 0.95],
    [0.2, 0.85]
  ];
  b.M(u[0][0] * w, u[0][1] * h);
  for (let i = 1; i + 2 < u.length + 1; i += 3) {
    b.C(
      u[i][0] * w,
      u[i][1] * h,
      u[i + 1][0] * w,
      u[i + 1][1] * h,
      u[i + 2][0] * w,
      u[i + 2][1] * h
    );
  }
  return b.Z();
}
function presetPath(preset, w, h, adjust) {
  if (!preset || w <= 0 || h <= 0) return null;
  const ss = Math.min(w, h);
  const cx = w / 2;
  const cy = h / 2;
  const frac = (name, dflt) => Math.min(Math.max((adjust?.[name] ?? dflt) / 1e5, 0), 1);
  const ang = (name, dflt) => (adjust?.[name] ?? dflt) / 6e4;
  const sweepCW = (a1, a2) => ((a2 - a1) % 360 + 360) % 360;
  switch (preset) {
    case "arc": {
      const a1 = ang("adj1", 162e5);
      const a2 = ang("adj2", 0);
      const sw = sweepCW(a1, a2) || 90;
      const fill = new PathB().M(cx, cy);
      fill.arc(cx, cy, cx, cy, a1, sw, "L").Z();
      const stroke = new PathB().arc(cx, cy, cx, cy, a1, sw, "M");
      return { fillPath: fill.d(), strokePath: stroke.d() };
    }
    case "chord": {
      const a1 = ang("adj1", 27e5);
      const a2 = ang("adj2", 162e5);
      return {
        path: new PathB().arc(cx, cy, cx, cy, a1, sweepCW(a1, a2) || 180, "M").Z().d()
      };
    }
    case "pie": {
      const a1 = ang("adj1", 0);
      const a2 = ang("adj2", 162e5);
      return {
        path: new PathB().M(cx, cy).arc(cx, cy, cx, cy, a1, sweepCW(a1, a2) || 270, "L").Z().d()
      };
    }
    case "blockArc": {
      const a1 = ang("adj1", 108e5);
      const a2 = ang("adj2", 0);
      const sw = sweepCW(a1, a2) || 180;
      const t = ss * frac("adj3", 25e3);
      const rxI = Math.max(cx - t, 0);
      const ryI = Math.max(cy - t, 0);
      const b = new PathB().arc(cx, cy, cx, cy, a1, sw, "M");
      b.arc(cx, cy, rxI, ryI, a1 + sw, -sw, "L").Z();
      return { path: b.d() };
    }
    case "circularArrow":
    case "leftCircularArrow": {
      const left = preset === "leftCircularArrow";
      const t = ss * frac("adj1", 12500);
      const headSpan = Math.min(ang("adj2", 1142319), 90);
      const end = ang("adj3", left ? 1142319 : 20457681);
      const start = ang("adj4", 108e5);
      const headR = Math.max(ss * frac("adj5", 12500), t * 0.75);
      const rxO = Math.max(cx - headR, 0);
      const ryO = Math.max(cy - headR, 0);
      const rxI = Math.max(rxO - t, 0);
      const ryI = Math.max(ryO - t, 0);
      const rxM = (rxO + rxI) / 2;
      const ryM = (ryO + ryI) / 2;
      const total = left ? -(sweepCW(end, start) || 250) : sweepCW(start, end) || 250;
      const headAng = left ? -headSpan : headSpan;
      const sweep = total - headAng;
      const bodyEnd = start + sweep;
      const pt = (deg, rx, ry) => [cx + rx * Math.cos(deg * D2R), cy + ry * Math.sin(deg * D2R)];
      const b = new PathB().arc(cx, cy, rxO, ryO, start, sweep, "M");
      const [hx1, hy1] = pt(bodyEnd, rxM + headR, ryM + headR);
      const [tx, ty] = pt(bodyEnd + headAng, rxM, ryM);
      const [hx2, hy2] = pt(bodyEnd, Math.max(rxM - headR, 0), Math.max(ryM - headR, 0));
      b.L(hx1, hy1).L(tx, ty).L(hx2, hy2);
      b.arc(cx, cy, rxI, ryI, bodyEnd, -sweep, "L").Z();
      return { path: b.d() };
    }
    case "donut": {
      const t = ss * frac("adj", 25e3);
      const b = new PathB();
      ellipseSub(b, cx, cy, cx, cy);
      ellipseSub(b, cx, cy, Math.max(cx - t, 0), Math.max(cy - t, 0), true);
      return { path: b.d() };
    }
    case "frame": {
      const t = ss * frac("adj1", 12500);
      const b = new PathB().M(0, 0).L(w, 0).L(w, h).L(0, h).Z();
      b.M(t, t).L(t, h - t).L(w - t, h - t).L(w - t, t).Z();
      return { path: b.d() };
    }
    case "round1Rect": {
      const r = ss * frac("adj", 16667);
      return { path: mixedCornerRect(w, h, [0, r, 0, 0], ["none", "round", "none", "none"]) };
    }
    case "round2SameRect": {
      const r1 = ss * frac("adj1", 16667);
      const r2 = ss * frac("adj2", 0);
      return { path: mixedCornerRect(w, h, [r1, r1, r2, r2], ["round", "round", "round", "round"]) };
    }
    case "round2DiagRect": {
      const r1 = ss * frac("adj1", 16667);
      const r2 = ss * frac("adj2", 0);
      return { path: mixedCornerRect(w, h, [r1, r2, r1, r2], ["round", "round", "round", "round"]) };
    }
    case "snipRoundRect": {
      const r1 = ss * frac("adj1", 16667);
      const r2 = ss * frac("adj2", 16667);
      return { path: mixedCornerRect(w, h, [r1, r2, 0, 0], ["round", "snip", "none", "none"]) };
    }
    case "mathEqual": {
      const a1 = Math.min(frac("adj1", 23520), 0.36745);
      const a2 = Math.min(frac("adj2", 11760), 1 - 2 * a1);
      const dy1 = h * a1;
      const dy2 = h * a2 / 2;
      const dx1 = w * 73490 / 2e5;
      const x1 = cx - dx1;
      const x2 = cx + dx1;
      const y2 = cy - dy2;
      const y3 = cy + dy2;
      const y1 = y2 - dy1;
      const y4 = y3 + dy1;
      const b = new PathB().M(x1, y1).L(x2, y1).L(x2, y2).L(x1, y2).Z();
      b.M(x1, y3).L(x2, y3).L(x2, y4).L(x1, y4).Z();
      return { path: b.d() };
    }
    case "funnel": {
      const d = ss / 20;
      const hd4 = h / 4;
      const rw3 = cx / 4;
      const rh3 = hd4 / 4;
      const par = (deg) => Math.atan2(cx * Math.sin(deg * D2R), hd4 * Math.cos(deg * D2R)) / D2R;
      const da = Math.atan2(hd4 * Math.sin(8 * D2R), cx * Math.cos(8 * D2R)) / D2R;
      const t0 = par(180 - da);
      const t3 = par(da);
      const b = new PathB();
      b.arc(cx, hd4, cx, hd4, t0, t3 + 360 - t0, "M");
      b.arc(cx, h - rh3, rw3, rh3, t3, par(180 - da) - t3, "L").Z();
      b.arc(cx, hd4, Math.max(cx - d, 0), Math.max(hd4 - d, 0), 180, -360, "M").Z();
      return { path: b.d() };
    }
    case "heart": {
      const b = new PathB().M(0.5 * w, 0.3 * h);
      b.C(0.5 * w, 0.12 * h, 0.36 * w, 0.01 * h, 0.22 * w, 0.01 * h);
      b.C(0.06 * w, 0.01 * h, 0, 0.15 * h, 0, 0.28 * h);
      b.C(0, 0.5 * h, 0.2 * w, 0.65 * h, 0.5 * w, h);
      b.C(0.8 * w, 0.65 * h, w, 0.5 * h, w, 0.28 * h);
      b.C(w, 0.15 * h, 0.94 * w, 0.01 * h, 0.78 * w, 0.01 * h);
      b.C(0.64 * w, 0.01 * h, 0.5 * w, 0.12 * h, 0.5 * w, 0.3 * h);
      return { path: b.Z().d() };
    }
    case "moon": {
      const g = frac("adj", 5e4);
      const b = new PathB().arc(w, cy, w, cy, 270, -180, "M");
      b.arc(w, cy, w * (1 - g), cy, 90, 180);
      return { path: b.Z().d() };
    }
    case "sun": {
      const g = frac("adj", 25e3);
      const rx = w * g;
      const ry = h * g;
      const b = new PathB();
      for (let k = 0; k < 8; k++) {
        const a = k * 45 * D2R;
        const tipX = cx + cx * Math.cos(a);
        const tipY = cy + cy * Math.sin(a);
        const br = 1.35;
        const a1 = a - 12 * D2R;
        const a2 = a + 12 * D2R;
        b.M(cx + rx * br * Math.cos(a1), cy + ry * br * Math.sin(a1)).L(tipX, tipY).L(cx + rx * br * Math.cos(a2), cy + ry * br * Math.sin(a2)).Z();
      }
      ellipseSub(b, cx, cy, rx, ry);
      return { path: b.d() };
    }
    case "cloud":
      return { path: cloudBlob(w, h).d() };
    case "cloudCallout": {
      const tipX = cx + w * adjRaw(adjust, "adj1", -20833);
      const tipY = cy + h * adjRaw(adjust, "adj2", 62500);
      const b = cloudBlob(w, h);
      for (const [t, r] of [
        [0.72, 0.075],
        [0.92, 0.045]
      ]) {
        ellipseSub(b, cx + (tipX - cx) * t, cy + (tipY - cy) * t, ss * r, ss * r);
      }
      return { path: b.d() };
    }
    case "teardrop": {
      const a = Math.min(Math.max((adjust?.adj ?? 1e5) / 1e5, 0), 2);
      const tipX = cx + cx * a;
      const tipY = cy - cy * a;
      const b = new PathB().arc(cx, cy, cx, cy, 0, 270, "M");
      b.Q(cx + (tipX - cx) / 2, tipY, tipX, tipY).Q(w, (tipY + cy) / 2, w, cy);
      return { path: b.Z().d() };
    }
    case "plaque": {
      const r = ss * frac("adj", 16667);
      const b = new PathB().M(r, 0).L(w - r, 0);
      b.arc(w, 0, r, r, 180, -90).L(w, h - r);
      b.arc(w, h, r, r, 270, -90).L(r, h);
      b.arc(0, h, r, r, 0, -90).L(0, r);
      b.arc(0, 0, r, r, 90, -90);
      return { path: b.Z().d() };
    }
    case "cube": {
      const d = ss * frac("adj", 25e3);
      const path = new PathB().M(0, d).L(d, 0).L(w, 0).L(w, h - d).L(w - d, h).L(0, h).Z().d();
      const inner = new PathB().M(0, d).L(w - d, d).L(w, 0).M(w - d, d).L(w - d, h).d();
      return { path, strokePath: inner };
    }
    case "can": {
      const ry = h * frac("adj", 25e3) / 2;
      const b = new PathB().M(0, ry).L(0, h - ry);
      b.arc(cx, h - ry, cx, ry, 180, -180).L(w, ry);
      b.arc(cx, ry, cx, ry, 0, -180).Z();
      const rim = new PathB().arc(cx, ry, cx, ry, 180, -180, "M").d();
      return { path: b.d(), strokePath: rim };
    }
    case "flowChartMagneticDisk": {
      const ry = h / 6;
      const b = new PathB().M(0, ry).L(0, h - ry);
      b.arc(cx, h - ry, cx, ry, 180, -180).L(w, ry);
      b.arc(cx, ry, cx, ry, 0, -180).Z();
      const rim = new PathB().arc(cx, ry, cx, ry, 180, -180, "M").d();
      return { path: b.d(), strokePath: rim };
    }
    case "flowChartMagneticDrum": {
      const rx = w / 6;
      const b = new PathB().arc(w - rx, cy, rx, cy, 270, 180, "M").L(rx, h);
      b.arc(rx, cy, rx, cy, 90, 180).Z();
      const rim = new PathB().arc(w - rx, cy, rx, cy, 270, -180, "M").d();
      return { path: b.d(), strokePath: rim };
    }
    case "bevel": {
      const t = ss * frac("adj", 12500);
      const path = new PathB().M(0, 0).L(w, 0).L(w, h).L(0, h).Z().d();
      const inner = new PathB().M(t, t).L(w - t, t).L(w - t, h - t).L(t, h - t).Z().M(0, 0).L(t, t).M(w, 0).L(w - t, t).M(w, h).L(w - t, h - t).M(0, h).L(t, h - t);
      return { path, strokePath: inner.d() };
    }
    case "foldedCorner": {
      const f = ss * frac("adj", 16667);
      const path = new PathB().M(0, 0).L(w, 0).L(w, h - f).L(w - f, h).L(0, h).Z().d();
      const fold = new PathB().M(w - f, h).L(w - 0.8 * f, h - 0.8 * f).L(w, h - f).d();
      return { path, strokePath: fold };
    }
    case "smileyFace": {
      const b = new PathB();
      ellipseSub(b, cx, cy, cx, cy);
      const g = adjRaw(adjust, "adj", 4653);
      const face = new PathB();
      ellipseSub(face, 0.35 * w, 0.37 * h, 0.05 * w, 0.05 * h);
      ellipseSub(face, 0.65 * w, 0.37 * h, 0.05 * w, 0.05 * h);
      face.M(0.3 * w, 0.67 * h).Q(cx, h * Math.min(Math.max(0.67 + 4 * g, 0.4), 0.95), 0.7 * w, 0.67 * h);
      return { path: b.d(), strokePath: face.d() };
    }
    case "noSmoking": {
      const t = ss * frac("adj", 18750);
      const b = new PathB();
      ellipseSub(b, cx, cy, cx, cy);
      const rxI = Math.max(cx - t, 0);
      const ryI = Math.max(cy - t, 0);
      ellipseSub(b, cx, cy, rxI, ryI, true);
      const p1x = cx + rxI * Math.cos(225 * D2R);
      const p1y = cy + ryI * Math.sin(225 * D2R);
      const p2x = cx + rxI * Math.cos(45 * D2R);
      const p2y = cy + ryI * Math.sin(45 * D2R);
      const len = Math.hypot(p2x - p1x, p2y - p1y) || 1;
      const nx = -(p2y - p1y) / len * (t / 2);
      const ny = (p2x - p1x) / len * (t / 2);
      b.M(p1x + nx, p1y + ny).L(p2x + nx, p2y + ny).L(p2x - nx, p2y - ny).L(p1x - nx, p1y - ny).Z();
      return { path: b.d() };
    }
    case "ribbon": {
      const b = new PathB();
      b.M(0, 0.25 * h).L(0.25 * w, 0.25 * h).L(0.25 * w, h).L(0, h).L(0.0833 * w, 0.625 * h).Z();
      b.M(w, 0.25 * h).L(0.75 * w, 0.25 * h).L(0.75 * w, h).L(w, h).L(0.9167 * w, 0.625 * h).Z();
      b.M(0.125 * w, 0).L(0.875 * w, 0).L(0.875 * w, 0.75 * h).L(0.125 * w, 0.75 * h).Z();
      return { path: b.d() };
    }
    case "ribbon2": {
      const b = new PathB();
      b.M(0, 0.75 * h).L(0.25 * w, 0.75 * h).L(0.25 * w, 0).L(0, 0).L(0.0833 * w, 0.375 * h).Z();
      b.M(w, 0.75 * h).L(0.75 * w, 0.75 * h).L(0.75 * w, 0).L(w, 0).L(0.9167 * w, 0.375 * h).Z();
      b.M(0.125 * w, h).L(0.875 * w, h).L(0.875 * w, 0.25 * h).L(0.125 * w, 0.25 * h).Z();
      return { path: b.d() };
    }
    case "wave": {
      const a = h * Math.min(frac("adj1", 12500), 0.25);
      const b = new PathB().M(0, a);
      b.C(w / 6, 0, w / 3, 0, w / 2, a).C(2 * w / 3, 2 * a, 5 * w / 6, 2 * a, w, a);
      b.L(w, h - a);
      b.C(5 * w / 6, h, 2 * w / 3, h, w / 2, h - a).C(
        w / 3,
        h - 2 * a,
        w / 6,
        h - 2 * a,
        0,
        h - a
      );
      return { path: b.Z().d() };
    }
    case "doubleWave": {
      const a = h * Math.min(frac("adj1", 6250), 0.2);
      const b = new PathB().M(0, a);
      b.C(w / 12, 0, w / 6, 0, w / 4, a).C(w / 3, 2 * a, 5 * w / 12, 2 * a, w / 2, a);
      b.C(7 * w / 12, 0, 2 * w / 3, 0, 3 * w / 4, a).C(
        5 * w / 6,
        2 * a,
        11 * w / 12,
        2 * a,
        w,
        a
      );
      b.L(w, h - a);
      b.C(11 * w / 12, h, 5 * w / 6, h, 3 * w / 4, h - a).C(
        2 * w / 3,
        h - 2 * a,
        7 * w / 12,
        h - 2 * a,
        w / 2,
        h - a
      );
      b.C(5 * w / 12, h, w / 3, h, w / 4, h - a).C(w / 6, h - 2 * a, w / 12, h - 2 * a, 0, h - a);
      return { path: b.Z().d() };
    }
    case "uturnArrow": {
      const t = ss * frac("adj1", 25e3);
      const hw = 0.75 * t;
      const hl = t;
      const xrc = w - hw;
      const rxO = (xrc + t / 2) / 2;
      const ryO = Math.min(h / 2, rxO);
      const b = new PathB().M(0, h).L(0, ryO);
      b.arc(rxO, ryO, rxO, ryO, 180, 180);
      const yh = h - hl;
      b.L(xrc + t / 2, yh).L(xrc + hw, yh).L(xrc, h).L(xrc - hw, yh).L(xrc - t / 2, yh).L(xrc - t / 2, ryO);
      b.arc(rxO, ryO, Math.max(rxO - t, 0), Math.max(ryO - t, ryO * 0.2), 0, -180);
      b.L(t, h).Z();
      return { path: b.d() };
    }
    case "curvedRightArrow": {
      const t = ss * frac("adj1", 25e3);
      const b = new PathB().M(0, 0);
      b.arc(0, cy, w, cy, 270, 90);
      const bi = Math.max(w - 1.5 * t, 0);
      b.L((w + bi) / 2, Math.min(h, cy + 1.2 * t)).L(bi, cy).L(w - t, cy);
      b.arc(0, cy, Math.max(w - t, 0), Math.max(cy - t, 0), 0, -90);
      b.L(0, 0).Z();
      return { path: b.d() };
    }
    case "stripedRightArrow": {
      const thick = h * frac("adj1", 5e4);
      const head = Math.min(w, ss * frac("adj2", 5e4));
      const y1 = (h - thick) / 2;
      const y2 = (h + thick) / 2;
      const xh = w - head;
      const bs = ss * 5 / 32;
      const b = new PathB();
      b.M(bs, y1).L(xh, y1).L(xh, 0).L(w, cy).L(xh, h).L(xh, y2).L(bs, y2).Z();
      b.M(0, y1).L(ss / 32, y1).L(ss / 32, y2).L(0, y2).Z();
      b.M(ss / 16, y1).L(ss / 8, y1).L(ss / 8, y2).L(ss / 16, y2).Z();
      return { path: b.d() };
    }
    case "swooshArrow": {
      const a1 = Math.min(Math.max(adjust?.adj1 ?? 25e3, 1), 75e3) / 1e5;
      const maxAdj2 = 7e4 * w / ss;
      const a2 = Math.min(Math.max(adjust?.adj2 ?? 16667, 0), maxAdj2);
      const ad1 = h * a1;
      const ad2 = ss * a2 / 1e5;
      const ssd8 = ss / 8;
      const tanAlfa = Math.tan(Math.PI / 2 / 14);
      const xB = w - ad2;
      const yB = ssd8;
      const xC = xB - ssd8 * tanAlfa;
      const yF = yB + ad1;
      const xF = xB + ad1 * tanAlfa;
      const xE = xF + ssd8 * tanAlfa;
      const yE = yF + ssd8;
      const yD = yE / 2 + h / 20;
      const b = new PathB();
      b.M(0, h).Q(w / 6, h / 3, xB, yB).L(xC, 0).L(w, yD).L(xE, yE).L(xF, yF).Q(w / 4, yF + h / 12, 0, h).Z();
      return { path: b.d() };
    }
    case "wedgeRoundRectCallout": {
      const r = ss * frac("adj3", 16667);
      const tipX = cx + w * adjRaw(adjust, "adj1", -20833);
      const tipY = cy + h * adjRaw(adjust, "adj2", 62500);
      const b = new PathB();
      b.M(r, 0).L(w - r, 0).arc(w - r, r, r, r, 270, 90).L(w, h - r);
      b.arc(w - r, h - r, r, r, 0, 90).L(r, h).arc(r, h - r, r, r, 90, 90).L(0, r);
      b.arc(r, r, r, r, 180, 90).Z();
      appendWedgeTail(b, w, h, tipX, tipY);
      return { path: b.d() };
    }
    case "wedgeEllipseCallout": {
      const tipX = cx + w * adjRaw(adjust, "adj1", -20833);
      const tipY = cy + h * adjRaw(adjust, "adj2", 62500);
      const b = new PathB();
      ellipseSub(b, cx, cy, cx, cy);
      const th = Math.atan2(tipY - cy, tipX - cx);
      b.M(cx + cx * Math.cos(th - 0.3), cy + cy * Math.sin(th - 0.3)).L(tipX, tipY).L(cx + cx * Math.cos(th + 0.3), cy + cy * Math.sin(th + 0.3)).Z();
      return { path: b.d() };
    }
    case "flowChartPredefinedProcess": {
      const path = new PathB().M(0, 0).L(w, 0).L(w, h).L(0, h).Z().d();
      const lines = new PathB().M(w / 8, 0).L(w / 8, h).M(7 * w / 8, 0).L(7 * w / 8, h).d();
      return { path, strokePath: lines };
    }
    case "flowChartInternalStorage": {
      const path = new PathB().M(0, 0).L(w, 0).L(w, h).L(0, h).Z().d();
      const lines = new PathB().M(w / 8, 0).L(w / 8, h).M(0, h / 8).L(w, h / 8).d();
      return { path, strokePath: lines };
    }
    case "flowChartDocument": {
      const b = new PathB().M(0, 0).L(w, 0).L(w, 0.83 * h);
      b.C(0.75 * w, 0.72 * h, 0.58 * w, 0.72 * h, 0.5 * w, 0.83 * h);
      b.C(0.42 * w, 0.94 * h, 0.25 * w, 0.94 * h, 0, 0.83 * h);
      return { path: b.Z().d() };
    }
    case "flowChartMultidocument": {
      const b = new PathB().M(0, 0.12 * h).L(0.88 * w, 0.12 * h).L(0.88 * w, 0.85 * h);
      b.C(0.66 * w, 0.74 * h, 0.51 * w, 0.74 * h, 0.44 * w, 0.85 * h);
      b.C(0.37 * w, 0.96 * h, 0.22 * w, 0.96 * h, 0, 0.85 * h);
      b.Z();
      const backs = new PathB().M(0.06 * w, 0.12 * h).L(0.06 * w, 0.06 * h).L(0.94 * w, 0.06 * h).L(0.94 * w, 0.6 * h).M(0.12 * w, 0.06 * h).L(0.12 * w, 0).L(w, 0).L(w, 0.53 * h);
      return { path: b.d(), strokePath: backs.d() };
    }
    case "flowChartConnector": {
      const b = new PathB();
      ellipseSub(b, cx, cy, cx, cy);
      return { path: b.d() };
    }
    case "flowChartOr": {
      const b = new PathB();
      ellipseSub(b, cx, cy, cx, cy);
      const lines = new PathB().M(cx, 0).L(cx, h).M(0, cy).L(w, cy).d();
      return { path: b.d(), strokePath: lines };
    }
    case "flowChartSummingJunction": {
      const b = new PathB();
      ellipseSub(b, cx, cy, cx, cy);
      const dx = cx * Math.SQRT1_2;
      const dy = cy * Math.SQRT1_2;
      const lines = new PathB().M(cx - dx, cy - dy).L(cx + dx, cy + dy).M(cx + dx, cy - dy).L(cx - dx, cy + dy).d();
      return { path: b.d(), strokePath: lines };
    }
    case "flowChartSort": {
      const path = new PathB().M(cx, 0).L(w, cy).L(cx, h).L(0, cy).Z().d();
      return { path, strokePath: new PathB().M(0, cy).L(w, cy).d() };
    }
    case "flowChartDelay": {
      const b = new PathB().M(0, 0).L(cx, 0);
      b.arc(cx, cy, cx, cy, 270, 180).L(0, h).Z();
      return { path: b.d() };
    }
    case "flowChartDisplay": {
      const b = new PathB().M(0, cy).L(w / 6, 0).L(5 * w / 6, 0);
      b.arc(5 * w / 6, cy, w / 6, cy, 270, 180).L(w / 6, h).Z();
      return { path: b.d() };
    }
    case "flowChartPunchedTape": {
      const a = 0.1 * h;
      const b = new PathB().M(0, a);
      b.C(w / 6, 0, w / 3, 0, w / 2, a).C(2 * w / 3, 2 * a, 5 * w / 6, 2 * a, w, a);
      b.L(w, h - a);
      b.C(5 * w / 6, h - 2 * a, 2 * w / 3, h - 2 * a, w / 2, h - a);
      b.C(w / 3, h, w / 6, h, 0, h - a);
      return { path: b.Z().d() };
    }
    case "leftBracket": {
      const r = Math.min(h / 2, ss * frac("adj", 8333));
      const b = new PathB().arc(w, r, w, r, 270, -90, "M").L(0, h - r).arc(w, h - r, w, r, 180, -90);
      return { strokePath: b.d() };
    }
    case "rightBracket": {
      const r = Math.min(h / 2, ss * frac("adj", 8333));
      const b = new PathB().arc(0, r, w, r, 270, 90, "M").L(w, h - r).arc(0, h - r, w, r, 0, 90);
      return { strokePath: b.d() };
    }
    case "leftBrace": {
      const r = Math.min(h / 4, ss * frac("adj1", 8333));
      const mid = h * frac("adj2", 5e4);
      const xm = w / 2;
      const b = new PathB().arc(w, r, xm, r, 270, -90, "M").L(xm, mid - r);
      b.arc(0, mid - r, xm, r, 0, 90).arc(0, mid + r, xm, r, 270, 90).L(xm, h - r);
      b.arc(w, h - r, xm, r, 180, -90);
      return { strokePath: b.d() };
    }
    case "rightBrace": {
      const r = Math.min(h / 4, ss * frac("adj1", 8333));
      const mid = h * frac("adj2", 5e4);
      const xm = w / 2;
      const b = new PathB().arc(0, r, xm, r, 270, 90, "M").L(xm, mid - r);
      b.arc(w, mid - r, xm, r, 180, -90).arc(w, mid + r, xm, r, 270, -90).L(xm, h - r);
      b.arc(0, h - r, xm, r, 0, 90);
      return { strokePath: b.d() };
    }
    default:
      return null;
  }
}
function appendWedgeTail(b, w, h, tipX, tipY) {
  const g = Math.min(w, h) * 0.1;
  const clamp = (v, lo, hi) => Math.min(Math.max(v, lo), hi);
  const nx = (tipX - w / 2) / w;
  const ny = (tipY - h / 2) / h;
  if (Math.abs(ny) >= Math.abs(nx)) {
    const bx = clamp(tipX, 2 * g, w - 2 * g);
    const ey = ny >= 0 ? h : 0;
    b.M(bx - g, ey).L(tipX, tipY).L(bx + g, ey).Z();
  } else {
    const by = clamp(tipY, 2 * g, h - 2 * g);
    const ex = nx >= 0 ? w : 0;
    b.M(ex, by - g).L(tipX, tipY).L(ex, by + g).Z();
  }
}
