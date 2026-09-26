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

// Inline Word drawings: wpg groups with their wps shapes, pictures, text
// boxes, preset silhouettes, freeform paths and connectors. Word wraps these
// in mc:AlternateContent as a Choice, with a VML Fallback beside it; the
// Choice is the one with the pictures, so dropping it falls back to a blank
// gap where a diagram should be.
//
// The group is one inline-block box at its wp:extent size. Children are
// positioned from their own a:xfrm, mapped through the group's chOff/chExt
// into the box, in document order so the z-order is Word's.

const { childrenOf, firstOf, findFirst, tagName, attr, attrInt, attrNum } = require("../shared/xml");
const { emuToPx } = require("../shared/units");
const { colorFromContainer } = require("../shared/color");
const { presetPath, presetPolygon, connectorPoints, isConnectorPreset } = require("../pptx/presets");

const SVG_NS = "http://www.w3.org/2000/svg";
const DEFAULT_LINE_EMU = 12700;
const DEFAULT_INSET_L_EMU = 91440;
const DEFAULT_INSET_T_EMU = 45720;

function renderShapeGroup(parent, run, opts) {
  const doc = parent.ownerDocument;
  const node = run.node;
  if (!node || !doc) return;
  const extent = groupExtent(node, run);
  if (!extent.width || !extent.height) return;

  const wrapper = parent.createSpan("ov-docx-shapegroup");
  wrapper.style.width = extent.width + "px";
  wrapper.style.height = extent.height + "px";
  if (run.alt || run.name) wrapper.title = run.alt || run.name;

  if (tagName(node) === "wgp") {
    // The group's own child space maps onto the wrapper box; nested groups
    // compose their transforms on top of it.
    const map = childTransform(node, { sx: 1, sy: 1, ox: 0, oy: 0 });
    drawGroupChildren(node, wrapper, map, opts, wrapper);
  } else {
    // A standalone shape: its box is the wp:extent, its own transform inside it.
    drawShape(node, wrapper, { sx: 1, sy: 1, ox: 0, oy: 0 }, opts, wrapper);
  }
}

function groupExtent(node, run) {
  if (run.widthPx > 0 && run.heightPx > 0) {
    return { width: run.widthPx, height: run.heightPx };
  }
  const xfrm = groupTransform(node);
  const ext = xfrm ? firstOf(xfrm, "ext") : null;
  return {
    width: ext ? emuToPx(attrNum(ext, "cx", 0)) : 0,
    height: ext ? emuToPx(attrNum(ext, "cy", 0)) : 0,
  };
}

function groupTransform(node) {
  const pr = firstOf(node, "grpSpPr");
  return pr ? firstOf(pr, "xfrm") : null;
}

// The child space of a group maps onto its box: chOff becomes the box origin
// and chExt scales to ext. An off/ext of zero is an identity transform, where
// children keep their own coordinates (a collapsed group would otherwise draw
// everything at the origin).
function childTransform(node, map) {
  const xfrm = groupTransform(node);
  if (!xfrm) return map;
  const off = firstOf(xfrm, "off");
  const ext = firstOf(xfrm, "ext");
  const chOff = firstOf(xfrm, "chOff");
  const chExt = firstOf(xfrm, "chExt");
  const offX = off ? attrNum(off, "x", 0) : 0;
  const offY = off ? attrNum(off, "y", 0) : 0;
  const extX = ext ? attrNum(ext, "cx", 0) : 0;
  const extY = ext ? attrNum(ext, "cy", 0) : 0;
  const chX = chOff ? attrNum(chOff, "x", 0) : 0;
  const chY = chOff ? attrNum(chOff, "y", 0) : 0;
  const chW = chExt ? attrNum(chExt, "cx", 0) : 0;
  const chH = chExt ? attrNum(chExt, "cy", 0) : 0;
  // A group with off/ext of zero is an identity transform: the child space
  // keeps its own coordinates and the box is chOff/chExt.
  const gx = extX > 0 && chW > 0 ? extX / chW : 1;
  const gy = extY > 0 && chH > 0 ? extY / chH : 1;
  const scaleX = map.sx * gx;
  const scaleY = map.sy * gy;
  const ox = (emuToPx(offX) - emuToPx(chX) * gx) * map.sx + map.ox;
  const oy = (emuToPx(offY) - emuToPx(chY) * gy) * map.sy + map.oy;
  return { sx: scaleX, sy: scaleY, ox, oy };
}

function drawGroupChildren(node, container, map, opts, wrapper) {
  for (const child of node.children || []) {
    const tag = tagName(child);
    if (tag === "wsp" || tag === "shape") drawShape(child, container, map, opts, wrapper);
    else if (tag === "pic" || tag === "picture") drawPictureShape(child, container, map, opts, wrapper);
    else if (tag === "grpSp" || tag === "wgp") drawNestedGroup(child, container, map, opts, wrapper);
  }
}

function drawNestedGroup(node, container, map, opts, wrapper) {
  const inner = childTransform(node, map);
  drawGroupChildren(node, container, inner, opts, wrapper);
}

// ---------- shape boxes ----------

function shapeXfrm(spPr) {
  const xfrm = spPr ? firstOf(spPr, "xfrm") : null;
  if (!xfrm) return null;
  const off = firstOf(xfrm, "off");
  const ext = firstOf(xfrm, "ext");
  return {
    x: off ? attrNum(off, "x", 0) : 0,
    y: off ? attrNum(off, "y", 0) : 0,
    cx: ext ? attrNum(ext, "cx", 0) : 0,
    cy: ext ? attrNum(ext, "cy", 0) : 0,
    rot: attrNum(xfrm, "rot", 0) / 60000,
    flipH: attr(xfrm, "flipH") === "1",
    flipV: attr(xfrm, "flipV") === "1",
  };
}

function boxFor(xfrm, map) {
  return {
    left: emuToPx(xfrm.x) * map.sx + map.ox,
    top: emuToPx(xfrm.y) * map.sy + map.oy,
    width: Math.max(0, emuToPx(xfrm.cx) * map.sx),
    height: Math.max(0, emuToPx(xfrm.cy) * map.sy),
    rot: xfrm.rot,
    flipH: xfrm.flipH,
    flipV: xfrm.flipV,
  };
}

function placeElement(el, box) {
  el.style.left = box.left + "px";
  el.style.top = box.top + "px";
  el.style.width = box.width + "px";
  el.style.height = box.height + "px";
  if (box.rot) {
    el.style.transform = "rotate(" + box.rot + "deg)";
    el.style.transformOrigin = "center";
  }
}

function drawShape(node, container, map, opts, wrapper) {
  const spPr = firstOf(node, "spPr");
  const xfrm = shapeXfrm(spPr);
  if (!xfrm) return;
  const box = boxFor(xfrm, map);
  const geometry = readGeometry(spPr);
  const style = firstOf(node, "style");
  const fill = style ? readFill(spPr, style, opts) : readFill(spPr, null, opts);
  const line = style ? readLine(spPr, style, opts) : readLine(spPr, null, opts);

  const el = container.ownerDocument.createElement("div");
  el.className = "ov-docx-shape";
  placeElement(el, box);
  container.appendChild(el);

  const isTextOnly = geometry.preset === "textNoShape";
  const hasGeometry = (box.width > 0 || box.height > 0) && !isTextOnly;
  if (hasGeometry && (fill || line)) {
    drawGeometry(el, geometry, box, fill, line, opts, container.ownerDocument);
  }
  renderTextBody(el, node, opts);
}

function drawPictureShape(node, container, map, opts, wrapper) {
  const spPr = firstOf(node, "spPr");
  const xfrm = shapeXfrm(spPr);
  if (!xfrm) return;
  const box = boxFor(xfrm, map);
  const blipFill = firstOf(node, "blipFill");
  const blip = blipFill ? firstOf(blipFill, "blip") : null;
  const rid = blip ? (attr(blip, "embed") || attr(blip, "link")) : null;
  const url = rid && opts.mediaUrl ? opts.mediaUrl(rid) : null;
  const el = container.ownerDocument.createElement("div");
  el.className = "ov-docx-shape ov-docx-shape-picture";
  placeElement(el, box);
  container.appendChild(el);
  if (!url) {
    el.addClass ? el.addClass("is-missing") : el.classList.add("is-missing");
    el.textContent = "Image could not be drawn";
    return;
  }
  const img = container.ownerDocument.createElement("img");
  img.src = url;
  const nv = findFirst(node, "nvPicPr");
  const cnv = nv ? firstOf(nv, "cNvPr") : null;
  img.alt = cnv ? (attr(cnv, "descr") || attr(cnv, "name") || "") : "";
  img.setAttribute("draggable", "false");
  img.style.width = "100%";
  img.style.height = "100%";
  const stretch = blipFill ? firstOf(blipFill, "stretch") : null;
  img.style.objectFit = stretch || !blipFill ? "fill" : "contain";
  const geometry = readGeometry(spPr);
  if (geometry.preset && geometry.preset !== "rect" && geometry.preset !== "rectangle") {
    applyGeometryClip(el, geometry, box, opts);
  }
  el.appendChild(img);
}

// ---------- fills and lines ----------

function readFill(spPr, style, opts) {
  if (spPr) {
    for (const child of spPr.children || []) {
      const tag = tagName(child);
      if (tag === "noFill") return null;
      if (tag === "solidFill") {
        const color = colorFromContainer(child, opts.theme, null);
        if (color) return { type: "solid", color };
      } else if (tag === "gradFill") {
        const color = gradientFirstColor(child, opts);
        if (color) return { type: "solid", color };
      } else if (tag === "blipFill") {
        const blip = firstOf(child, "blip");
        const rid = blip ? (attr(blip, "embed") || attr(blip, "link")) : null;
        const url = rid && opts.mediaUrl ? opts.mediaUrl(rid) : null;
        if (url) return { type: "image", url };
      }
    }
  }
  const ref = style ? firstOf(style, "fillRef") : null;
  if (ref && attrInt(ref, "idx", 0) > 0) {
    const color = colorFromContainer(ref, opts.theme, null);
    if (color) return { type: "solid", color };
  }
  return null;
}

function gradientFirstColor(grad, opts) {
  const stops = grad ? childrenOf(grad, "gs") : [];
  let best = null;
  let bestPos = Infinity;
  for (const stop of stops) {
    const pos = attrNum(stop, "pos", 0);
    if (pos < bestPos) {
      bestPos = pos;
      best = stop;
    }
  }
  return best ? colorFromContainer(best, opts.theme, null) : null;
}

function readLine(spPr, style, opts) {
  const ln = spPr ? firstOf(spPr, "ln") : null;
  const ref = style ? firstOf(style, "lnRef") : null;
  const refActive = ref && attrInt(ref, "idx", 0) > 0;
  const refColor = refActive ? colorFromContainer(ref, opts.theme, null) : null;
  if (ln) {
    if (firstOf(ln, "noFill")) return null;
    const fill = firstOf(ln, "solidFill") || firstOf(ln, "gradFill");
    // An explicit a:ln overrides only what it states. A line that names a
    // width and no colour still takes the style's line colour.
    const color = (fill ? colorFromContainer(fill, opts.theme, null) : null) || refColor;
    const head = firstOf(ln, "headEnd");
    const tail = firstOf(ln, "tailEnd");
    const widthEmu = attrInt(ln, "w", DEFAULT_LINE_EMU) || DEFAULT_LINE_EMU;
    if (!color && !head && !tail) return null;
    return {
      color: color || "#000000",
      widthEmu,
      dash: attr(firstOf(ln, "prstDash"), "val") || null,
      customDash: readCustomDash(ln),
      headEnd: head && attr(head, "type") !== "none" ? attr(head, "type") : null,
      tailEnd: tail && attr(tail, "type") !== "none" ? attr(tail, "type") : null,
    };
  }
  if (refColor) return { color: refColor, widthEmu: DEFAULT_LINE_EMU, dash: null, customDash: null, headEnd: null, tailEnd: null };
  return null;
}

// a:custDash writes each dash and gap as thousandths of the line width.
function readCustomDash(ln) {
  const cust = firstOf(ln, "custDash");
  if (!cust) return null;
  const out = [];
  for (const ds of childrenOf(cust, "ds")) {
    const d = attrNum(ds, "d", 0) / 100000;
    const sp = attrNum(ds, "sp", 0) / 100000;
    if (d > 0) out.push(d);
    if (sp > 0) out.push(sp);
  }
  return out.length ? out : null;
}

// ---------- geometry ----------

function readGeometry(spPr) {
  if (!spPr) return { preset: "rect", adjustments: null };
  const prst = firstOf(spPr, "prstGeom");
  if (prst) return { preset: attr(prst, "prst") || "rect", adjustments: readAdjustments(prst) };
  const cust = firstOf(spPr, "custGeom");
  if (cust) return { preset: "custom", node: cust };
  return { preset: "rect", adjustments: null };
}

function readAdjustments(geom) {
  const avLst = firstOf(geom, "avLst");
  if (!avLst) return null;
  const out = {};
  for (const gd of childrenOf(avLst, "gd")) {
    const name = attr(gd, "name");
    const fmla = attr(gd, "fmla") || "";
    const value = /val\s+(-?\d+)/.exec(fmla);
    if (name && value) out[name] = Number(value[1]);
  }
  return Object.keys(out).length ? out : null;
}

// ---------- geometry drawing ----------

function drawGeometry(el, geometry, box, fill, line, opts, doc) {
  const width = Math.max(box.width, 1);
  const height = Math.max(box.height, 1);
  const lineWidthPx = line ? Math.max(0.5, emuToPx(line.widthEmu)) : 0;
  const preset = geometry.preset || "rect";

  if (isConnectorPreset(preset)) {
    drawConnector(el, geometry, box, line, lineWidthPx, doc);
    return;
  }
  if (preset === "rect" || preset === "rectangle") {
    const svg = makeSvg(doc, width, height);
    const rect = svgEl(doc, "rect");
    rect.setAttribute("x", "0");
    rect.setAttribute("y", "0");
    rect.setAttribute("width", String(width));
    rect.setAttribute("height", String(height));
    applySvgFill(rect, fill, opts);
    applySvgStroke(rect, line, lineWidthPx);
    svg.appendChild(rect);
    el.appendChild(svg);
    return;
  }
  if (preset === "ellipse" || preset === "circle") {
    const svg = makeSvg(doc, width, height);
    const ellipse = svgEl(doc, "ellipse");
    ellipse.setAttribute("cx", String(width / 2));
    ellipse.setAttribute("cy", String(height / 2));
    ellipse.setAttribute("rx", String(width / 2));
    ellipse.setAttribute("ry", String(height / 2));
    applySvgFill(ellipse, fill, opts);
    applySvgStroke(ellipse, line, lineWidthPx);
    svg.appendChild(ellipse);
    el.appendChild(svg);
    return;
  }

  let pathData = "";
  let fillData = "";
  if (preset === "custom") {
    const paths = customPaths(geometry.node, box.width, box.height);
    pathData = paths.stroke;
    fillData = paths.fill;
  } else {
    // Straight edged silhouettes come back as points; curved ones as a path.
    const points = box.width > 0 && box.height > 0
      ? presetPolygon(preset, box.width, box.height, geometry.adjustments)
      : null;
    if (points && points.length >= 6) {
      const svg = makeSvg(doc, width, height);
      const poly = svgEl(doc, "polygon");
      const list = [];
      for (let i = 0; i + 1 < points.length; i += 2) {
        list.push(round1(points[i]) + "," + round1(points[i + 1]));
      }
      poly.setAttribute("points", list.join(" "));
      applySvgFill(poly, fill, opts);
      applySvgStroke(poly, line, lineWidthPx);
      svg.appendChild(poly);
      el.appendChild(svg);
      return;
    }
    const drawn = presetPath(preset, box.width, box.height, geometry.adjustments);
    if (drawn) {
      pathData = drawn.strokePath || drawn.path || drawn.fillPath || "";
      fillData = drawn.fillPath || drawn.path || "";
    }
  }
  if (!pathData && !fillData) return;
  const svg = makeSvg(doc, width, height);
  if (fillData && fill) {
    const filled = svgEl(doc, "path");
    filled.setAttribute("d", fillData);
    applySvgFill(filled, fill, opts);
    filled.setAttribute("stroke", "none");
    svg.appendChild(filled);
  } else if (pathData && fill && fill.type === "solid") {
    const filled = svgEl(doc, "path");
    filled.setAttribute("d", pathData);
    applySvgFill(filled, fill, opts);
    filled.setAttribute("stroke", "none");
    svg.appendChild(filled);
  }
  if (line && pathData) {
    const stroked = svgEl(doc, "path");
    stroked.setAttribute("d", pathData);
    stroked.setAttribute("fill", "none");
    applySvgStroke(stroked, line, lineWidthPx);
    svg.appendChild(stroked);
  }
  el.appendChild(svg);
}

// a:custGeom path coordinate space is its own; scale it to the shape box. The
// path list keeps a fill path and a stroke path separate, because a path with
// fill="none" is a line and must not become a solid silhouette.
function customPaths(geom, width, height) {
  const pathLst = firstOf(geom, "pathLst");
  const fill = [];
  const stroke = [];
  if (!pathLst) return { fill: "", stroke: "" };
  for (const path of childrenOf(pathLst, "path")) {
    const w = attrNum(path, "w", 0);
    const h = attrNum(path, "h", 0);
    const sx = w > 0 ? width / w : 1;
    const sy = h > 0 ? height / h : 1;
    const parts = [];
    let cx = 0;
    let cy = 0;
    let startX = 0;
    let startY = 0;
    for (const segment of path.children || []) {
      const tag = tagName(segment);
      if (tag === "moveTo" || tag === "lnTo") {
        const pt = firstOf(segment, "pt");
        if (!pt) continue;
        cx = attrNum(pt, "x", 0);
        cy = attrNum(pt, "y", 0);
        if (tag === "moveTo") {
          startX = cx;
          startY = cy;
          parts.push("M" + round1(cx * sx) + " " + round1(cy * sy));
        } else {
          parts.push("L" + round1(cx * sx) + " " + round1(cy * sy));
        }
      } else if (tag === "cubicBezTo") {
        const pts = childrenOf(segment, "pt");
        if (pts.length < 3) continue;
        const [c1, c2, end] = pts;
        parts.push("C" + round1(attrNum(c1, "x", 0) * sx) + " " + round1(attrNum(c1, "y", 0) * sy) + " "
          + round1(attrNum(c2, "x", 0) * sx) + " " + round1(attrNum(c2, "y", 0) * sy) + " "
          + round1(attrNum(end, "x", 0) * sx) + " " + round1(attrNum(end, "y", 0) * sy));
        cx = attrNum(end, "x", 0);
        cy = attrNum(end, "y", 0);
      } else if (tag === "quadBezTo") {
        const pts = childrenOf(segment, "pt");
        if (pts.length < 2) continue;
        const [c1, end] = pts;
        parts.push("Q" + round1(attrNum(c1, "x", 0) * sx) + " " + round1(attrNum(c1, "y", 0) * sy) + " "
          + round1(attrNum(end, "x", 0) * sx) + " " + round1(attrNum(end, "y", 0) * sy));
        cx = attrNum(end, "x", 0);
        cy = attrNum(end, "y", 0);
      } else if (tag === "arcTo") {
        const wR = attrNum(segment, "wR", 0);
        const hR = attrNum(segment, "hR", 0);
        const stAng = attrNum(segment, "stAng", 0);
        const swAng = attrNum(segment, "swAng", 0);
        const startAngle = (stAng / 60000) * Math.PI / 180;
        const sweepAngle = (swAng / 60000) * Math.PI / 180;
        const centerX = cx - wR * Math.cos(startAngle);
        const centerY = cy - hR * Math.sin(startAngle);
        const endX = centerX + wR * Math.cos(startAngle + sweepAngle);
        const endY = centerY + hR * Math.sin(startAngle + sweepAngle);
        const large = Math.abs(swAng) > 10800000 ? 1 : 0;
        const sweep = swAng >= 0 ? 1 : 0;
        parts.push("A" + round1(wR * sx) + " " + round1(hR * sy) + " 0 " + large + " " + sweep + " "
          + round1(endX * sx) + " " + round1(endY * sy));
        cx = endX;
        cy = endY;
      } else if (tag === "close") {
        parts.push("Z");
        cx = startX;
        cy = startY;
      }
    }
    const data = parts.join(" ");
    stroke.push(data);
    if (attr(path, "fill") !== "none") fill.push(data);
  }
  return { fill: fill.join(" "), stroke: stroke.join(" ") };
}

// A connector or a plain line: one SVG line, with the arrowheads the file asks
// for. A horizontal or vertical connector has a zero extent on one axis, so the
// viewport gets one line width there and the line is centred in it.
function drawConnector(el, geometry, box, line, lineWidthPx, doc) {
  const preset = geometry.preset || "line";
  const width = box.width;
  const height = box.height;
  const flipH = box.flipH;
  const flipV = box.flipV;
  const pts = connectorPoints(preset, width, height, flipH, flipV, geometry.adjustments);
  if (!pts || pts.length < 4) return;
  const lw = Math.max(lineWidthPx || 1.33, 1);
  const svgWidth = Math.max(width, lw);
  const svgHeight = Math.max(height, lw);
  const svg = makeSvg(doc, svgWidth, svgHeight);
  if (width <= 0) svg.style.left = -svgWidth / 2 + "px";
  if (height <= 0) svg.style.top = -svgHeight / 2 + "px";
  const x1 = pts[0];
  const y1 = pts[1];
  const x2 = pts[pts.length - 2];
  const y2 = pts[pts.length - 1];
  const poly = svgEl(doc, "polyline");
  const points = [];
  for (let i = 0; i + 1 < pts.length; i += 2) points.push(pts[i] + "," + pts[i + 1]);
  poly.setAttribute("points", points.join(" "));
  poly.setAttribute("fill", "none");
  applySvgStroke(poly, line, lineWidthPx || 1);
  svg.appendChild(poly);
  if (line) {
    if (line.headEnd) addArrow(doc, svg, line.headEnd, { x: x1, y: y1 }, { x: x2, y: y2 }, line.color, lineWidthPx || 1);
    if (line.tailEnd) addArrow(doc, svg, line.tailEnd, { x: x2, y: y2 }, { x: x1, y: y1 }, line.color, lineWidthPx || 1);
  }
  el.appendChild(svg);
}

const ARROW_SHAPES = {
  triangle: [[0, 0], [-2, 0.72], [-2, -0.72]],
  arrow: [[0, 0], [-2, 0.72], [-1.2, 0], [-2, -0.72]],
  stealth: [[0, 0], [-2, 0.72], [-1.2, 0], [-2, -0.72]],
};

function addArrow(doc, svg, spec, tip, from, color, lineWidthPx) {
  const shape = ARROW_SHAPES[spec] || ARROW_SHAPES.triangle;
  const size = Math.max(6, lineWidthPx * 3.2);
  const angle = Math.atan2(tip.y - from.y, tip.x - from.x);
  const points = shape.map(([x, y]) => {
    const px = tip.x + (x * Math.cos(angle) - y * Math.sin(angle)) * size * 0.65;
    const py = tip.y + (x * Math.sin(angle) + y * Math.cos(angle)) * size * 0.65;
    return round1(px) + "," + round1(py);
  });
  const polygon = svgEl(doc, "polygon");
  polygon.setAttribute("points", points.join(" "));
  polygon.setAttribute("fill", color || "#000000");
  svg.appendChild(polygon);
}

function applySvgFill(el, fill, opts) {
  if (!fill) {
    el.setAttribute("fill", "none");
    return;
  }
  if (fill.type === "image") {
    el.setAttribute("fill", "#ffffff");
    return;
  }
  el.setAttribute("fill", fill.color || "none");
}

function applySvgStroke(el, line, widthPx) {
  if (!line) {
    el.setAttribute("stroke", "none");
    return;
  }
  el.setAttribute("stroke", line.color || "#000000");
  el.setAttribute("stroke-width", String(Math.max(0.5, widthPx)));
  if (line.customDash) {
    const unit = Math.max(0.5, widthPx);
    el.setAttribute("stroke-dasharray", line.customDash.map((v) => String(round1(v * unit))).join(" "));
  } else if (line.dash) {
    const pattern = dashArray(line.dash, widthPx);
    if (pattern) el.setAttribute("stroke-dasharray", pattern);
  }
}

function dashArray(dash, widthPx) {
  const unit = Math.max(1, widthPx);
  switch (dash) {
    case "dot": return unit + " " + unit * 2;
    case "dash": return unit * 4 + " " + unit * 3;
    case "lgDash": return unit * 8 + " " + unit * 3;
    case "dashDot": return unit * 4 + " " + unit * 2 + " " + unit + " " + unit * 2;
    case "sysDash": return unit * 3 + " " + unit;
    default: return null;
  }
}

function applyGeometryClip(el, geometry, box, opts) {
  if (geometry.preset === "custom") return;
  const drawn = presetPath(geometry.preset, box.width, box.height, geometry.adjustments);
  const data = drawn ? (drawn.path || drawn.fillPath) : "";
  if (data) el.style.clipPath = 'path("' + data + '")';
}

// ---------- text ----------

function renderTextBody(box, node, opts) {
  const txbx = firstOf(node, "txbx");
  const content = txbx ? firstOf(txbx, "txbxContent") : null;
  if (!content) return;
  const bodyPr = firstOf(node, "bodyPr");
  const anchor = bodyPr ? (attr(bodyPr, "anchor") || "t") : "t";
  box.style.display = "flex";
  box.style.flexDirection = "column";
  box.style.justifyContent = anchor === "ctr" ? "center" : anchor === "b" ? "flex-end" : "flex-start";
  if (bodyPr) {
    box.style.paddingLeft = emuToPx(attrInt(bodyPr, "lIns", DEFAULT_INSET_L_EMU)) + "px";
    box.style.paddingRight = emuToPx(attrInt(bodyPr, "rIns", DEFAULT_INSET_L_EMU)) + "px";
    box.style.paddingTop = emuToPx(attrInt(bodyPr, "tIns", DEFAULT_INSET_T_EMU)) + "px";
    box.style.paddingBottom = emuToPx(attrInt(bodyPr, "bIns", DEFAULT_INSET_T_EMU)) + "px";
  }
  for (const p of childrenOf(content, "p")) {
    let block = null;
    try {
      block = opts.parseParagraph(p);
    } catch (err) {
      block = null;
    }
    if (block) opts.drawParagraph(block, box);
  }
}

// ---------- svg helpers ----------

function makeSvg(doc, width, height) {
  const svg = doc.createElementNS ? doc.createElementNS(SVG_NS, "svg") : doc.createElement("svg");
  svg.setAttribute("width", String(width));
  svg.setAttribute("height", String(height));
  svg.setAttribute("viewBox", "0 0 " + width + " " + height);
  svg.setAttribute("preserveAspectRatio", "none");
  svg.style.position = "absolute";
  svg.style.left = "0";
  svg.style.top = "0";
  svg.style.overflow = "visible";
  if (svg.classList) svg.classList.add("ov-docx-shape-svg");
  return svg;
}

function svgEl(doc, name) {
  return doc.createElementNS ? doc.createElementNS(SVG_NS, name) : doc.createElement(name);
}

function round1(value) {
  return Math.round(value * 100) / 100;
}

module.exports = { renderShapeGroup };
