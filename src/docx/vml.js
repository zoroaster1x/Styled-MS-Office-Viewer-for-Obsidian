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

// Legacy VML drawings: w:pict shapes from older Word files, headers and the
// Fallback branch of mc:AlternateContent when no DrawingML Choice exists. The
// vocabulary is small: shape, rect, roundrect, oval, line, polyline, group and
// the fill, stroke and textbox children. Coordinates inside a group are in the
// group's coordinate space (often twips), while a top level shape writes points
// in its style.

const { childrenOf, firstOf, findAll, findFirst, tagName, attr, attrNs } = require("../shared/xml");
const { applyAnchorStyles, makeSvg, svgEl, applySvgFill, applySvgStroke, round1 } = require("./drawing");

const SHAPE_TAGS = new Set(["shape", "rect", "roundrect", "oval", "arc", "curve", "line", "polyline", "group"]);
const NAMED_COLORS = {
  black: "#000000", white: "#ffffff", red: "#ff0000", green: "#008000", blue: "#0000ff",
  yellow: "#ffff00", gray: "#808080", grey: "#808080", silver: "#c0c0c0", maroon: "#800000",
  navy: "#000080", purple: "#800080", fuchsia: "#ff00ff", aqua: "#00ffff", teal: "#008080",
  lime: "#00ff00", olive: "#808000", windowtext: "#000000", window: "#ffffff",
  activeborder: "#000000", inactiveborder: "#808080", buttonface: "#f0f0f0",
};

// ---------- parse ----------

// A run for the flow model: size for the estimator, out-of-flow offsets for a
// position:absolute pict, and the text for search and the fidelity sweep.
function parseVmlRun(node) {
  const root = findRootShape(node);
  if (!root) return null;
  const style = parseStyle(attr(root, "style"));
  const isLine = tagName(root) === "line";
  let width = absLength(style.width);
  let height = absLength(style.height);
  let left = absLength(style["margin-left"]) || absLength(style.left) || 0;
  let top = absLength(style["margin-top"]) || absLength(style.top) || 0;
  // A top level VML line writes its ends as lengths in points. Its box is the
  // bounding box of the two ends, and the from/to points are what the fraction
  // bars in an equation sheet are: style margins are absent on those shapes.
  let line = null;
  if (isLine) {
    const from = parsePoint(attr(root, "from"));
    const to = parsePoint(attr(root, "to"));
    if (from && to) {
      line = { x1: from.x, y1: from.y, x2: to.x, y2: to.y };
      left = Math.min(from.x, to.x);
      top = Math.min(from.y, to.y);
      width = Math.abs(to.x - from.x);
      height = Math.abs(to.y - from.y);
    }
  }
  // A VML pict is an absolutely positioned overlay anchored to its paragraph.
  // It never takes room in the line by itself: the wrap type decides whether
  // the text flows beside it (square, tight), is pushed above and below it
  // (topAndBottom) or flows through it (none, through).
  const relativeH = (style["mso-position-horizontal-relative"] || "text").toLowerCase();
  const relativeV = (style["mso-position-vertical-relative"] || "text").toLowerCase();
  // Word writes the wrap in a w10:wrap child when it is not the default; the
  // style carries it for shapes edited from the text box tool. A picture with
  // no wrap information still floats (Word's default for an anchored
  // picture), while a bare vector shape (the rectangles drawn around an
  // equation) is an overlay that does not move the text.
  const wrapEl = firstOf(root, "wrap");
  const hasPicture = Boolean(findFirst(root, "imagedata"));
  // A line is never a wrap obstacle: it has no height to reserve, and treating
  // it as a float moves the text it belongs to.
  const wrapStyle = isLine ? "none" : (style["mso-wrap-style"]
    || (wrapEl && attr(wrapEl, "type"))
    || (hasPicture ? "square" : "none")).toLowerCase();
  // Page- and margin-relative verticals are placed against the page, not the
  // paragraph, and reserve nothing in the flow.
  const pageVertical = relativeV === "page" || relativeV === "margin";
  const context = pageVertical ? "page" : "paragraph";
  const texts = [];
  for (const t of findAll(node, "t")) {
    const value = t.textContent || "";
    if (value.trim()) texts.push(value);
  }
  return {
    type: "vml",
    node,
    root,
    widthPx: width || 0,
    heightPx: height || 0,
    texts,
    relativeH,
    relativeV,
    wrapStyle,
    isLine,
    line,
    offsetPx: { left, top },
    anchor: {
      outOfFlow: pageVertical,
      context,
      h: { offsetPx: left, from: relativeH },
      v: { offsetPx: top, from: relativeV },
      behindDoc: Number(style["z-index"] || 0) < 0,
    },
  };
}

// "105.65pt,17.75pt" or "1200,300" as a pair of pixel values. A group's child
// coordinates are unitless; a top level line writes points.
function parsePoint(value) {
  const parts = String(value || "").split(",");
  if (parts.length < 2) return null;
  const x = absLength(parts[0]);
  const y = absLength(parts[1]);
  if (!parts[0].trim() || !parts[1].trim()) return null;
  return { x, y };
}

function findRootShape(node) {
  for (const child of node.children || []) {
    if (SHAPE_TAGS.has(tagName(child))) return child;
    const inner = findRootShape(child);
    if (inner) return inner;
  }
  return null;
}

function parseStyle(value) {
  const out = {};
  for (const part of String(value || "").split(";")) {
    const i = part.indexOf(":");
    if (i === -1) continue;
    out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim();
  }
  return out;
}

function absLength(value) {
  if (value == null) return 0;
  if (/^-?[\d.]+%$/.test(String(value).trim())) return 0;
  const m = /^(-?[\d.]+)(pt|px|in|cm|mm|pc|em)?$/.exec(String(value).trim());
  if (!m) return 0;
  const n = parseFloat(m[1]);
  if (isNaN(n)) return 0;
  switch (m[2]) {
    case "pt": return n * 96 / 72;
    case "in": return n * 96;
    case "cm": return n * 96 / 2.54;
    case "mm": return n * 96 / 25.4;
    case "pc": return n * 16;
    case "em": return n * 16;
    default: return n;
  }
}

function cssColor(value) {
  const text = String(value == null ? "" : value).trim().toLowerCase();
  if (!text || text === "none") return null;
  if (NAMED_COLORS[text]) return NAMED_COLORS[text];
  const m = /^#?([0-9a-f]{6})$/.exec(text);
  if (m) return "#" + m[1];
  const short = /^#?([0-9a-f]{3})$/.exec(text);
  if (short) return "#" + short[1][0] + short[1][0] + short[1][1] + short[1][1] + short[1][2] + short[1][2];
  return text;
}

// ---------- render ----------

function renderVml(parent, run, opts) {
  const root = run.root || findRootShape(run.node);
  if (!root) return;
  const doc = parent.ownerDocument;
  const width = run.widthPx > 0 ? run.widthPx : 160;
  const height = run.heightPx > 0 ? run.heightPx : 0;
  const left = (run.offsetPx && run.offsetPx.left) || 0;
  const top = (run.offsetPx && run.offsetPx.top) || 0;
  const marginLeft = opts.pageMarginLeftPx || 0;
  const marginTop = opts.pageMarginTopPx || 0;
  const contentWidth = opts.contentWidthPx || 0;
  // The wrapper is absolute inside the anchor paragraph, whose border box
  // starts at the text margin, which is also the VML text/column origin. A
  // page-relative axis counts from the page edge instead, and its wrapper is
  // appended to the page content element, whose origin is the text margin.
  const pageVertical = run.relativeV === "page" || run.relativeV === "margin";
  const pageHost = pageVertical ? (opts.pageHost || null) : null;
  const cssLeft = run.relativeH === "page" ? left - marginLeft : left;
  const cssTop = run.relativeV === "page" ? top - marginTop : top;

  // The flow effect of the shape. Square and tight wrap float, so the lines
  // beside the box keep their place (the Fv' fractions left of the y2 box).
  // topAndBottom pushes the following lines below the shape's bottom edge; the
  // paragraph carries that height, so two boxes in one paragraph reserve the
  // taller of the two instead of the sum. None and through let the text pass
  // under the drawing and reserve nothing.
  if (!pageVertical) {
    if (run.wrapStyle === "square" || run.wrapStyle === "tight") {
      const spacer = parent.createSpan("ov-docx-vml-spacer");
      spacer.style.display = "block";
      spacer.style.width = width + "px";
      spacer.style.height = Math.max(0, height) + "px";
      spacer.style.marginTop = top + "px";
      if (contentWidth && left + width / 2 > contentWidth / 2) {
        // A float's own box starts at the paragraph content edge, so its left
        // margin takes the paragraph indent off, while its right margin needs
        // no such correction. A negative margin is legal and lets a shape the
        // file puts past the margin sit there.
        spacer.style.float = "right";
        spacer.style.marginRight = (contentWidth - (left + width)) + "px";
      } else {
        spacer.style.float = "left";
        spacer.style.marginLeft = (left - (opts.plannedLeft || 0)) + "px";
      }
      parent.appendChild(spacer);
    } else if ((run.wrapStyle === "topandbottom" || run.wrapStyle === "through")) {
      // Top-and-bottom pushes the lines below the box; through lets the text
      // pass over the drawing, but the reference suites still give the anchor
      // paragraph the shape's room (a cover canvas would otherwise let the
      // signature line run through it). One block spacer per paragraph, at the
      // first such box, so the paragraph's own text lands under it.
      if (opts.takeReserve) {
        const room = opts.takeReserve();
        if (room > 0) {
          const spacer = parent.createSpan("ov-docx-vml-spacer");
          spacer.style.display = "block";
          spacer.style.width = "100%";
          // The paragraph's own line follows the spacer, so it stands short by
          // one line and the paragraph still ends at the box's bottom edge.
          spacer.style.height = Math.max(0, room - (opts.lineHeightPx || 0)) + "px";
          parent.appendChild(spacer);
        }
      } else if (opts.reserve) {
        opts.reserve(Math.max(0, top + height));
      }
    }
  }

  // The shape itself sits at its own offset, above the line it reserved.
  const wrapper = (pageHost || parent).createSpan("ov-docx-vml");
  wrapper.style.width = width + "px";
  if (height) wrapper.style.height = height + "px";
  wrapper.style.position = "absolute";
  wrapper.style.left = cssLeft + "px";
  wrapper.style.top = cssTop + "px";
  // A negative z-index in the file means behind the text, not behind the page:
  // a real -1 paints under the white page background in CSS and disappears.
  wrapper.style.zIndex = "1";
  if (wrapper.addClass) wrapper.addClass("ov-docx-anchor");
  (pageHost || parent).appendChild(wrapper);

  // The wrapper is the root shape's own box: draw its contents at the origin.
  // Passing the margins again as a root map is what doubled every offset and
  // pushed the equation boxes off the page.
  const stroke = vmlStroke(root);
  if (run.isLine && run.line) {
    const line = run.line;
    drawLineBox(wrapper, line.x1 - left, line.y1 - top, line.x2 - left, line.y2 - top, stroke || { color: "#000000", widthPx: 1 }, doc);
    return;
  }
  drawShape(root, wrapper, { sx: 1, sy: 1, ox: 0, oy: 0, root: true }, doc, opts);
}

function drawShape(shape, container, map, doc, opts) {
  const tag = tagName(shape);
  if (tag === "group") {
    drawGroup(shape, container, map, doc, opts);
    return;
  }
  const style = parseStyle(attr(shape, "style"));
  const box = shapeBox(shape, style, map);
  const stroke = vmlStroke(shape);
  if (tag === "line") {
    drawChildLine(container, shape, map, stroke, doc);
    return;
  }
  const host = doc.createElement("div");
  host.className = "ov-docx-shape";
  host.style.left = (map.root ? 0 : box.left) + "px";
  host.style.top = (map.root ? 0 : box.top) + "px";
  host.style.width = box.width + "px";
  host.style.height = box.height + "px";
  const rotation = parseFloat(style.rotation || "");
  if (!isNaN(rotation)) {
    host.style.transform = "rotate(" + rotation + "deg)";
    host.style.transformOrigin = "center";
  }
  container.appendChild(host);

  // A straight connector is a v:shape that references a one dimensional
  // shapetype; it has no path of its own, so it would otherwise draw nothing.
  if (vmlOneDimensional(shape)) {
    drawLineBox(host, 0, 0, box.width, box.height, stroke || { color: "#000000", widthPx: 1 }, doc);
    return;
  }

  if (tag === "polyline") {
    drawPolyline(host, shape, box, stroke, doc);
    return;
  }

  const imagedata = firstOf(shape, "imagedata");
  const url = imagedata ? resolveImage(imagedata, opts) : null;
  if (url && !firstOf(shape, "textbox")) {
    const img = doc.createElement("img");
    img.src = url;
    img.style.width = "100%";
    img.style.height = "100%";
    img.setAttribute("draggable", "false");
    host.appendChild(img);
    if (stroke) drawOutline(host, tag, box, null, stroke, doc);
    return;
  }

  const fill = vmlFill(shape);
  const geometry = pathData(shape, box.width, box.height);
  const hasText = Boolean(firstOf(shape, "textbox"));
  const drawFill = hasText ? fill : (fill || { type: "none" });
  if (tag === "oval") {
    const svg = makeSvg(doc, Math.max(box.width, 1), Math.max(box.height, 1));
    const ellipse = svgEl(doc, "ellipse");
    ellipse.setAttribute("cx", String(box.width / 2));
    ellipse.setAttribute("cy", String(box.height / 2));
    ellipse.setAttribute("rx", String(Math.max(box.width / 2 - 0.5, 0.5)));
    ellipse.setAttribute("ry", String(Math.max(box.height / 2 - 0.5, 0.5)));
    applySvgFill(ellipse, fill, opts);
    applyVmlStroke(ellipse, stroke);
    svg.appendChild(ellipse);
    host.appendChild(svg);
  } else if (geometry) {
    const svg = makeSvg(doc, Math.max(box.width, 1), Math.max(box.height, 1));
    const path = svgEl(doc, "path");
    path.setAttribute("d", geometry);
    applySvgFill(path, drawFill, opts);
    applyVmlStroke(path, stroke);
    svg.appendChild(path);
    host.appendChild(svg);
  } else if (fill || stroke) {
    const svg = makeSvg(doc, Math.max(box.width, 1), Math.max(box.height, 1));
    const rect = svgEl(doc, "rect");
    const radius = tag === "roundrect" ? roundRectRadius(shape, box) : 0;
    rect.setAttribute("x", "0");
    rect.setAttribute("y", "0");
    rect.setAttribute("width", String(box.width));
    rect.setAttribute("height", String(box.height));
    if (radius) {
      rect.setAttribute("rx", String(radius));
      rect.setAttribute("ry", String(radius));
    }
    applySvgFill(rect, fill, opts);
    applyVmlStroke(rect, stroke);
    svg.appendChild(rect);
    host.appendChild(svg);
  }

  renderVmlText(host, shape, doc, opts);
}

function drawGroup(group, container, map, doc, opts) {
  const style = parseStyle(attr(group, "style"));
  const width = absLength(style.width);
  const height = absLength(style.height);
  const host = doc.createElement("div");
  host.className = "ov-docx-shape ov-docx-vml-group";
  const box = shapeBox(group, style, map);
  host.style.left = (map.root ? 0 : box.left) + "px";
  host.style.top = (map.root ? 0 : box.top) + "px";
  host.style.width = (width || box.width) + "px";
  host.style.height = (height || box.height) + "px";
  container.appendChild(host);

  const coordsize = String(attr(group, "coordsize") || "").split(",");
  const coordorigin = String(attr(group, "coordorigin") || "0,0").split(",");
  const sizeX = parseFloat(coordsize[0]) || 0;
  const sizeY = parseFloat(coordsize[1]) || 0;
  const originX = parseFloat(coordorigin[0]) || 0;
  const originY = parseFloat(coordorigin[1]) || 0;
  const sx = sizeX > 0 ? (width || box.width) / sizeX : 1;
  const sy = sizeY > 0 ? (height || box.height) / sizeY : 1;
  // Children are placed inside the group's own box, so the group's map origin
  // does not apply to them; only the coordinate origin and scale do. Carrying
  // the parent's offset in would offset a nested group's children twice.
  const inner = {
    sx: map.sx * sx,
    sy: map.sy * sy,
    ox: (0 - originX * sx) * map.sx,
    oy: (0 - originY * sy) * map.sy,
  };
  for (const child of group.children || []) {
    if (SHAPE_TAGS.has(tagName(child))) drawShape(child, host, inner, doc, opts);
  }
}

function shapeBox(shape, style, map) {
  const left = absLength(style["margin-left"]) || absLength(style.left) || 0;
  const top = absLength(style["margin-top"]) || absLength(style.top) || 0;
  const width = absLength(style.width);
  const height = absLength(style.height);
  return {
    left: left * map.sx + map.ox,
    top: top * map.sy + map.oy,
    width: width * map.sx,
    height: height * map.sy,
  };
}

function vmlFill(shape) {
  if (attr(shape, "filled") === "f") return null;
  const fill = firstOf(shape, "fill");
  if (fill && (attr(fill, "opacity") === "0" || attr(fill, "type") === "none")) return null;
  const color = attr(shape, "fillcolor") || (fill && attr(fill, "color")) || null;
  if (!color) return null;
  return { type: "solid", color: cssColor(color) };
}

function vmlStroke(shape) {
  if (attr(shape, "stroked") === "f") return null;
  const stroke = firstOf(shape, "stroke");
  const color = attr(shape, "strokecolor") || (stroke && attr(stroke, "color")) || null;
  const weight = attr(shape, "strokeweight") || (stroke && attr(stroke, "weight")) || null;
  const dash = (stroke && attr(stroke, "dashstyle")) || null;
  if (!color && !stroke) return null;
  return {
    color: cssColor(color || "#000000") || "#000000",
    widthPx: absLength(weight) || 1,
    dash: vmlDash(dash),
  };
}

function vmlDash(dash) {
  switch (String(dash || "").toLowerCase()) {
    case "dash": return [4, 3];
    case "dot": return [1, 2];
    case "dashdot": return [4, 2, 1, 2];
    case "longdash": return [8, 3];
    case "sysdash": return [3, 1];
    default: return null;
  }
}

function applyVmlStroke(el, stroke) {
  if (!stroke) {
    el.setAttribute("stroke", "none");
    return;
  }
  if (stroke.dash) {
    el.setAttribute("stroke", stroke.color);
    el.setAttribute("stroke-width", String(Math.max(0.5, stroke.widthPx)));
    el.setAttribute("stroke-dasharray", stroke.dash.map((v) => round1(v * stroke.widthPx)).join(" "));
    return;
  }
  applySvgStroke(el, { color: stroke.color }, stroke.widthPx);
}

function drawOutline(host, tag, box, fill, stroke, doc) {
  const svg = makeSvg(doc, Math.max(box.width, 1), Math.max(box.height, 1));
  const rect = svgEl(doc, "rect");
  rect.setAttribute("x", "0");
  rect.setAttribute("y", "0");
  rect.setAttribute("width", String(box.width));
  rect.setAttribute("height", String(box.height));
  rect.setAttribute("fill", "none");
  applyVmlStroke(rect, stroke);
  svg.appendChild(rect);
  host.appendChild(svg);
}

// A line child of a group: its endpoints are in the group's coordinate space,
// not a style box, so the host box is their mapped bounding box and the line
// keeps its direction inside it.
function drawChildLine(container, shape, map, stroke, doc) {
  const from = String(attr(shape, "from") || "0,0").split(",").map(Number);
  const to = String(attr(shape, "to") || "0,0").split(",").map(Number);
  const x1 = (from[0] || 0) * map.sx + map.ox;
  const y1 = (from[1] || 0) * map.sy + map.oy;
  const x2 = (to[0] || 0) * map.sx + map.ox;
  const y2 = (to[1] || 0) * map.sy + map.oy;
  const host = doc.createElement("div");
  host.className = "ov-docx-shape";
  host.style.left = Math.min(x1, x2) + "px";
  host.style.top = Math.min(y1, y2) + "px";
  host.style.width = Math.abs(x2 - x1) + "px";
  host.style.height = Math.abs(y2 - y1) + "px";
  container.appendChild(host);
  drawLineBox(host, x1, y1, x2, y2, stroke, doc);
}

// Draws a line inside a host whose origin is the line's bounding box.
function drawLineBox(host, x1, y1, x2, y2, stroke, doc) {
  const width = Math.max(Math.abs(x2 - x1), 1);
  const height = Math.max(Math.abs(y2 - y1), 1);
  const ox = Math.min(x1, x2);
  const oy = Math.min(y1, y2);
  const svg = makeSvg(doc, width, height);
  svg.style.left = "0";
  svg.style.top = "0";
  const line = svgEl(doc, "line");
  line.setAttribute("x1", String(x1 - ox));
  line.setAttribute("y1", String(y1 - oy));
  line.setAttribute("x2", String(x2 - ox));
  line.setAttribute("y2", String(y2 - oy));
  line.setAttribute("fill", "none");
  applyVmlStroke(line, stroke || { color: "#000000", widthPx: 1 });
  svg.appendChild(line);
  host.appendChild(svg);
}

// The shapetype behind a straight connector (Word writes t32) is one
// dimensional: a v:shape with that type draws a line across its own box.
function vmlOneDimensional(shape) {
  const type = String(attr(shape, "type") || "");
  if (attr(shape, "oned") === "t") return true;
  if (/t32|t33|t34|t35/.test(type)) return true;
  const shapetype = firstOf(shape, "shapetype");
  return Boolean(shapetype && attr(shapetype, "oned") === "t");
}

function drawPolyline(host, shape, box, stroke, doc) {
  const points = String(attr(shape, "points") || "").trim().split(/\s+/).filter(Boolean);
  if (!points.length) return;
  const svg = makeSvg(doc, Math.max(box.width, 1), Math.max(box.height, 1));
  const poly = svgEl(doc, "polyline");
  poly.setAttribute("points", points.join(" "));
  poly.setAttribute("fill", "none");
  applyVmlStroke(poly, stroke || { color: "#000000", widthPx: 1 });
  svg.appendChild(poly);
  host.appendChild(svg);
}

function roundRectRadius(shape, box) {
  const arc = absLength(attr(shape, "arcsize")) || 0;
  if (arc) return arc;
  return Math.min(box.width, box.height) * 0.16;
}

// ---------- paths ----------

// VML path syntax: m/l/c/q with relative forms and omitted coordinates. Word
// writes text box shapetypes as a rectangle path, so a shape with no path of
// its own draws its box.
function pathData(shape, width, height) {
  const path = attr(shape, "path");
  if (!path || path === "m,l,21600r21600,l21600,xe") return null;
  const coordsize = String(attr(shape, "coordsize") || "").split(",");
  const sizeX = parseFloat(coordsize[0]) || 21600;
  const sizeY = parseFloat(coordsize[1]) || 21600;
  const sx = width > 0 ? width / sizeX : 1;
  const sy = height > 0 ? height / sizeY : 1;
  const tokens = String(path).match(/[a-zA-Z]|-?\d+(?:\.\d+)?/g) || [];
  const parts = [];
  let command = "l";
  let index = 0;
  let cx = 0;
  let cy = 0;
  let startX = 0;
  let startY = 0;
  const number = () => {
    const token = tokens[index];
    if (token == null || /[a-zA-Z]/.test(token)) return null;
    index++;
    return parseFloat(token);
  };
  while (index < tokens.length) {
    const token = tokens[index];
    if (/[a-zA-Z]/.test(token)) {
      command = token.toLowerCase();
      index++;
      if (command === "x" || command === "e") {
        parts.push("Z");
        continue;
      }
      if (command === "m") {
        const x = number();
        const y = number();
        if (x == null || y == null) continue;
        cx = x;
        cy = y;
        startX = x;
        startY = y;
        parts.push("M" + round1(cx * sx) + " " + round1(cy * sy));
        command = "l";
      }
      continue;
    }
    if (command === "l" || command === "r") {
      let x = number();
      let y = number();
      if (x == null) { x = command === "r" ? 0 : cx; }
      if (y == null) { y = command === "r" ? 0 : cy; }
      cx = command === "r" ? cx + x : x;
      cy = command === "r" ? cy + y : y;
      parts.push("L" + round1(cx * sx) + " " + round1(cy * sy));
    } else if (command === "c") {
      const x1 = number();
      const y1 = number();
      const x2 = number();
      const y2 = number();
      const x3 = number();
      const y3 = number();
      if (x1 == null || y1 == null || x2 == null || y2 == null || x3 == null || y3 == null) break;
      parts.push("C" + round1(x1 * sx) + " " + round1(y1 * sy) + " "
        + round1(x2 * sx) + " " + round1(y2 * sy) + " "
        + round1(x3 * sx) + " " + round1(y3 * sy));
      cx = x3;
      cy = y3;
    } else if (command === "q") {
      const x1 = number();
      const y1 = number();
      const x2 = number();
      const y2 = number();
      if (x1 == null || y1 == null || x2 == null || y2 == null) break;
      parts.push("Q" + round1(x1 * sx) + " " + round1(y1 * sy) + " "
        + round1(x2 * sx) + " " + round1(y2 * sy));
      cx = x2;
      cy = y2;
    } else {
      index++;
    }
  }
  if (!parts.length) return null;
  return parts.join(" ");
}

// ---------- text and images ----------

function renderVmlText(host, shape, doc, opts) {
  const textbox = firstOf(shape, "textbox");
  if (!textbox) return;
  const content = firstOf(textbox, "txbxContent");
  if (!content) return;
  const box = doc.createElement("div");
  box.className = "ov-docx-vml-text";
  const inset = String(attr(textbox, "inset") || "").split(/[,\s]+/).filter(Boolean);
  if (inset.length >= 4) {
    box.style.paddingLeft = absLength(inset[0] + "pt") + "px";
    box.style.paddingTop = absLength(inset[1] + "pt") + "px";
    box.style.paddingRight = absLength(inset[2] + "pt") + "px";
    box.style.paddingBottom = absLength(inset[3] + "pt") + "px";
  }
  const anchor = /v-text-anchor:\s*(\w+)/.exec(attr(shape, "style") || "");
  const where = anchor ? anchor[1] : "top";
  box.style.display = "flex";
  box.style.flexDirection = "column";
  box.style.justifyContent = where === "middle" || where === "center" ? "center"
    : where === "bottom" ? "flex-end" : "flex-start";
  host.appendChild(box);
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

function resolveImage(imagedata, opts) {
  const rid = attrNs(imagedata, "id") || attr(imagedata, "id") || attrNs(imagedata, "embed");
  if (!rid || !opts.mediaUrl) return null;
  return opts.mediaUrl(rid);
}

module.exports = { parseVmlRun, renderVml };
