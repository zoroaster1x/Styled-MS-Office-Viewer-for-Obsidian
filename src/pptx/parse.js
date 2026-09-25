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

// Reads a PresentationML deck into the shape tree the renderer draws. Slide
// masters and layouts are folded in so inherited placeholder geometry, fills
// and text styles resolve the way PowerPoint shows them.

const { Package, relTypeIs } = require("../shared/package");
const { parseMath, mathText } = require("../docx/math");
const {
  childrenOf, firstOf, findFirst, attr, attrInt, attrBool, flag, tagName,
} = require("../shared/xml");
const { parseTheme, colorFromContainer, themeColor } = require("../shared/color");
const { createMediaCache, mediaUrl } = require("../media/media");

function parsePptx(input) {
  const pkg = Package.open(input);
  const presentationPath = findPresentationPart(pkg);
  if (!presentationPath) {
    throw new Error("This file has no ppt/presentation.xml part, so it is not a readable presentation.");
  }
  const presentationXml = pkg.xml(presentationPath);
  if (!presentationXml) throw new Error("ppt/presentation.xml could not be parsed.");
  const root = presentationXml.documentElement;
  const rels = pkg.rels(presentationPath);
  const mediaCache = createMediaCache();

  const slideSize = firstOf(root, "sldSz");
  const widthPx = slideSize ? attrInt(slideSize, "cx", 9144000) / 9525 : 960;
  const heightPx = slideSize ? attrInt(slideSize, "cy", 6858000) / 9525 : 720;

  const themes = parseThemes(pkg);
  const masters = parseMasters(pkg, mediaCache);
  const layouts = parseLayouts(pkg, masters);

  const slides = [];
  const slideList = firstOf(root, "sldIdLst");
  const slideIds = slideList ? childrenOf(slideList, "sldId") : [];
  slideIds.forEach((slideId, index) => {
    // <p:sldId id="404" r:id="rId2"/> carries the slide number in id and the
    // relationship in r:id, so the namespaced attribute has to win.
    const rid = attr(slideId, "r:id") || attr(slideId, "id");
    const rel = rels.get(rid);
    if (!rel || !relTypeIs(rel, "slide")) return;
    const path = pkg.resolve(presentationPath, rel);
    if (!path) return;
    const parsed = parseSlide(pkg, path, {
      index,
      layout: findLayoutFor(pkg, path, layouts),
      masters,
      themes,
      mediaCache,
    });
    if (slides.length < MAX_SLIDES) slides.push(parsed);
  });

  const model = {
    kind: "pptx",
    pkg,
    widthPx,
    heightPx,
    slides,
    masters,
    layouts,
    themes,
    mediaCache,
    embeddedFonts: parseEmbeddedFonts(root, rels, presentationPath, pkg),
    notesSize: parseNotesSize(root),
    properties: parseProperties(pkg),
    slideSize,
    // Media is inflated and encoded on first draw, not at open.
    mediaUrl(partPath) {
      if (!partPath) return null;
      return mediaUrl(pkg, partPath, mediaCache);
    },
  };
  return model;
}

function findPresentationPart(pkg) {
  const rels = pkg.rels("");
  for (const rel of rels.values()) {
    if (relTypeIs(rel, "officeDocument")) {
      const path = pkg.resolve("", rel);
      if (path && path.indexOf("ppt/") === 0) return path;
    }
  }
  if (pkg.has("ppt/presentation.xml")) return "ppt/presentation.xml";
  return null;
}

function parseNotesSize(root) {
  const notesSz = firstOf(root, "notesSz");
  return {
    widthPx: notesSz ? attrInt(notesSz, "cx", 6858000) / 9525 : 720,
    heightPx: notesSz ? attrInt(notesSz, "cy", 9144000) / 9525 : 960,
  };
}

// ---------- embedded fonts ----------

// A presentation can carry its own fonts as EOT containers under ppt/fonts.
// The list maps a typeface name and a style to a part; the slides only ever
// name the typeface.
const EMBEDDED_FONT_STYLES = [
  ["regular", "400", "normal"],
  ["bold", "700", "normal"],
  ["italic", "400", "italic"],
  ["boldItalic", "700", "italic"],
];

function parseEmbeddedFonts(root, rels, presentationPath, pkg) {
  const out = [];
  const list = firstOf(root, "embeddedFontLst");
  for (const entry of childrenOf(list, "embeddedFont")) {
    const font = firstOf(entry, "font");
    const family = font ? attr(font, "typeface") : null;
    if (!family) continue;
    for (const [tag, weight, style] of EMBEDDED_FONT_STYLES) {
      const ref = firstOf(entry, tag);
      if (!ref) continue;
      const rid = attr(ref, "r:id") || attr(ref, "id");
      const rel = rid ? rels.get(rid) : null;
      if (!rel || !relTypeIs(rel, "font")) continue;
      const path = pkg.resolve(presentationPath, rel);
      if (path && pkg.has(path)) out.push({ family, weight, style, tag, path });
      if (out.length >= 32) return out;
    }
  }
  return out;
}

// ---------- themes ----------

// Theme parts are looked up by path. Each slide master points at exactly one
// theme, and layouts and slides inherit it through their master.
function parseThemes(pkg) {
  const out = new Map();
  for (const name of pkg.list()) {
    if (!/^ppt\/theme\/theme\d+\.xml$/.test(name)) continue;
    const doc = pkg.xml(name);
    if (!doc) continue;
    out.set(name, { path: name, theme: parseTheme(doc) });
  }
  return out;
}

// ---------- masters ----------

function parseMasters(pkg, mediaCache) {
  const out = new Map();
  for (const name of pkg.list()) {
    if (!/^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(name)) continue;
    const doc = pkg.xml(name);
    if (!doc) continue;
    const root = doc.documentElement;
    const clrMap = parseColorMap(firstOf(root, "clrMap"));
    const tree = firstOf(root, "cSld") ? firstOf(firstOf(root, "cSld"), "spTree") : null;
    const themePath = findThemeFor(pkg, name);
    const master = {
      path: name,
      clrMap,
      themePath,
      theme: themePath ? parseTheme(pkg.xml(themePath)) : null,
      background: parseBackground(root, null, {
        pkg,
        partPath: name,
        rels: pkg.rels(name),
        theme: themePath ? parseTheme(pkg.xml(themePath)) : null,
      }),
      shapes: tree ? parseShapeTree(tree, { pkg, partPath: name, mediaCache, master: null, layout: null, theme: null }) : [],
    };
    const txStyles = firstOf(root, "txStyles");
    master.textStyles = txStyles ? {
      title: parseListStyle(firstOf(txStyles, "titleStyle")),
      body: parseListStyle(firstOf(txStyles, "bodyStyle")),
      other: parseListStyle(firstOf(txStyles, "otherStyle")),
    } : null;
    out.set(name, master);
  }
  return out;
}

function findThemeFor(pkg, partPath) {
  const rels = pkg.rels(partPath);
  for (const rel of rels.values()) {
    if (!relTypeIs(rel, "theme")) continue;
    const path = pkg.resolve(partPath, rel);
    if (path) return path;
  }
  return null;
}

function parseColorMap(el) {
  const map = {};
  if (!el) {
    // The default mapping PowerPoint applies when a master omits clrMap.
    return { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2", accent1: "accent1", accent2: "accent2", accent3: "accent3", accent4: "accent4", accent5: "accent5", accent6: "accent6", hlink: "hlink", folHlink: "folHlink" };
  }
  for (const key of ["bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]) {
    map[key] = attr(el, key) || key;
  }
  return map;
}

// ---------- layouts ----------

function parseLayouts(pkg, masters) {
  const out = new Map();
  for (const name of pkg.list()) {
    if (!/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name)) continue;
    const doc = pkg.xml(name);
    if (!doc) continue;
    const root = doc.documentElement;
    const masterPath = findMasterFor(pkg, name);
    const master = masterPath ? masters.get(masterPath) : null;
    const tree = firstOf(root, "cSld") ? firstOf(firstOf(root, "cSld"), "spTree") : null;
    const clrMapOvr = firstOf(root, "clrMapOvr");
    out.set(name, {
      path: name,
      type: attr(root, "type"),
      masterPath,
      master,
      clrMap: master ? master.clrMap : null,
      theme: master ? master.theme : null,
      themePath: master ? master.themePath : null,
      background: parseBackground(root, master, {
        pkg,
        partPath: name,
        rels: pkg.rels(name),
        theme: master ? master.theme : null,
      }),
      shapes: tree ? parseShapeTree(tree, { pkg, partPath: name, mediaCache: null, master, layout: null, theme: master ? master.theme : null, rels: pkg.rels(name), clrMap: master ? master.clrMap : null }) : [],
      placeholderMap: buildPlaceholderMap(tree),
    });
  }
  return out;
}

function findMasterFor(pkg, layoutPath) {
  const rels = pkg.rels(layoutPath);
  for (const rel of rels.values()) {
    if (!relTypeIs(rel, "slideMaster")) continue;
    const path = pkg.resolve(layoutPath, rel);
    if (path) return path;
  }
  return null;
}

function findLayoutFor(pkg, slidePath, layouts) {
  const rels = pkg.rels(slidePath);
  for (const rel of rels.values()) {
    if (!relTypeIs(rel, "slideLayout")) continue;
    const path = pkg.resolve(slidePath, rel);
    if (path && layouts.has(path)) return layouts.get(path);
  }
  return null;
}

function buildPlaceholderMap(tree) {
  const map = new Map();
  if (!tree) return map;
  const walk = (node) => {
    for (const child of node.children || []) {
      const tag = tagName(child);
      if (tag === "sp") {
        const nv = firstOf(child, "nvSpPr");
        const ph = findFirst(nv, "ph");
        if (ph) {
          const key = placeholderKey(attr(ph, "type") || "body", attrInt(ph, "idx", 0));
          map.set(key, child);
        }
      }
      if (tag === "pic" || tag === "graphicFrame" || tag === "grpSp" || tag === "cxnSp") {
        const nv = firstOf(child, tag === "grpSp" ? "nvGrpSpPr" : tag === "graphicFrame" ? "nvGraphicFramePr" : tag === "pic" ? "nvPicPr" : "nvCxnSpPr");
        const ph = findFirst(nv, "ph");
        if (ph) map.set(placeholderKey(attr(ph, "type") || "body", attrInt(ph, "idx", 0)), child);
      }
      walk(child);
    }
  };
  walk(tree);
  return map;
}

function placeholderKey(type, idx) {
  return String(type) + ":" + idx;
}

// ---------- slides ----------

function parseSlide(pkg, path, env) {
  const doc = pkg.xml(path);
  if (!doc) throw new Error("Slide " + path + " could not be parsed.");
  const root = doc.documentElement;
  const cSld = firstOf(root, "cSld");
  const tree = cSld ? firstOf(cSld, "spTree") : null;
  const rels = pkg.rels(path);
  const layout = env.layout;
  const master = layout && layout.master ? layout.master : (layout && layout.masterPath ? env.masters.get(layout.masterPath) : null);
  const theme = (layout && layout.theme) || (master && master.theme) || null;
  const clrMap = (layout && layout.clrMap) || (master && master.clrMap) || null;
  // The slide's diagram drawings, in relationship order, so a SmartArt frame
  // can claim the drawing that belongs to it.
  const diagramDrawings = [];
  for (const rel of rels.values()) {
    if (!relTypeIs(rel, "diagramDrawing")) continue;
    const target = pkg.resolve(path, rel);
    if (target && pkg.has(target)) diagramDrawings.push(target);
  }
  const ctx = {
    pkg,
    partPath: path,
    mediaCache: env.mediaCache,
    master,
    layout,
    theme,
    clrMap,
    rels,
    diagramDrawings,
    diagramIndex: 0,
    slide: null,
  };
  const shapes = tree ? parseShapeTree(tree, ctx) : [];
  const notes = parseNotes(pkg, path, env);
  const textParts = [];
  collectShapeText(shapes, textParts);
  let imageCount = 0;
  countImages(shapes, (n) => { imageCount += n; });
  const background = parseBackground(root, layout || master, {
    pkg,
    partPath: path,
    rels,
    theme,
  });
  return {
    index: env.index,
    path,
    shapes,
    decorations: buildDecorations(layout, master),
    background,
    clrMap,
    theme,
    layoutPath: layout ? layout.path : null,
    masterPath: master ? master.path : null,
    notes,
    text: textParts.join(" ").replace(/\s+/g, " ").trim(),
    title: findTitle(shapes),
    imageCount,
    hidden: attr(root, "show") === "0",
  };
}

// Shapes from the layout and the master that a slide always shows: logos,
// rules, frames. Every placeholder is a prototype for the slide's own content,
// not content: a date, footer or slide number appears only when the slide
// carries that placeholder itself. The deck's own PowerPoint thumbnail,
// The reference renderers all leave an uninstantiated date and slide number
// off the slide, and drawing the layout's put a ghost date on every page.

// Bounds so a corrupt or crafted deck cannot exhaust memory or the stack.
const MAX_SLIDES = 2000;
const MAX_SHAPES_PER_TREE = 5000;
const MAX_GROUP_DEPTH = 16;
const MAX_TABLE_ROWS = 5000;
const MAX_TABLE_CELLS = 500;

function buildDecorations(layout, master) {
  const furniture = [];
  const collect = (shapes) => {
    for (const shape of shapes || []) {
      // A master or layout shape with no placeholder is design furniture.
      if (!shape.placeholder) furniture.push(shape);
    }
  };
  if (master) collect(master.shapes);
  if (layout) collect(layout.shapes);
  return furniture;
}

function findTitle(shapes) {
  for (const shape of shapes) {
    if (shape.placeholder && (shape.placeholder.type === "title" || shape.placeholder.type === "ctrTitle")) {
      return shape.text || "";
    }
    if (shape.type === "group") {
      const inner = findTitle(shape.shapes);
      if (inner) return inner;
    }
  }
  return "";
}

function collectShapeText(shapes, out) {
  for (const shape of shapes) {
    if (shape.text) out.push(shape.text);
    if (shape.type === "group") collectShapeText(shape.shapes, out);
    // Table cells carry text of their own, and search should reach it.
    if (shape.table) {
      for (const row of shape.table.rows) {
        for (const cell of row.cells) {
          if (cell.text && cell.text.text) out.push(cell.text.text);
        }
      }
    }
  }
}

function countImages(shapes, add) {
  for (const shape of shapes) {
    if (shape.type === "picture") add(1);
    else if (shape.type === "group") countImages(shape.shapes, add);
  }
}

function parseNotes(pkg, slidePath, env) {
  const rels = pkg.rels(slidePath);
  for (const rel of rels.values()) {
    if (!relTypeIs(rel, "notesSlide")) continue;
    const path = pkg.resolve(slidePath, rel);
    const doc = pkg.xml(path);
    if (!doc) continue;
    const tree = firstOf(firstOf(doc.documentElement, "cSld"), "spTree");
    if (!tree) continue;
    const ctx = {
      pkg,
      partPath: path,
      mediaCache: env.mediaCache,
      master: null,
      layout: null,
      theme: env.themes ? null : null,
      clrMap: null,
      rels: pkg.rels(path),
    };
    const shapes = parseShapeTree(tree, ctx);
    const parts = [];
    collectShapeText(shapes, parts);
    return parts.join("\n").trim();
  }
  return "";
}

// ---------- shape tree ----------

function parseShapeTree(tree, ctx) {
  const out = [];
  // Document order is the z-order. Collecting by tag type draws every picture
  // behind every shape, so an arrow that should point across a picture ends up
  // underneath it.
  for (const el of tree.children || []) {
    if (out.length >= MAX_SHAPES_PER_TREE) break;
    const shape = parseShapeElement(el, ctx, 0);
    if (shape) out.push(shape);
  }
  return out;
}

// One element of a shape tree. PowerPoint wraps a newer construct in
// mc:AlternateContent: the mc:Choice branch is the one to read, and the
// mc:Fallback is a picture of the same thing for older viewers. Reading both
// drew the shape twice and dropped the choice entirely, which is how a slide's
// equations became empty boxes.
function parseShapeElement(el, ctx, depth) {
  const tag = tagName(el);
  if (tag === "sp") return parseShape(el, ctx);
  if (tag === "pic") return parsePicture(el, ctx);
  if (tag === "graphicFrame") return parseGraphicFrame(el, ctx);
  if (tag === "cxnSp") return parseConnector(el, ctx);
  if (tag === "grpSp") return parseGroup(el, ctx, depth || 0);
  if (tag === "AlternateContent") {
    const branch = firstOf(el, "Choice") || firstOf(el, "Fallback");
    for (const inner of (branch && branch.children) || []) {
      const shape = parseShapeElement(inner, ctx, depth);
      if (shape) return shape;
    }
  }
  return null;
}

// The property holder that carries the shape name and placeholder reference.
const NON_VISUAL_CONTAINER = {
  sp: "nvSpPr",
  pic: "nvPicPr",
  graphicFrame: "nvGraphicFramePr",
  grpSp: "nvGrpSpPr",
  cxnSp: "nvCxnSpPr",
};

function nonVisual(el, kind) {
  const container = NON_VISUAL_CONTAINER[kind];
  return container ? firstOf(el, container) : null;
}

function parsePlaceholder(el, kind) {
  const nv = nonVisual(el, kind);
  const ph = nv ? findFirst(nv, "ph") : null;
  if (!ph) return null;
  return {
    type: attr(ph, "type") || "body",
    idx: attrInt(ph, "idx", 0),
    size: attr(ph, "sz"),
    orient: attr(ph, "orient"),
  };
}

function parseShape(el, ctx) {
  const placeholder = parsePlaceholder(el, "sp");
  const shape = {
    type: "shape",
    name: shapeName(el, "sp"),
    placeholder,
    ...parseCommonShape(el, ctx),
  };
  const txBody = firstOf(el, "txBody");
  shape.textInfo = txBody ? parseTextBody(txBody, ctx, shape) : null;
  shape.text = shape.textInfo ? shape.textInfo.text : "";
  shape.geometry = parsePresetGeometry(el);
  return shape;
}

function shapeName(el, kind) {
  const nv = nonVisual(el, kind);
  if (!nv) return "";
  const cNvPr = findFirst(nv, "cNvPr");
  return cNvPr ? attr(cNvPr, "name") || "" : "";
}

// Position and size, falling back to the placeholder on the layout and then
// the master, which is how PowerPoint lays out untouched placeholders.
// <a:xfrm> carries its offsets in <a:off> and <a:ext> children.
function parseCommonShape(el, ctx) {
  // A group keeps its transform in grpSpPr; every other shape keeps it in
  // spPr. Reading spPr for a group finds nothing, which used to send every
  // group down the "zero sized box" path and place its children at chOff.
  const props = firstOf(el, "spPr") || firstOf(el, "grpSpPr");
  const xfrm = props ? firstOf(props, "xfrm") : null;
  const off = xfrm ? firstOf(xfrm, "off") : null;
  const ext = xfrm ? firstOf(xfrm, "ext") : null;
  const geom = xfrm ? {
    x: off ? attrInt(off, "x", null) : null,
    y: off ? attrInt(off, "y", null) : null,
    cx: ext ? attrInt(ext, "cx", null) : null,
    cy: ext ? attrInt(ext, "cy", null) : null,
    rot: attrInt(xfrm, "rot", 0) / 60000,
    flipH: attrBool(xfrm, "flipH", false),
    flipV: attrBool(xfrm, "flipV", false),
  } : {};
  const placeholder = parsePlaceholder(el, shapeKindFor(el));
  const inherited = placeholder ? inheritedPlaceholderBox(ctx, placeholder) : null;
  const final = {
    x: geom.x != null ? geom.x : inherited ? inherited.x : 0,
    y: geom.y != null ? geom.y : inherited ? inherited.y : 0,
    cx: geom.cx != null ? geom.cx : inherited ? inherited.cx : 0,
    cy: geom.cy != null ? geom.cy : inherited ? inherited.cy : 0,
    rot: geom.rot || 0,
    flipH: geom.flipH,
    flipV: geom.flipV,
  };
  const style = firstOf(el, "style");
  return {
    xEmu: final.x,
    yEmu: final.y,
    cxEmu: final.cx,
    cyEmu: final.cy,
    rotation: final.rot,
    flipH: final.flipH,
    flipV: final.flipV,
    fill: parseFill(props, ctx, style),
    line: parseLine(props, ctx, style),
    effects: parseEffects(props),
    placeholderBox: inherited,
  };
}

function shapeKindFor(el) {
  for (const kind of Object.keys(NON_VISUAL_CONTAINER)) {
    if (nonVisual(el, kind)) return kind;
  }
  return "sp";
}

function inheritedPlaceholderBox(ctx, placeholder) {
  if (!ctx || !placeholder) return null;
  const layers = [ctx.layout, ctx.master];
  for (const layer of layers) {
    if (!layer) continue;
    const tree = layer.shapes;
    const found = findPlaceholderShape(tree, placeholder);
    if (found) return { x: found.xEmu, y: found.yEmu, cx: found.cxEmu, cy: found.cyEmu, fill: found.fill, line: found.line, textInfo: found.textInfo };
  }
  return null;
}

function findPlaceholderShape(shapes, placeholder) {
  for (const shape of shapes || []) {
    if (shape.placeholder && (shape.placeholder.type === placeholder.type || (shape.placeholder.idx === placeholder.idx && placeholder.type !== "body"))) {
      return shape;
    }
  }
  return null;
}

function parsePresetGeometry(el) {
  const spPr = firstOf(el, "spPr");
  const geom = spPr ? firstOf(spPr, "prstGeom") : null;
  if (geom) {
    return { preset: attr(geom, "prst") || "rect", custom: null, adjustments: parseAdjustments(geom) };
  }
  const custGeom = spPr ? firstOf(spPr, "custGeom") : null;
  if (custGeom) {
    return { preset: "custom", custom: parseCustGeom(custGeom), adjustments: null };
  }
  return { preset: "rect", custom: null, adjustments: null };
}

// a:avLst holds the shape's adjustment handles. Only literal values are read;
// a handle expressed as a formula keeps its default, which is where the shape
// looks the way the preset defines it.
function parseAdjustments(geom) {
  const avLst = firstOf(geom, "avLst");
  if (!avLst) return null;
  const out = {};
  for (const gd of childrenOf(avLst, "gd")) {
    const name = attr(gd, "name");
    const match = /^\s*val\s+(-?\d+)\s*$/.exec(attr(gd, "fmla") || "");
    if (name && match) out[name] = parseInt(match[1], 10);
  }
  return Object.keys(out).length ? out : null;
}

function parseCustGeom(el) {
  const pathLst = firstOf(el, "pathLst");
  if (!pathLst) return null;
  const paths = [];
  for (const path of childrenOf(pathLst, "path")) {
    const w = attrInt(path, "w", 0);
    const h = attrInt(path, "h", 0);
    const segments = [];
    const fillMode = attr(path, "fill") || null;
    for (const child of path.children || []) {
      const tag = tagName(child);
      if (tag === "moveTo" || tag === "lnTo") {
        const pt = firstOf(child, "pt");
        segments.push({ type: tag === "moveTo" ? "M" : "L", x: pt ? attrInt(pt, "x", 0) : 0, y: pt ? attrInt(pt, "y", 0) : 0 });
      } else if (tag === "cubicBezTo") {
        const pts = childrenOf(child, "pt").map((p) => ({ x: attrInt(p, "x", 0), y: attrInt(p, "y", 0) }));
        if (pts.length === 3) segments.push({ type: "C", pts });
      } else if (tag === "quadBezTo") {
        const pts = childrenOf(child, "pt").map((p) => ({ x: attrInt(p, "x", 0), y: attrInt(p, "y", 0) }));
        if (pts.length === 2) segments.push({ type: "Q", pts });
      } else if (tag === "close") {
        segments.push({ type: "Z" });
      } else if (tag === "arcTo") {
        segments.push({
          type: "A",
          wR: attrInt(child, "wR", 0),
          hR: attrInt(child, "hR", 0),
          stAng: attrInt(child, "stAng", 0),
          swAng: attrInt(child, "swAng", 0),
        });
      }
    }
    paths.push({ w, h, segments, fill: fillMode });
  }
  return paths.length ? paths : null;
}

function parsePicture(el, ctx) {
  const placeholder = parsePlaceholder(el, "pic");
  const common = parseCommonShape(el, ctx);
  const blipFill = firstOf(el, "blipFill");
  const blip = blipFill ? findFirst(blipFill, "blip") : null;
  const rid = blip ? attr(blip, "embed") || attr(blip, "link") : null;
  // Office stores a picture that has a vector version as a raster fallback
  // plus an asvg:svgBlip beside it. PowerPoint and the reference renderers draw the
  // vector, so prefer it and keep the raster as the fallback.
  const svgBlip = blip ? findFirst(blip, "svgBlip") : null;
  const svgRid = svgBlip ? (attr(svgBlip, "embed") || attr(svgBlip, "link")) : null;
  let path = null;
  let rasterPath = null;
  if (rid && ctx.rels) {
    const rel = ctx.rels.get(rid);
    if (rel) rasterPath = ctx.pkg.resolve(ctx.partPath, rel);
  }
  if (svgRid && ctx.rels) {
    const rel = ctx.rels.get(svgRid);
    if (rel) path = ctx.pkg.resolve(ctx.partPath, rel);
  }
  if (!path) path = rasterPath;
  const srcRect = blipFill ? firstOf(blipFill, "srcRect") : null;
  const stretch = blipFill ? firstOf(blipFill, "stretch") : null;
  const fillRect = stretch ? firstOf(stretch, "fillRect") : null;
  const nvPicPr = nonVisual(el, "pic");
  const cNvPr = nvPicPr ? findFirst(nvPicPr, "cNvPr") : null;
  return {
    type: "picture",
    name: shapeName(el, "pic"),
    placeholder,
    ...common,
    path,
    rasterPath,
    vector: Boolean(svgRid),
    alt: cNvPr ? (attr(cNvPr, "descr") || attr(cNvPr, "name") || "") : "",
    crop: srcRect ? {
      l: attrInt(srcRect, "l", 0) / 100000,
      t: attrInt(srcRect, "t", 0) / 100000,
      r: attrInt(srcRect, "r", 0) / 100000,
      b: attrInt(srcRect, "b", 0) / 100000,
    } : null,
    fillMode: fillRect || (firstOf(blipFill, "tile") ? "tile" : "stretch"),
  };
}

function parseConnector(el, ctx) {
  const placeholder = parsePlaceholder(el, "cxnSp");
  return {
    type: "connector",
    name: shapeName(el, "cxnSp"),
    placeholder,
    ...parseCommonShape(el, ctx),
    geometry: parsePresetGeometry(el),
  };
}

function parseGroup(el, ctx, depth) {
  const level = (depth || 0) + 1;
  // Deeply nested groups are a stack overflow waiting to happen.
  if (level > MAX_GROUP_DEPTH) return null;
  const placeholder = parsePlaceholder(el, "grpSp");
  const common = parseCommonShape(el, ctx);
  const xfrm = firstOf(firstOf(el, "grpSpPr"), "xfrm");
  // A group can declare a zero sized box and put the real one in chOff/chExt,
  // which means "identity transform": children keep their own coordinates. The
  // group's frame is then the child box, or the children would all collapse
  // onto the origin.
  if ((!common.cxEmu || !common.cyEmu) && xfrm) {
    const chOffEl = firstOf(xfrm, "chOff");
    const chExtEl = firstOf(xfrm, "chExt");
    const chExtX = attrInt(chExtEl, "cx", 0);
    const chExtY = attrInt(chExtEl, "cy", 0);
    if (chExtX && chExtY) {
      common.xEmu = attrInt(chOffEl, "x", common.xEmu);
      common.yEmu = attrInt(chOffEl, "y", common.yEmu);
      common.cxEmu = chExtX;
      common.cyEmu = chExtY;
    }
  }
  const offEl = xfrm ? firstOf(xfrm, "off") : null;
  const extEl = xfrm ? firstOf(xfrm, "ext") : null;
  const chOffEl = xfrm ? firstOf(xfrm, "chOff") : null;
  const chExtEl = xfrm ? firstOf(xfrm, "chExt") : null;
  const groupTransform = xfrm ? {
    offX: attrInt(offEl, "x", 0),
    offY: attrInt(offEl, "y", 0),
    extX: attrInt(extEl, "cx", 0) || attrInt(chExtEl, "cx", 0),
    extY: attrInt(extEl, "cy", 0) || attrInt(chExtEl, "cy", 0),
    chOffX: attrInt(chOffEl, "x", 0),
    chOffY: attrInt(chOffEl, "y", 0),
    chExtX: attrInt(chExtEl, "cx", 0),
    chExtY: attrInt(chExtEl, "cy", 0),
  } : null;
  const children = [];
  for (const child of el.children || []) {
    if (children.length >= MAX_SHAPES_PER_TREE) break;
    const shape = parseShapeElement(child, ctx, level);
    if (shape) children.push(shape);
  }
  return {
    type: "group",
    name: shapeName(el, "grpSp"),
    placeholder,
    ...common,
    groupTransform,
    shapes: children,
  };
}

function parseGraphicFrame(el, ctx) {
  const placeholder = parsePlaceholder(el, "graphicFrame");
  const common = parseCommonShape(el, ctx);
  const xfrm = firstOf(el, "xfrm");
  const offEl = xfrm ? firstOf(xfrm, "off") : null;
  const extEl = xfrm ? firstOf(xfrm, "ext") : null;
  const offset = {
    x: attrInt(offEl, "x", common.xEmu),
    y: attrInt(offEl, "y", common.yEmu),
    cx: attrInt(extEl, "cx", common.cxEmu),
    cy: attrInt(extEl, "cy", common.cyEmu),
  };
  const graphic = firstOf(el, "graphic");
  const graphicData = graphic ? firstOf(graphic, "graphicData") : null;
  const out = {
    type: "graphicFrame",
    name: shapeName(el, "graphicFrame"),
    placeholder,
    ...common,
    xEmu: offset.x,
    yEmu: offset.y,
    cxEmu: offset.cx,
    cyEmu: offset.cy,
    table: null,
    chart: null,
    diagram: null,
  };
  if (!graphicData) return out;
  const tbl = firstOf(graphicData, "tbl");
  if (tbl) {
    out.table = parseTable(tbl, ctx);
    out.type = "table";
  } else {
    const chart = findFirst(graphicData, "chart");
    if (chart) {
      const rid = attr(chart, "r:id") || attr(chart, "id");
      out.chart = { rid, label: "Chart" };
      out.type = "chart";
    } else if (findFirst(graphicData, "relIds")) {
      out.diagram = { label: "SmartArt diagram" };
      out.type = "diagram";
      // PowerPoint stores the rendered SmartArt as a shape tree in a sibling
      // part (ppt/diagrams/drawing1.xml). Rendering that gives the real shapes
      // and their labels instead of a placeholder. The parts are paired with
      // the frames on the slide in document order, which is how the relIds in
      // the frame and the drawing relationships line up.
      const drawing = takeDiagramDrawing(ctx);
      if (drawing) {
        const shapes = parseDiagramDrawing(drawing, ctx);
        if (shapes && shapes.length) {
          const bounds = shapesBounds(shapes);
          if (bounds && bounds.cx > 0 && bounds.cy > 0) {
            out.type = "group";
            out.shapes = shapes;
            out.groupTransform = {
              offX: offset.x,
              offY: offset.y,
              extX: offset.cx,
              extY: offset.cy,
              chOffX: bounds.x,
              chOffY: bounds.y,
              chExtX: bounds.cx,
              chExtY: bounds.cy,
            };
            out.diagram = { label: "SmartArt diagram", rendered: true };
          }
        }
      }
    }
  }
  return out;
}

// The next unclaimed diagram drawing part for this slide, in relationship
// order. parseSlide fills ctx.diagramDrawings from the slide's rels.
function takeDiagramDrawing(ctx) {
  if (!ctx || !ctx.diagramDrawings || !ctx.diagramDrawings.length) return null;
  const index = ctx.diagramIndex || 0;
  if (index >= ctx.diagramDrawings.length) return null;
  ctx.diagramIndex = index + 1;
  return ctx.diagramDrawings[index];
}

function parseDiagramDrawing(path, ctx) {
  if (!ctx || !ctx.pkg) return null;
  const doc = ctx.pkg.xml(path);
  if (!doc) return null;
  const tree = findFirst(doc.documentElement, "spTree");
  if (!tree) return null;
  const drawingCtx = Object.assign({}, ctx, {
    partPath: path,
    rels: ctx.pkg.rels(path),
    layout: null,
    master: null,
  });
  return parseShapeTree(tree, drawingCtx);
}

// The bounding box of a parsed shape tree, so the drawing can be scaled into
// the frame the slide gives it. Groups are measured by their own box.
function shapesBounds(shapes) {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  const visit = (list) => {
    for (const shape of list || []) {
      if (!shape) continue;
      const x = Number(shape.xEmu) || 0;
      const y = Number(shape.yEmu) || 0;
      const cx = Number(shape.cxEmu) || 0;
      const cy = Number(shape.cyEmu) || 0;
      if (cx > 0 && cy > 0) {
        minX = Math.min(minX, x);
        minY = Math.min(minY, y);
        maxX = Math.max(maxX, x + cx);
        maxY = Math.max(maxY, y + cy);
      }
      if (shape.type === "group") visit(shape.shapes);
    }
  };
  visit(shapes);
  if (!isFinite(minX) || !isFinite(minY) || maxX <= minX || maxY <= minY) return null;
  return { x: minX, y: minY, cx: maxX - minX, cy: maxY - minY };
}

// ---------- text ----------

function parseTextBody(txBody, ctx, shape) {
  const bodyPr = firstOf(txBody, "bodyPr");
  const own = parseListStyle(firstOf(txBody, "lstStyle"));
  const lstStyle = mergeListStyles([own].concat(placeholderListStyles(shape, ctx)));
  const paragraphs = [];
  for (const p of childrenOf(txBody, "p")) {
    paragraphs.push(parseParagraph(p, ctx, lstStyle, shape));
  }
  const text = paragraphs.map((p) => p.text).join("\n");
  // Body properties resolve through the placeholder chain too, so a slide
  // placeholder keeps the master's centre anchor and insets.
  const inheritedBodies = placeholderShapes(shape, ctx)
    .map((placeholder) => placeholder.textInfo && placeholder.textInfo.bodyProps)
    .filter(Boolean);
  const bodyProps = bodyPr ? parseBodyProps(bodyPr, inheritedBodies[0] || null) : (inheritedBodies[0] || null);
  return {
    paragraphs,
    text,
    bodyProps,
    listStyle: lstStyle,
  };
}

function parseBodyProps(bodyPr, fallback) {
  // Insets, wrap, vertical text, autofit and the vertical anchor inherit from
  // the layout placeholder and then the master's. The anchor used to be dropped
  // because centring one deck's title pushed a wrapped line behind the picture
  // that follows; the file is the authority and PowerPoint inherits it.
  const base = fallback || null;
  const baseInsets = base && base.insets ? base.insets : null;
  const out = {
    anchor: attr(bodyPr, "anchor") || (base && base.anchor) || "t",
    wrap: attr(bodyPr, "wrap") || (base ? base.wrap : null) || "square",
    insets: {
      l: attrInt(bodyPr, "lIns", baseInsets ? baseInsets.l : 91440),
      t: attrInt(bodyPr, "tIns", baseInsets ? baseInsets.t : 45720),
      r: attrInt(bodyPr, "rIns", baseInsets ? baseInsets.r : 91440),
      b: attrInt(bodyPr, "bIns", baseInsets ? baseInsets.b : 45720),
    },
    rot: attrInt(bodyPr, "rot", 0),
    vert: attr(bodyPr, "vert") || (base ? base.vert : null) || "horz",
    autofit: null,
  };
  const normAutofit = firstOf(bodyPr, "normAutofit");
  if (normAutofit) out.autofit = { type: "norm", fontScale: attrInt(normAutofit, "fontScale", 100000) / 100000, lnSpcReduction: attrInt(normAutofit, "lnSpcReduction", 0) / 100000 };
  const spAutoFit = firstOf(bodyPr, "spAutoFit");
  if (spAutoFit) out.autofit = { type: "shape" };
  if (firstOf(bodyPr, "noAutofit")) out.autofit = { type: "none" };
  if (!out.autofit && base && base.autofit) out.autofit = base.autofit;
  return out;
}

function parseParagraph(p, ctx, lstStyle, shape) {
  const pPr = firstOf(p, "pPr");
  const out = {
    props: parseParaLevelProps(pPr),
    runs: [],
    text: "",
    endParaRPr: null,
  };
  for (const child of p.children || []) {
    const tag = tagName(child);
    if (tag === "r") out.runs.push(parseTextRun(child, ctx));
    else if (tag === "br") out.runs.push({ type: "br" });
    else if (tag === "m") {
      // An equation sits in an a14:m wrapper; the OMML inside is the same
      // grammar Word writes, so the shared tree and renderer read it.
      const mathEl = findFirst(child, "oMath") || findFirst(child, "oMathPara");
      if (mathEl) {
        const node = parseMath(mathEl);
        if (node) out.runs.push({ type: "math", node });
      }
    }
    else if (tag === "fld") {
      const text = firstOf(child, "t");
      out.runs.push({
        type: "field",
        kind: attr(child, "type") || "text",
        props: parseRunProps(firstOf(child, "rPr"), ctx),
        text: text ? text.textContent : "",
      });
    } else if (tag === "endParaRPr") {
      out.endParaRPr = parseRunProps(child, ctx);
    }
  }
  // Paragraph level from the list style or the placeholder hierarchy.
  out.level = out.props.level || 0;
  out.text = out.runs.map((run) => {
    if (run.type === "br") return "\n";
    if (run.type === "math") return mathText(run.node);
    return run.text || "";
  }).join("");
  out.bullet = resolveBullet(out.props, lstStyle, ctx, shape);
  // marL and indent inherit exactly the way the bullet does: the paragraph's
  // own attributes, then the shape's list style, then the master's style set.
  // A paragraph in a plain text box resolves to the master's otherStyle, whose
  // indent is zero; handing it the body list indent shifts every text box to
  // the right and narrows its wrap until words break letter by letter.
  const level = out.level + 1;
  const listLevel = lstStyle ? lstStyle.levelAt(level) : null;
  const styleSet = masterTextStyle(ctx, shape);
  const masterLevel = styleSet ? styleSet.levelAt(level) : null;
  out.marLEmu = firstGiven(out.props.marginLeftEmu, listLevel && listLevel.marginLeftEmu, masterLevel && masterLevel.marginLeftEmu, 0);
  out.indentEmu = firstGiven(out.props.indentEmu, listLevel && listLevel.indentEmu, masterLevel && masterLevel.indentEmu, 0);
  // Alignment and bullet styling come down the same chain. So does the
  // paragraph's own spacing: a title inherits the master title style's 90% line
  // spacing and a body list its spacing before. Dropping those flattens every
  // deck to single spacing and stacks paragraphs on top of each other.
  out.props.align = out.props.align || (listLevel && listLevel.align) || (masterLevel && masterLevel.align) || null;
  out.props.lineSpacingPct = firstGiven(out.props.lineSpacingPct, listLevel && listLevel.lineSpacingPct, masterLevel && masterLevel.lineSpacingPct, null);
  out.props.lineSpacingPt = firstGiven(out.props.lineSpacingPt, listLevel && listLevel.lineSpacingPt, masterLevel && masterLevel.lineSpacingPt, null);
  out.props.spaceBeforePt = firstGiven(out.props.spaceBeforePt, listLevel && listLevel.spaceBeforePt, masterLevel && masterLevel.spaceBeforePt, null);
  out.props.spaceAfterPt = firstGiven(out.props.spaceAfterPt, listLevel && listLevel.spaceAfterPt, masterLevel && masterLevel.spaceAfterPt, null);
  if (out.props.bulletColor == null && listLevel && listLevel.bulletColor) out.props.bulletColor = listLevel.bulletColor;
  if (out.props.bulletSize == null && listLevel && listLevel.bulletSize) out.props.bulletSize = listLevel.bulletSize;
  // Run defaults for the level fill in whatever a run leaves unset, which is
  // how a level 2 bullet keeps the 20pt the layout gives it.
  const levelRuns = mergeRunProps(listLevel && listLevel.runProps, masterLevel && masterLevel.runProps);
  if (levelRuns) {
    for (const run of out.runs) {
      if (run.type === "run" || run.type === "field") run.props = fillRunProps(run.props, levelRuns);
    }
  }
  return out;
}

// The first value that is neither null nor undefined, with the last argument as
// the fallback. Zero counts as given.
function firstGiven(...values) {
  const fallback = values.pop();
  for (const value of values) if (value != null) return value;
  return fallback;
}

// The layout and master placeholders a shape inherits from, nearest first.
function placeholderShapes(shape, ctx) {
  const out = [];
  if (!shape || !shape.placeholder || !ctx) return out;
  for (const layer of [ctx.layout, ctx.master]) {
    if (!layer) continue;
    const found = findPlaceholderShape(layer.shapes, shape.placeholder);
    if (found) out.push(found);
  }
  return out;
}

function placeholderListStyles(shape, ctx) {
  return placeholderShapes(shape, ctx)
    .map((placeholder) => placeholder.textInfo && placeholder.textInfo.listStyle)
    .filter(Boolean);
}

// Levels come from the shape's own list style, then the layout placeholder's,
// then the master placeholder's, filling gaps level by level. A subtitle keeps
// the layout's centred, bullet free level instead of the body style's bullet.
function mergeListStyles(styles) {
  const present = styles.filter(Boolean);
  if (!present.length) return null;
  const levels = new Map();
  for (let lvl = 1; lvl <= 9; lvl++) {
    const merged = {};
    let found = false;
    for (const style of present) {
      const level = style.levelAt(lvl);
      if (!level) continue;
      found = true;
      for (const [key, value] of Object.entries(level)) {
        // undefined means the source said nothing; null on bullet means buNone,
        // which is a real answer that must win.
        if (value === undefined || (value === null && key !== "bullet")) continue;
        if (!(key in merged)) merged[key] = value;
      }
    }
    if (found) levels.set(lvl, merged);
  }
  if (!levels.size) return null;
  return {
    size: levels.size,
    levelAt(level) {
      return levels.get(Math.max(1, Math.min(9, level))) || null;
    },
  };
}

// Run level defaults, first source winning per property.
function mergeRunProps(...layers) {
  const present = layers.filter(Boolean);
  if (!present.length) return null;
  const out = {};
  for (const layer of present) {
    for (const [key, value] of Object.entries(layer)) {
      if (value == null) continue;
      if (!(key in out)) out[key] = value;
    }
  }
  return out;
}

function fillRunProps(props, defaults) {
  if (!props || !defaults) return props;
  for (const [key, value] of Object.entries(defaults)) {
    if (props[key] == null) props[key] = value;
  }
  return props;
}

function parseParaLevelProps(pPr) {
  if (!pPr) return { level: 0 };
  const out = {
    level: attrInt(pPr, "lvl", 0),
    align: attr(pPr, "algn") || null,
    marginLeftEmu: attrInt(pPr, "marL", null),
    indentEmu: attrInt(pPr, "indent", null),
    rtl: attrBool(pPr, "rtl", false),
  };
  const lnSpc = firstOf(pPr, "lnSpc");
  if (lnSpc) {
    const pct = firstOf(lnSpc, "spcPct");
    const pts = firstOf(lnSpc, "spcPts");
    if (pct) out.lineSpacingPct = attrInt(pct, "val", 100000) / 100000;
    else if (pts) out.lineSpacingPt = attrInt(pts, "val", 0) / 100;
  }
  const spcBef = firstOf(pPr, "spcBef");
  if (spcBef) out.spaceBeforePt = spacingPt(spcBef);
  const spcAft = firstOf(pPr, "spcAft");
  if (spcAft) out.spaceAfterPt = spacingPt(spcAft);
  const buNone = firstOf(pPr, "buNone");
  const buChar = firstOf(pPr, "buChar");
  const buAutoNum = firstOf(pPr, "buAutoNum");
  if (buNone) out.bullet = null;
  else if (buChar) out.bullet = { kind: "char", char: attr(buChar, "char") || "\u2022" };
  else if (buAutoNum) out.bullet = { kind: "auto", type: attr(buAutoNum, "type") || "arabicPeriod", startAt: attrInt(buAutoNum, "startAt", 1) };
  const buClr = firstOf(pPr, "buClr");
  if (buClr) out.bulletColor = colorFromContainer(buClr, null, null);
  const buSzPct = firstOf(pPr, "buSzPct");
  if (buSzPct) out.bulletSize = attrInt(buSzPct, "val", 100000) / 100000;
  const defRPr = firstOf(pPr, "defRPr");
  if (defRPr) out.runProps = parseRunProps(defRPr, null);
  const tabLst = firstOf(pPr, "tabLst");
  if (tabLst) out.tabs = childrenOf(tabLst, "tab").map((t) => attrInt(t, "pos", 0));
  return out;
}

function spacingPt(el) {
  const pct = firstOf(el, "spcPct");
  const pts = firstOf(el, "spcPts");
  if (pct) return { pct: attrInt(pct, "val", 0) / 100000 };
  if (pts) return attrInt(pts, "val", 0) / 100;
  return 0;
}

function parseTextRun(r, ctx) {
  const rPr = firstOf(r, "rPr");
  const t = firstOf(r, "t");
  return {
    type: "run",
    props: parseRunProps(rPr, ctx),
    text: t ? t.textContent || "" : "",
  };
}

function parseRunProps(rPr, ctx) {
  const out = {
    sizePt: null,
    bold: null,
    italic: null,
    underline: null,
    strike: null,
    color: null,
    fontFamily: null,
    fontFamilyTheme: null,
    baseline: null,
    caps: null,
    spacingPt: null,
    highlight: null,
  };
  if (!rPr) return out;
  const sz = attrInt(rPr, "sz", null);
  if (sz != null) out.sizePt = sz / 100;
  const b = attrInt(rPr, "b", null);
  if (b != null) out.bold = b === 1;
  const i = attrInt(rPr, "i", null);
  if (i != null) out.italic = i === 1;
  out.underline = attr(rPr, "u") || null;
  out.strike = attr(rPr, "strike") || null;
  const baseline = attrInt(rPr, "baseline", null);
  if (baseline != null) out.baseline = baseline / 1000;
  const caps = attr(rPr, "cap") || null;
  out.caps = caps;
  const spc = attrInt(rPr, "spc", null);
  if (spc != null) out.spacingPt = spc / 100;
  out.rtl = attrBool(rPr, "rtl", false);
  const solidFill = firstOf(rPr, "solidFill");
  if (solidFill) out.color = colorFromContainer(solidFill, ctxTheme(ctx), null);
  const highlight = firstOf(rPr, "highlight");
  if (highlight) out.highlight = colorFromContainer(highlight, ctxTheme(ctx), null);
  const latin = firstOf(rPr, "latin");
  if (latin) {
    const typeface = attr(latin, "typeface") || null;
    // "+mn-lt" is a theme reference, not a font name: keep it as a token and
    // resolve it against the theme when the run is drawn. Passing it through to
    // CSS gives a browser an invalid family and the default face wins.
    if (typeface && typeface.charAt(0) === "+") out.fontFamilyTheme = typeface;
    else out.fontFamily = typeface;
  }
  const ea = firstOf(rPr, "ea");
  if (ea) out.fontFamilyEastAsia = attr(ea, "typeface") || null;
  return out;
}

// The theme a shape resolves colours against: the slide's own theme, inherited
// from its layout, which inherits from its master.
function ctxTheme(ctx) {
  return ctx && ctx.theme ? ctx.theme : null;
}

// A list style is keyed by level, 1 through 9. The returned object exposes
// levelAt(n) so callers do not care that the source tags carry the number.
function parseListStyle(lstStyle) {
  if (!lstStyle) return null;
  const levels = new Map();
  for (const child of lstStyle.children || []) {
    const match = /^lvl(\d)pPr$/.exec(tagName(child));
    if (match) levels.set(Number(match[1]), parseParaLevelProps(child));
  }
  if (!levels.size) return null;
  return {
    size: levels.size,
    levelAt(level) {
      return levels.get(Math.max(1, Math.min(9, level))) || null;
    },
  };
}

function resolveBullet(props, lstStyle, ctx, shape) {
  if (props.bullet !== undefined) return props.bullet;
  const level = (props.level || 0) + 1;
  const fromList = lstStyle ? lstStyle.levelAt(level) : null;
  if (fromList && fromList.bullet !== undefined) return fromList.bullet;
  const styleSet = masterTextStyle(ctx, shape);
  const fromMaster = styleSet ? styleSet.levelAt(level) : null;
  if (fromMaster && fromMaster.bullet !== undefined) return fromMaster.bullet;
  return null;
}

// Which of the master's three text style sets a placeholder follows. The title
// and the body have sets of their own; everything else, which includes the
// subtitle, the date and the slide number, follows otherStyle.
function masterTextStyle(ctx, shape) {
  const master = ctx && ctx.master;
  if (!master || !master.textStyles) return null;
  const type = shape && shape.placeholder ? shape.placeholder.type : null;
  if (type === "title" || type === "ctrTitle") return master.textStyles.title;
  if (type === "body") return master.textStyles.body;
  return master.textStyles.other;
}

// ---------- tables ----------

function parseTable(tbl, ctx) {
  // Column widths live in <a:tblGrid>, one <a:gridCol> each.
  const grid = [];
  const tblGrid = firstOf(tbl, "tblGrid");
  for (const col of childrenOf(tblGrid, "gridCol")) grid.push(attrInt(col, "w", 0));
  const rows = [];
  for (const tr of childrenOf(tbl, "tr")) {
    if (rows.length >= MAX_TABLE_ROWS) break;
    const cells = [];
    for (const tc of childrenOf(tr, "tc")) {
      if (cells.length >= MAX_TABLE_CELLS) break;
      const txBody = firstOf(tc, "txBody");
      const cell = {
        gridSpan: attrInt(tc, "gridSpan", 1),
        rowSpan: attrInt(tc, "rowSpan", 1),
        hMerge: attrBool(tc, "hMerge", false),
        vMerge: attrBool(tc, "vMerge", false),
        fill: parseFill(firstOf(tc, "tcPr"), ctx, null),
        borders: parseTableBorders(firstOf(tc, "tcPr"), ctx),
        text: txBody ? parseTextBody(txBody, ctx, null) : null,
        margins: parseCellMargins(firstOf(tc, "tcPr")),
        anchor: attrInt(firstOf(tc, "tcPr"), "anchor", null),
      };
      cell.text2 = cell.text ? cell.text.text : "";
      cells.push(cell);
    }
    const height = attrInt(tr, "h", null);
    rows.push({ height, cells });
  }
  return { grid, rows };
}

function parseCellMargins(tcPr) {
  if (!tcPr) return null;
  const out = {};
  for (const side of ["marL", "marR", "marT", "marB"]) {
    const v = attrInt(tcPr, side, null);
    if (v != null) out[side] = v;
  }
  return Object.keys(out).length ? out : null;
}

function parseTableBorders(tcPr, ctx) {
  if (!tcPr) return null;
  const out = {};
  for (const side of ["lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr"]) {
    const ln = firstOf(tcPr, side);
    if (!ln) continue;
    out[side] = parseLineElement(ln, ctx);
  }
  return Object.keys(out).length ? out : null;
}

// ---------- fills, lines, effects ----------

function parseFill(spPr, ctx, style) {
  if (!spPr) {
    const fromStyle = style ? firstOf(style, "fillRef") : null;
    if (fromStyle) {
      return styleRefFill(fromStyle, ctx);
    }
    return null;
  }
  const noFill = firstOf(spPr, "noFill");
  if (noFill) return { type: "none" };
  const solid = firstOf(spPr, "solidFill");
  if (solid) {
    return { type: "solid", color: colorFromContainer(solid, ctxTheme(ctx), "#ffffff") };
  }
  const grad = firstOf(spPr, "gradFill");
  if (grad) return parseGradient(grad, ctx);
  const patt = firstOf(spPr, "pattFill");
  if (patt) {
    return {
      type: "pattern",
      pattern: attr(patt, "prst") || "pct5",
      fg: colorFromContainer(firstOf(patt, "fgClr"), ctxTheme(ctx), "#000000"),
      bg: colorFromContainer(firstOf(patt, "bgClr"), ctxTheme(ctx), "#ffffff"),
    };
  }
  const blip = firstOf(spPr, "blipFill");
  if (blip) return parseBlipFill(blip, ctx);
  const grp = firstOf(spPr, "grpFill");
  if (grp) return { type: "group" };
  if (style) {
    const fillRef = firstOf(style, "fillRef");
    if (fillRef) return styleRefFill(fillRef, ctx);
  }
  return null;
}

function styleRefFill(el, ctx) {
  const idx = attrInt(el, "idx", 0);
  const solid = firstOf(el, "solidFill");
  if (solid) return { type: "solid", color: colorFromContainer(solid, ctxTheme(ctx), "#ffffff") };
  const scheme = firstOf(el, "schemeClr");
  if (scheme) {
    const base = themeColor(ctxTheme(ctx), attr(scheme, "val"), null);
    if (base) return { type: "solid", color: base };
  }
  if (idx === 0) return { type: "none" };
  return null;
}

function parseGradient(grad, ctx) {
  const gsLst = firstOf(grad, "gsLst");
  const stops = [];
  if (gsLst) {
    for (const gs of childrenOf(gsLst, "gs")) {
      const pos = attrInt(gs, "pos", 0) / 100000;
      const color = colorFromContainer(gs, ctxTheme(ctx), "#ffffff");
      stops.push({ pos, color });
    }
  }
  const lin = firstOf(grad, "lin");
  const path = firstOf(grad, "path");
  if (lin) {
    return { type: "gradient", stops, angleDeg: attrInt(lin, "ang", 0) / 60000, scaled: attrBool(lin, "scaled", false) };
  }
  if (path) {
    return { type: "gradient", stops, path: attr(path, "path") || "circle", fillToRect: firstOf(path, "fillToRect") };
  }
  return { type: "gradient", stops, angleDeg: 0 };
}

function parseLine(spPr, ctx, style) {
  if (!spPr) return styleLineFromRef(style, ctx);
  const ln = firstOf(spPr, "ln");
  if (!ln) {
    const fromStyle = style ? firstOf(style, "lnRef") : null;
    return fromStyle ? styleLineFromRef(style, ctx) : null;
  }
  return parseLineElement(ln, ctx);
}

function styleLineFromRef(style, ctx) {
  if (!style) return null;
  const lnRef = firstOf(style, "lnRef");
  if (!lnRef) return null;
  const color = colorFromContainer(lnRef, ctxTheme(ctx), "#000000");
  return { color, widthEmu: 12700, dash: null, cap: null, arrowHead: null, arrowTail: null };
}

function parseLineElement(ln, ctx) {
  const out = {
    color: null,
    widthEmu: attrInt(ln, "w", 12700),
    dash: attr(firstOf(ln, "prstDash"), "val") || null,
    cap: attr(ln, "cap") || null,
    compound: attr(ln, "cmpd") || null,
    arrowHead: null,
    arrowTail: null,
    style: null,
  };
  if (firstOf(ln, "noFill")) out.color = "none";
  else {
    const solid = firstOf(ln, "solidFill");
    if (solid) out.color = colorFromContainer(solid, ctxTheme(ctx), "#000000");
    const grad = firstOf(ln, "gradFill");
    if (grad) {
      const g = parseGradient(grad, ctx);
      out.color = g.stops.length ? g.stops[0].color : "#000000";
    }
    const scheme = firstOf(ln, "schemeClr");
    if (scheme) out.color = themeColor(ctxTheme(ctx), attr(scheme, "val"), null);
    if (firstOf(ln, "pattFill")) out.color = "#000000";
  }
  const head = firstOf(ln, "headEnd");
  const tail = firstOf(ln, "tailEnd");
  if (head) out.arrowHead = { type: attr(head, "type") || "triangle", w: attr(head, "w"), len: attr(head, "len") };
  if (tail) out.arrowTail = { type: attr(tail, "type") || "triangle", w: attr(tail, "w"), len: attr(tail, "len") };
  return out;
}

function parseEffects(spPr) {
  if (!spPr) return null;
  const out = {};
  const outer = firstOf(spPr, "effectLst");
  if (outer) {
    const shadow = firstOf(outer, "outerShdw");
    if (shadow) {
      out.shadow = {
        blurEmu: attrInt(shadow, "blurRad", 0),
        distEmu: attrInt(shadow, "dist", 0),
        dir: attrInt(shadow, "dir", 0) / 60000,
        color: colorFromContainer(shadow, null, "rgba(0,0,0,0.4)"),
      };
    }
  }
  return Object.keys(out).length ? out : null;
}

// A background is a solid, a gradient, or an image. Images resolve through the
// relationships of the part that declares them, so a slide, a layout and a
// master each resolve their own. Text tells us when a slide has no background
// of its own and inherits the layout's, which inherits the master's.
function parseBackground(root, inherited, ctx) {
  // <p:bg> sits inside <p:cSld> on slides, layouts and masters alike.
  const cSld = firstOf(root, "cSld");
  const bgEl = (cSld ? firstOf(cSld, "bg") : null) || firstOf(root, "bg");
  if (!bgEl) {
    if (inherited && inherited.background) return inherited.background;
    return { type: "solid", color: "#ffffff" };
  }
  const bgPr = firstOf(bgEl, "bgPr");
  const bgRef = firstOf(bgEl, "bgRef");
  if (bgPr) {
    const solid = firstOf(bgPr, "solidFill");
    if (solid) return { type: "solid", color: colorFromContainer(solid, ctxTheme(ctx), "#ffffff") };
    const grad = firstOf(bgPr, "gradFill");
    if (grad) return parseGradient(grad, ctx);
    const blip = firstOf(bgPr, "blipFill");
    if (blip) {
      return parseBlipFill(blip, ctx);
    }
  }
  if (bgRef) {
    const idx = attrInt(bgRef, "idx", 0);
    if (idx === 0) return { type: "none" };
    const color = colorFromContainer(bgRef, ctxTheme(ctx), "#ffffff");
    return { type: "solid", color: color || "#ffffff" };
  }
  return { type: "solid", color: "#ffffff" };
}

// A blipFill is an image used as a fill, on a background, a shape or a table
// cell. Stretch maps the image to the frame; tile repeats it; srcRect crops it.
function parseBlipFill(blip, ctx) {
  const blipEl = findFirst(blip, "blip");
  const rid = blipEl ? (attr(blipEl, "embed") || attr(blipEl, "link")) : null;
  let path = null;
  if (rid && ctx && ctx.rels && ctx.pkg && ctx.partPath) {
    const rel = ctx.rels.get(rid);
    if (rel) path = ctx.pkg.resolve(ctx.partPath, rel);
  }
  const srcRect = firstOf(blip, "srcRect");
  return {
    type: "image",
    rid,
    path,
    tile: Boolean(firstOf(blip, "tile")),
    fill: firstOf(blip, "stretch") ? "stretch" : "stretch",
    crop: srcRect ? {
      l: attrInt(srcRect, "l", 0) / 100000,
      t: attrInt(srcRect, "t", 0) / 100000,
      r: attrInt(srcRect, "r", 0) / 100000,
      b: attrInt(srcRect, "b", 0) / 100000,
    } : null,
    tint: firstOf(blip, "lum") ? "lum" : null,
  };
}

function parseProperties(pkg) {
  const core = pkg.xml("docProps/core.xml");
  const app = pkg.xml("docProps/app.xml");
  const out = { title: null, author: null, created: null, application: null, slides: null };
  if (core) {
    for (const child of core.documentElement.children || []) {
      const tag = tagName(child);
      if (tag === "title") out.title = child.textContent;
      else if (tag === "creator") out.author = child.textContent;
      else if (tag === "created") out.created = child.textContent;
    }
  }
  if (app) {
    for (const child of app.documentElement.children || []) {
      const tag = tagName(child);
      if (tag === "Application") out.application = child.textContent;
      else if (tag === "Slides") out.slides = child.textContent;
    }
  }
  return out;
}

module.exports = {
  parsePptx,
  parseColorMap,
  parseFill,
  parseLine,
  parseGradient,
  parseRunProps,
  parseParaLevelProps,
  parseTable,
  placeholderKey,
};
