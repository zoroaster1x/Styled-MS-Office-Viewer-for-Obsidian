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

// Writes an inventory of what the viewer would draw, as JSON on stdout. The
// Python side (test/compare-reference.py) builds the same inventory from a
// reference conversion and diffs the two.
//
//   bun test/inventory.mjs <file> > mine.json

import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { setupDom, createContainer } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);

const file = process.argv[2];
if (!file) {
  console.error("usage: bun test/inventory.mjs <file>");
  process.exit(2);
}
const lower = file.toLowerCase();
const bytes = new Uint8Array(readFileSync(file));
const out = { file, kind: null, pages: [], text: "", images: [] };

// A background image occupies the whole slide box.
function backgroundShape(model, slide) {
  return {
    type: "picture",
    xEmu: 0,
    yEmu: 0,
    cxEmu: Math.round(model.widthPx * 9525),
    cyEmu: Math.round(model.heightPx * 9525),
    path: slide.background ? slide.background.path : null,
    background: true,
  };
}

function textOfShapes(shapes, pageIndex, kind) {
  let text = "";
  for (const shape of shapes) {
    if (shape.type === "group") {
      text += textOfShapes(shape.shapes, pageIndex, kind);
      continue;
    }
    if (shape.textInfo) {
      for (const paragraph of shape.textInfo.paragraphs) {
        // Runs of one paragraph belong to one string; a newline between them
        // would split words that the file happens to split across runs.
        let line = "";
        for (const run of paragraph.runs) {
          if (run.type === "field" && String(run.kind || "").indexOf("slidenum") !== -1) {
            line += String((pageIndex || 0) + 1);
          } else if ((run.type === "run" || run.type === "field") && run.text) {
            line += run.text;
          }
        }
        text += line + "\n";
      }
    }
    if (shape.table) {
      for (const row of shape.table.rows) {
        for (const cell of row.cells) {
          if (cell.text) text += textOfShapes([{ textInfo: cell.text }], pageIndex, kind);
        }
      }
    }
  }
  return text;
}

// The extent an image occupies before its frame clips it. A frame that shows
// the middle 50% of an image is drawn at twice the frame's width, centred.
function uncroppedExtent(frame, crop) {
  const left = crop.l || 0;
  const right = crop.r || 0;
  const top = crop.t || 0;
  const bottom = crop.b || 0;
  const fullW = frame.w / Math.max(0.01, 1 - left - right);
  const fullH = frame.h / Math.max(0.01, 1 - top - bottom);
  return {
    x: frame.x - left * fullW,
    y: frame.y - top * fullH,
    w: fullW,
    h: fullH,
  };
}

// The extent of a picture including its border and its shadow, which is what a
// PDF reports as the drawn image rectangle.
function decoratedExtent(rect, shape) {
  const lineWidth = shape.line && shape.line.color !== "none" && shape.line.widthEmu
    ? emuToPx(shape.line.widthEmu)
    : 0;
  const half = lineWidth / 2;
  let x0 = rect.x - half;
  let y0 = rect.y - half;
  let x1 = rect.x + rect.w + half;
  let y1 = rect.y + rect.h + half;
  const shadow = shape.effects && shape.effects.shadow;
  if (shadow) {
    const distance = emuToPx(shadow.distEmu);
    const angle = ((shadow.dir || 45) * Math.PI) / 180;
    const blur = emuToPx(shadow.blurEmu);
    const sx = Math.cos(angle) * distance;
    const sy = Math.sin(angle) * distance;
    x0 = Math.min(x0, rect.x + sx - blur);
    y0 = Math.min(y0, rect.y + sy - blur);
    x1 = Math.max(x1, rect.x + rect.w + sx + blur);
    y1 = Math.max(y1, rect.y + rect.h + sy + blur);
  }
  if (lineWidth === 0 && !shadow) return rect;
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

function emuToPx(v) {
  return Math.round((Number(v || 0) / 9525) * 100) / 100;
}

// Effective geometry in slide coordinates.
//
// A group places its children in a child coordinate space: a child at c maps to
// off + (c - chOff) * ext/chExt. Nested groups compose, innermost first. A
// rotated picture is reported as the axis aligned box around it, because that
// is what a PDF describes.
// Effective geometry for one level of grouping, which is what the renderer
// applies with a CSS transform on the group element. Deeply nested groups with
// their own scaling are approximated rather than composed, so a rect here can
// differ from the pixel result for those decks; the image counts still say
// whether anything is missing.
function effectiveRect(shape, parent) {
  if (!parent) {
    return { x: shape.xEmu, y: shape.yEmu, cx: shape.cxEmu, cy: shape.cyEmu };
  }
  const t = parent.groupTransform || {};
  const sx = t.chExtX ? (t.extX || 0) / t.chExtX : 1;
  const sy = t.chExtY ? (t.extY || 0) / t.chExtY : 1;
  const x = (t.offX || 0) + (shape.xEmu - (t.chOffX || 0)) * sx;
  const y = (t.offY || 0) + (shape.yEmu - (t.chOffY || 0)) * sy;
  return { x, y, cx: shape.cxEmu * sx, cy: shape.cyEmu * sy };
}

function collectPictures(shapes, pageIndex, images, chain) {
  const ancestors = chain || [];
  for (const shape of shapes) {
    if (shape.type === "group") {
      collectPictures(shape.shapes, pageIndex, images, ancestors.concat([shape]));
      continue;
    }
    if (shape.type === "picture") {
      const rect = effectiveRect(shape, ancestors[ancestors.length - 1]);
      // A cropped picture is drawn inside its frame, but a PDF describes the
      // image's own extent before the frame clips it, so report the same thing
      // here or the two sides cannot be compared.
      const crop = shape.crop || null;
      const frame = {
        x: emuToPx(rect.x),
        y: emuToPx(rect.y),
        w: emuToPx(rect.cx),
        h: emuToPx(rect.cy),
      };
      let drawn = crop ? uncroppedExtent(frame, crop) : frame;
      // A PDF's idea of a picture's extent covers its border and shadow too.
      drawn = decoratedExtent(drawn, shape);
      images.push({
        page: pageIndex,
        x: drawn.x,
        y: drawn.y,
        w: drawn.w,
        h: drawn.h,
        frameX: frame.x,
        frameY: frame.y,
        frameW: frame.w,
        frameH: frame.h,
        cropped: Boolean(crop),
        decorated: drawn !== frame,
        vector: Boolean(shape.vector),
        source: shape.path || null,
        hasUrl: Boolean(shape.path),
        inGroup: ancestors.length > 0,
        rotated: Boolean(shape.rotation),
      });
    }
    if (shape.textInfo) {
      for (const paragraph of shape.textInfo.paragraphs) {
        for (const run of paragraph.runs) {
          if (run.type === "run" && run.text) out.text += run.text + "\n";
        }
      }
    }
  }
}

if (lower.endsWith(".pptx") || lower.endsWith(".pptm")) {
  const { parsePptx } = require("../src/pptx/parse.js");
  const model = parsePptx(bytes);
  out.kind = "presentation";
  out.widthPx = model.widthPx;
  out.heightPx = model.heightPx;
  out.pages = model.slides.map((slide, index) => {
    const images = [];
    // Everything the renderer puts on the page counts: the slide's own
    // pictures, the layout and master decorations, and the background image.
    collectPictures(slide.shapes, index, images);
    collectPictures(slide.decorations || [], index, images);
    collectPictures(slide.background && slide.background.type === "image"
      ? [backgroundShape(model, slide)]
      : [], index, images);
    out.images.push(...images);
    const decorationText = textOfShapes(slide.decorations || [], index, "decoration");
    const shapesText = textOfShapes(slide.shapes, index, "shape");
    return {
      index,
      shapeCount: slide.shapes.length + (slide.decorations || []).length,
      imageCount: images.length,
      text: (slide.text || "") + " " + decorationText + " " + shapesText,
      notes: slide.notes || "",
    };
  });
  out.text = model.slides.map((s, i) => (s.text || "") + "\n" + textOfShapes(s.decorations || [], i, "decoration")).join("\n");
} else if (lower.endsWith(".docx") || lower.endsWith(".docm")) {
  const { parseDocx } = require("../src/docx/parse.js");
  const model = parseDocx(bytes);
  out.kind = "document";
  let images = 0;
  const walkBlocks = (blocks) => {
    for (const block of blocks) {
      if (block.type === "table") {
        for (const row of block.rows) for (const cell of row.cells) walkBlocks(cell.blocks);
        continue;
      }
      if (block.type === "p") {
        for (const run of block.runs || []) walkRun(run);
        out.text += "\n";
      }
    }
  };
  const walkRun = (run) => {
    if (!run) return;
    if (run.type === "image") {
      images++;
      out.images.push({ page: 0, source: run.rid || null, hasUrl: Boolean(run.url), w: run.widthPx, h: run.heightPx });
    } else if (run.type === "run") {
      for (const inner of run.runs) walkRun(inner);
    } else if (run.type === "link") {
      for (const inner of run.link.runs) walkRun(inner);
    } else if (run.type === "text") {
      // Runs within a paragraph belong to one string. A newline here would
      // split words that Word happened to split across runs.
      out.text += run.text;
    }
  };
  walkBlocks(model.body);
  out.pages = [{ index: 0, shapeCount: 0, imageCount: images, text: out.text }];
} else {
  console.error("unsupported: " + file);
  process.exit(2);
}

// Round trip through a container so the renderer is exercised too, and count
// the image elements it actually produces.
try {
  const host = createContainer();
  if (out.kind === "presentation") {
    const { parsePptx } = require("../src/pptx/parse.js");
    const { createPptxRenderer } = require("../src/pptx/render.js");
    const model = parsePptx(bytes);
    const renderer = createPptxRenderer({ container: host, model, settings: { fit: "none" } });
    out.renderedImages = 0;
    for (let i = 0; i < model.slides.length; i++) {
      renderer.goToSlide(i);
      out.renderedImages = Math.max(out.renderedImages, host.querySelectorAll("img").length);
    }
    renderer.destroy();
  } else {
    const { parseDocx } = require("../src/docx/parse.js");
    const { createDocxRenderer } = require("../src/docx/render.js");
    const model = parseDocx(bytes);
    let ready = null;
    const renderer = createDocxRenderer({
      container: host,
      model,
      settings: {},
      onReady: (info) => { ready = info; },
    });
    out.renderedImages = host.querySelectorAll("img").length;
    out.renderedMissing = host.querySelectorAll(".is-missing").length;
    out.pageCount = ready ? ready.pageCount : host.querySelectorAll(".ov-docx-page").length;
    out.pageTexts = Array.from(host.querySelectorAll(".ov-docx-page")).map((el) =>
      (el.textContent || "").replace(/\s+/g, " ").trim());
    renderer.destroy();
  }
} catch (err) {
  out.renderError = err.message;
}

console.log(JSON.stringify(out));
