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

// Draws one slide as a scaled page: shapes as positioned boxes, text laid out
// with the browser's own metrics, pictures as images with their crop applied.
//
// Text layout mirrors PowerPoint closely enough to read: hard breaks, wrapped
// lines measured against the real font, paragraph spacing, bullets with their
// indent, and the shape's vertical anchor.

const { emuToPx, ptToPx } = require("../shared/units");
const { parseHex } = require("../shared/color");
const { presetPolygon, presetPath } = require("./presets");
const { renderMathDom } = require("../shared/math-dom");

const SHAPE_DEFAULT_FONT_PT = 18;

// Metric-compatible fallbacks for the fonts Office documents ask for. Without
// them a slide rendered on a machine without Arial falls back to a wider face
// and the text runs past its shape.
const FONT_FALLBACKS = {
  Arial: 'Arial, "Liberation Sans", Helvetica, sans-serif',
  "Arial Narrow": '"Arial Narrow", "Liberation Sans Narrow", Arial, sans-serif',
  Aptos: 'Aptos, Calibri, Carlito, sans-serif',
  "Aptos Display": 'Aptos Display, Calibri, Carlito, sans-serif',
  Calibri: 'Calibri, Carlito, "Segoe UI", sans-serif',
  // Office titles use Calibri Light, which most Linux installs lack. Falling
  // through to a default sans-serif is much wider and wraps titles that fit in
  // PowerPoint, so the fallback stays metric compatible.
  "Calibri Light": '"Calibri Light", Calibri, Carlito, "Segoe UI", sans-serif',
  Cambria: 'Cambria, "Liberation Serif", serif',
  Candara: 'Candara, Carlito, "Segoe UI", sans-serif',
  "Century Gothic": '"Century Gothic", "URW Gothic", "DejaVu Sans", sans-serif',
  "Comic Sans MS": '"Comic Sans MS", "Comic Neue", cursive',
  Consolas: 'Consolas, "Liberation Mono", monospace',
  Constantia: 'Constantia, "Liberation Serif", serif',
  Corbel: 'Corbel, "Segoe UI", "Noto Sans", sans-serif',
  "Courier New": '"Courier New", "Liberation Mono", monospace',
  "Franklin Gothic Book": '"Franklin Gothic Book", "Liberation Sans", Arial, sans-serif',
  Garamond: 'Garamond, "EB Garamond", "Liberation Serif", serif',
  Georgia: 'Georgia, "Liberation Serif", serif',
  "Gill Sans MT": '"Gill Sans MT", "Liberation Sans", Arial, sans-serif',
  "Lucida Console": '"Lucida Console", "DejaVu Sans Mono", monospace',
  "Palatino Linotype": '"Palatino Linotype", "Book Antiqua", "URW Palladio L", "Liberation Serif", serif',
  "Segoe UI": '"Segoe UI", "Noto Sans", sans-serif',
  Tahoma: 'Tahoma, Verdana, "DejaVu Sans", sans-serif',
  "Times New Roman": '"Times New Roman", "Liberation Serif", serif',
  "Trebuchet MS": '"Trebuchet MS", "Liberation Sans", sans-serif',
  Verdana: 'Verdana, "DejaVu Sans", sans-serif',
  Wingdings: 'Wingdings, "Zapf Dingbats", sans-serif',
};

function fontFamilyCss(name, slide, themeToken) {
  const resolved = themeFontName(themeToken, slide) || name;
  if (!resolved) return 'Calibri, Carlito, "Segoe UI", sans-serif';
  if (FONT_FALLBACKS[resolved]) return FONT_FALLBACKS[resolved];
  if (/^[A-Za-z0-9 ]+$/.test(resolved)) return "\"" + resolved + "\", sans-serif";
  return resolved;
}

// "+mj-lt" is the major theme font and "+mn-lt" the minor one; the East Asian
// and complex script slots fall back to the latin face, which the theme's
// latin entries describe more usefully than the default stack.
function themeFontName(token, slide) {
  if (!token || token.charAt(0) !== "+") return null;
  const fonts = slide && slide.theme ? slide.theme.fonts : null;
  if (!fonts) return null;
  if (token.indexOf("+mj") === 0) return fonts.major || null;
  if (token.indexOf("+mn") === 0) return fonts.minor || null;
  return null;
}
const DEFAULT_INSETS = { l: 91440, t: 45720, r: 91440, b: 45720 };
const MAX_AUTOFIT_SCALE = 1;
const MIN_AUTOFIT_SCALE = 0.45;

function createPptxRenderer(opts) {
  const container = opts.container;
  const model = opts.model;
  const doc = container.ownerDocument || document;
  const settings = Object.assign({
    zoom: 1,
    fit: "contain",
    showNotes: true,
    slideIndex: 0,
    theme: "white",
  }, opts.settings || {});
  const callbacks = {
    onSlideChange: opts.onSlideChange || function () {},
    onReady: opts.onReady || function () {},
    onExternalLink: opts.onExternalLink || function (href) { window.open(href, "_blank"); },
  };

  let destroyed = false;
  let rootEl = null;
  let railEl = null;
  let stageEl = null;
  let slideEl = null;
  let notesEl = null;
  let scale = settings.zoom;
  let current = clampSlide(settings.slideIndex);
  let resizeObserver = null;
  let resizePending = false;
  const thumbnailCache = new Map();
  const mediaCache = model.mediaCache || new Map();
  const fontMetrics = createFontMetrics(doc);

  const slideWidth = model.widthPx;
  const slideHeight = model.heightPx;

  function clampSlide(index) {
    const count = model.slides.length;
    if (!count) return 0;
    return Math.max(0, Math.min(count - 1, index));
  }

  function render() {
    rootEl = container.createDiv("ov-pptx");
    rootEl.style.setProperty("--ov-pptx-scale", String(scale));
    const body = rootEl.createDiv("ov-pptx-body");
    railEl = body.createDiv("ov-pptx-rail");
    const stageWrap = body.createDiv("ov-pptx-stagewrap");
    stageEl = stageWrap.createDiv("ov-pptx-stage");
    notesEl = rootEl.createDiv("ov-pptx-notes");
    watchStageSize();
    buildRail();
    showSlide(current);
    callbacks.onReady({ slideCount: model.slides.length });
  }

  function buildRail() {
    railEl.empty();
    model.slides.forEach((slide, index) => {
      const button = railEl.createDiv("ov-pptx-thumb");
      button.dataset.index = String(index);
      const number = button.createDiv("ov-pptx-thumbnum");
      number.setText(String(index + 1));
      const frame = button.createDiv("ov-pptx-thumbbox");
      frame.style.aspectRatio = slideWidth + " / " + slideHeight;
      button.addEventListener("click", () => showSlide(index));
      thumbnailObserver.observe(button);
    });
    railEl.addEventListener("scroll", () => {
      scheduleRailRender();
    });
    paintVisibleThumbs();
  }

  let railPending = false;
  function scheduleRailRender() {
    if (railPending) return;
    railPending = true;
    const raf = typeof requestAnimationFrame === "function" ? requestAnimationFrame : (cb) => setTimeout(cb, 16);
    raf(() => {
      railPending = false;
      paintVisibleThumbs();
    });
  }

  // Thumbnails are drawn only for the slides near the rail viewport, so a deck
  // with hundreds of slides does not pay for all of them at once.
  function paintVisibleThumbs() {
    if (!railEl) return;
    const railRect = measureBox(railEl) || { top: 0, bottom: 600 };
    for (const button of Array.from(railEl.children)) {
      const rect = measureBox(button);
      const visible = !rect || (rect.bottom > railRect.top - 400 && rect.top < railRect.bottom + 400);
      if (!visible) continue;
      const index = Number(button.dataset.index);
      if (button.dataset.painted === "1") continue;
      button.dataset.painted = "1";
      const frame = button.querySelector(".ov-pptx-thumbbox");
      if (!frame) continue;
      // The thumbnail is the whole slide, scaled to the width of its frame.
      const mini = drawSlide(model.slides[index], { thumbnail: true, width: slideWidth });
      const frameWidth = frame.clientWidth || 96;
      mini.style.transform = "scale(" + (frameWidth / slideWidth) + ")";
      mini.style.transformOrigin = "top left";
      frame.appendChild(mini);
    }
  }

  function showSlide(index) {
    current = clampSlide(index);
    if (!stageEl) return;
    stageEl.empty();
    const slide = model.slides[current];
    if (!slide) {
      stageEl.createDiv("ov-pptx-empty", (el) => el.setText("This presentation has no slides."));
      return;
    }
    slideEl = drawSlide(slide, { width: slideWidth, height: slideHeight });
    slideEl.addClass("ov-pptx-slide");
    const frame = stageEl.createDiv("ov-pptx-frame");
    frame.appendChild(slideEl);
    applyScale();
    updateRailSelection();
    renderNotes(slide);
    warmNeighbourMedia(current, slide);
    callbacks.onSlideChange({ index: current, count: model.slides.length, title: slide.title });
  }

  // Decode the next slide's pictures while the reader is still on this one, so
  // flipping through a large deck does not wait for a decompression each time.
  // The media cache owns the result and releases it with the model.
  function warmNeighbourMedia(index, slide) {
    if (!model.mediaUrl || slide.__ovWarm) return;
    slide.__ovWarm = true;
    const next = model.slides[index + 1];
    if (!next) return;
    const paths = [];
    const collect = (shapes) => {
      for (const shape of shapes || []) {
        if (!shape || paths.length >= 24) continue;
        if (shape.type === "group") collect(shape.shapes);
        else if (shape.path) paths.push(shape.path);
        else if (shape.fill && shape.fill.type === "image" && shape.fill.path) paths.push(shape.fill.path);
      }
    };
    collect(next.shapes);
    for (const path of paths) {
      try {
        model.mediaUrl(path);
      } catch (err) {
        // A part that cannot be decoded stays broken; the next slide shows its
        // own placeholder rather than failing to render.
      }
    }
  }

  function updateRailSelection() {
    for (const button of Array.from(railEl.children)) {
      const isCurrent = Number(button.dataset.index) === current;
      button.toggleClass("is-current", isCurrent);
    }
    const active = railEl.children[current];
    if (active && active.scrollIntoView) active.scrollIntoView({ block: "nearest" });
  }

  function applyScale() {
    if (!slideEl) return;
    const base = computeBaseScale();
    const s = base * scale;
    slideEl.style.transform = "scale(" + s + ")";
    slideEl.style.width = slideWidth + "px";
    slideEl.style.height = slideHeight + "px";
    const frame = slideEl.parentElement;
    if (frame) {
      frame.style.width = slideWidth * s + "px";
      frame.style.height = slideHeight * s + "px";
    }
  }

  // "contain" fits the whole slide, "width" fills the width; zoom multiplies
  // whichever fit is active.
  function computeBaseScale() {
    if (!stageEl) return 1;
    const available = measureBox(stageEl);
    const width = available && available.width ? available.width : slideWidth;
    const height = available && available.height ? available.height : slideHeight;
    if (settings.fit === "width") return Math.max(0.05, (width - 24) / slideWidth);
    if (settings.fit === "none") return 1;
    return Math.max(0.05, Math.min((width - 24) / slideWidth, (height - 24) / slideHeight));
  }

  // Fit follows the pane: entering or leaving fullscreen, opening a sidebar or
  // resizing the window all change how much room the slide has. Without this
  // the slide keeps the scale it was first drawn at, which left a deck zoomed
  // after a slideshow ended.
  function watchStageSize() {
    if (typeof ResizeObserver === "function") {
      resizeObserver = new ResizeObserver(scheduleResize);
      resizeObserver.observe(stageEl);
    } else if (typeof window !== "undefined" && window.addEventListener) {
      window.addEventListener("resize", scheduleResize);
    }
  }

  function scheduleResize() {
    if (resizePending || destroyed) return;
    resizePending = true;
    const raf = typeof requestAnimationFrame === "function" ? requestAnimationFrame : (cb) => setTimeout(cb, 16);
    raf(() => {
      resizePending = false;
      if (!destroyed) applyScale();
    });
  }

  function renderNotes(slide) {
    if (!notesEl) return;
    notesEl.empty();
    const show = settings.showNotes && slide.notes;
    notesEl.toggleClass("is-hidden", !show);
    if (!show) return;
    const head = notesEl.createDiv("ov-pptx-noteshead");
    head.setText("Speaker notes");
    notesEl.createDiv("ov-pptx-notesbody").setText(slide.notes);
  }

  // ---------- one slide ----------

  function drawSlide(slide, o) {
    const width = o.width || slideWidth;
    const height = o.height || slideHeight;
    const page = doc.createElement("div");
    page.className = "ov-pptx-page";
    page.style.width = width + "px";
    page.style.height = height + "px";
    page.style.position = "relative";
    page.style.overflow = "hidden";
    applyBackground(page, slide.background, slide);
    // Layout and master furniture sits under the slide's own content.
    for (const shape of slide.decorations || []) {
      const el = drawDecoration(shape, slide, o);
      if (el) page.appendChild(el);
    }
    for (const shape of slide.shapes) {
      const el = drawShape(shape, slide, { thumbnail: o.thumbnail });
      if (el) page.appendChild(el);
    }
    return page;
  }

  function applyBackground(page, background, slide) {
    if (!background) {
      page.style.background = "#ffffff";
      return;
    }
    if (background.type === "image") {
      page.style.background = "#ffffff";
      drawBackgroundImage(page, background, slide);
      return;
    }
    const css = backgroundCss(background, slide);
    page.style.background = css || "#ffffff";
  }

  // A slide background image covers the whole slide. PowerPoint stretches it
  // to the slide box unless the file asks for a tile.
  function drawBackgroundImage(page, background, slide) {
    const url = background.path && model.mediaUrl ? model.mediaUrl(background.path) : null;
    if (!url) return;
    const img = doc.createElement("img");
    img.className = "ov-pptx-bgimage";
    img.src = url;
    img.alt = "";
    img.setAttribute("draggable", "false");
    img.style.position = "absolute";
    img.style.inset = "0";
    img.style.width = "100%";
    img.style.height = "100%";
    img.style.objectFit = background.tile ? "none" : "fill";
    if (background.crop) {
      const crop = background.crop;
      const scaleX = 1 / Math.max(0.01, 1 - crop.l - crop.r);
      const scaleY = 1 / Math.max(0.01, 1 - crop.t - crop.b);
      if (scaleX !== 1 || scaleY !== 1) {
        img.style.transform = "scale(" + scaleX + ", " + scaleY + ")";
        img.style.transformOrigin = (crop.l > 0 ? "left" : crop.r > 0 ? "right" : "center") + " " +
          (crop.t > 0 ? "top" : crop.b > 0 ? "bottom" : "center");
      }
    }
    page.appendChild(img);
    void slide;
  }

  function backgroundCss(background, slide) {
    if (background.type === "solid") return background.color || "#ffffff";
    if (background.type === "gradient") return gradientCss(background, slide);
    if (background.type === "none") return "transparent";
    return "#ffffff";
  }

  // ---------- shapes ----------

  // Decorations come from the layout and the master, so they are drawn with
  // their own inherited box and their fields resolved for this slide.
  function drawDecoration(shape, slide, o) {
    if (shape.type === "group") {
      const el = doc.createElement("div");
      el.className = "ov-pptx-group";
      positionElement(el, shape, o);
      for (const child of shape.shapes) {
        const childEl = drawDecoration(child, slide, o);
        if (childEl) el.appendChild(childEl);
      }
      return el;
    }
    const hasText = shape.textInfo && shape.textInfo.text && shape.textInfo.text.trim();
    if (!hasText && !shape.fill && !shape.line) return null;
    return drawShape(shape, slide, Object.assign({}, o, { decoration: true }));
  }

  function drawShape(shape, slide, o) {
    if (shape.type === "group") return drawGroup(shape, slide, o);
    const el = doc.createElement("div");
    el.className = "ov-pptx-shape";
    positionElement(el, shape, o);
    if (shape.rotation) el.style.transform = "rotate(" + shape.rotation + "deg)";
    if (shape.flipH || shape.flipV) {
      el.style.scale = (shape.flipH ? "-1" : "1") + " " + (shape.flipV ? "-1" : "1");
    }
    applyShapeFill(el, shape, slide, o);
    applyShapeLine(el, shape, slide, o);
    applyShapeEffects(el, shape);

    if (shape.type === "picture") {
      drawPicture(el, shape, slide, o);
    } else if (shape.type === "table") {
      if (o.thumbnail) el.addClass("ov-pptx-tableblock");
      else drawTable(el, shape, slide, o);
    }
    else if (shape.type === "chart") drawChart(el, shape);
    else if (shape.type === "diagram") drawDiagram(el, shape, slide, o);
    else if (shape.type === "connector") drawConnector(el, shape, slide, o);
    // A line can also be a plain shape with the line preset; without this the
    // element only carried a border on a zero width box and nothing showed.
    else if (shape.geometry && shape.geometry.preset === "line") drawConnector(el, shape, slide, o);
    else if (shape.textInfo && shape.textInfo.text) drawTextBody(el, shape, slide, o);

    if (shape.geometry && shape.geometry.preset && shape.geometry.preset !== "rect" && shape.geometry.preset !== "line") {
      applyPresetGeometry(el, shape, slide, o);
    }
    return el;
  }

  function positionElement(el, shape, o) {
    const box = mappedBox(shape, o);
    el.style.left = box.x + "px";
    el.style.top = box.y + "px";
    el.style.width = box.cx + "px";
    el.style.height = box.cy + "px";
  }

  function drawGroup(group, slide, o) {
    const el = doc.createElement("div");
    el.className = "ov-pptx-group";
    // The wrapper holds nothing; every descendant is drawn in slide
    // coordinates through the group map, so a nested transform composes in
    // plain arithmetic. A CSS scale on the wrapper turned a 0.0006 px child
    // space into a 348000 px element and Chrome stopped rasterising its SVG.
    el.style.left = "0px";
    el.style.top = "0px";
    const t = group.groupTransform;
    let childOptions = o;
    if (t) {
      const parent = o.map || { ax: 1, ay: 1, bx: 0, by: 0 };
      const sx = t.chExtX ? t.extX / t.chExtX : 1;
      const sy = t.chExtY ? t.extY / t.chExtY : 1;
      // p -> off + (p - chOff) * S, composed with the parent's map.
      childOptions = Object.assign({}, o, {
        map: {
          ax: parent.ax * sx,
          ay: parent.ay * sy,
          bx: parent.bx + parent.ax * (t.offX - t.chOffX * sx),
          by: parent.by + parent.ay * (t.offY - t.chOffY * sy),
        },
      });
    }
    for (const child of group.shapes) {
      if (!child) continue;
      const childEl = drawShape(child, slide, childOptions);
      if (childEl) el.appendChild(childEl);
    }
    return el;
  }

  // A shape's box after every enclosing group transform, in slide pixels.
  // Group scaling multiplies geometry; it never multiplies a font size or a
  // stroke width, which stay absolute the way PowerPoint keeps them.
  function mappedBox(shape, o) {
    const map = o && o.map;
    const xEmu = map ? shape.xEmu * map.ax + map.bx : shape.xEmu;
    const yEmu = map ? shape.yEmu * map.ay + map.by : shape.yEmu;
    const cxEmu = map ? shape.cxEmu * map.ax : shape.cxEmu;
    const cyEmu = map ? shape.cyEmu * map.ay : shape.cyEmu;
    return {
      xEmu,
      yEmu,
      cxEmu,
      cyEmu,
      x: emuToPx(xEmu),
      y: emuToPx(yEmu),
      cx: emuToPx(cxEmu),
      cy: emuToPx(cyEmu),
    };
  }

  function applyShapeFill(el, shape, slide, o) {
    const fill = shape.fill || (shape.placeholderBox && shape.placeholderBox.fill) || null;
    if (!fill || fill.type === "none") return;
    if (fill.type === "solid") {
      el.style.background = fill.color || "#ffffff";
      return;
    }
    if (fill.type === "gradient") {
      const css = gradientCss(fill, slide);
      if (css) el.style.background = css;
      return;
    }
    if (fill.type === "pattern") {
      el.style.backgroundColor = fill.bg || "#ffffff";
      el.style.backgroundImage = patternImage(fill);
      if (fill.pattern === "pct5" || fill.pattern === "pct10" || fill.pattern === "pct20" || fill.pattern === "pct25" || fill.pattern === "dotGrid") {
        el.style.backgroundSize = "6px 6px";
      }
      return;
    }
    if (fill.type === "image") {
      // A shape or background filled with a picture, which is how pasted
      // screenshots and some SmartArt pieces arrive.
      const url = fill.path && model.mediaUrl ? model.mediaUrl(fill.path) : null;
      if (!url) return;
      el.style.backgroundImage = 'url("' + url + '")';
      if (fill.tile) {
        el.style.backgroundSize = "auto";
        el.style.backgroundRepeat = "repeat";
        return;
      }
      const crop = fill.crop || {};
      const left = crop.l || 0;
      const right = crop.r || 0;
      const top = crop.t || 0;
      const bottom = crop.b || 0;
      const scaleX = 1 / Math.max(0.01, 1 - left - right);
      const scaleY = 1 / Math.max(0.01, 1 - top - bottom);
      el.style.backgroundSize = scaleX * 100 + "% " + scaleY * 100 + "%";
      el.style.backgroundRepeat = "no-repeat";
      const posX = left + right > 0 ? (-left / (left + right)) * 100 : 50;
      const posY = top + bottom > 0 ? (-top / (top + bottom)) * 100 : 50;
      el.style.backgroundPosition = posX + "% " + posY + "%";
      return;
    }
  }

  function gradientCss(fill, slide) {
    if (!fill.stops || !fill.stops.length) return null;
    // CSS clamps a stop that sits before its predecessor to that position,
    // which collapses the list; a file can list the stops unordered and the
    // cone wash is what that looks like.
    const stops = fill.stops.slice().sort((a, b) => (a.pos || 0) - (b.pos || 0)).map((stop) => {
      const color = resolveCss(stop.color, slide) || "#ffffff";
      return color + " " + Math.round(stop.pos * 100) + "%";
    });
    if (fill.path) {
      return "radial-gradient(circle at center, " + stops.join(", ") + ")";
    }
    const angle = 90 - (fill.angleDeg || 0);
    return "linear-gradient(" + angle + "deg, " + stops.join(", ") + ")";
  }

  // A small subset of the DrawingML preset patterns, drawn with CSS gradients.
  // Each entry returns the background-image value for one foreground colour.
  const PATTERN_BUILDERS = {
    pct5: (c) => "radial-gradient(" + c + " 1px, transparent 1px)",
    pct10: (c) => "radial-gradient(" + c + " 1px, transparent 1px)",
    pct20: (c) => "radial-gradient(" + c + " 1.6px, transparent 1.6px)",
    pct25: (c) => "radial-gradient(" + c + " 1.8px, transparent 1.8px)",
    dotGrid: (c) => "radial-gradient(" + c + " 1px, transparent 1px)",
    smGrid: (c) => "linear-gradient(" + c + " 1px, transparent 1px), linear-gradient(90deg, " + c + " 1px, transparent 1px)",
    lgGrid: (c) => "linear-gradient(" + c + " 1.6px, transparent 1.6px), linear-gradient(90deg, " + c + " 1.6px, transparent 1.6px)",
    horz: (c) => "repeating-linear-gradient(0deg, " + c + " 0 1.6px, transparent 1.6px 8px)",
    vert: (c) => "repeating-linear-gradient(90deg, " + c + " 0 1.6px, transparent 1.6px 8px)",
    ltHorz: (c) => "repeating-linear-gradient(0deg, " + c + " 0 1px, transparent 1px 9px)",
    ltVert: (c) => "repeating-linear-gradient(90deg, " + c + " 0 1px, transparent 1px 9px)",
    diag: (c) => "repeating-linear-gradient(45deg, " + c + " 0 1.6px, transparent 1.6px 8px)",
    cross: (c) => "repeating-linear-gradient(45deg, " + c + " 0 1px, transparent 1px 8px), repeating-linear-gradient(-45deg, " + c + " 0 1px, transparent 1px 8px)",
    diagCross: (c) => "repeating-linear-gradient(45deg, " + c + " 0 2px, transparent 2px 8px), repeating-linear-gradient(-45deg, " + c + " 0 2px, transparent 2px 8px)",
  };

  function patternImage(fill) {
    const builder = PATTERN_BUILDERS[fill.pattern] || PATTERN_BUILDERS.pct5;
    return builder(fill.fg || "#000000");
  }

  function applyShapeLine(el, shape, slide, o) {
    const line = shape.line;
    if (!line || line.color === "none") return;
    const width = Math.max(0.5, emuToPx(line.widthEmu));
    const color = resolveCss(line.color, slide) || "#000000";
    const dash = DASH_PATTERNS[line.dash] || "solid";
    el.style.outline = width + "px " + dash + " " + color;
    el.style.outlineOffset = "-" + width / 2 + "px";
    void o;
  }

  const DASH_PATTERNS = {
    dash: "dashed",
    dashDot: "dashed",
    dot: "dotted",
    lgDash: "dashed",
    lgDashDot: "dashed",
    lgDashDotDot: "dashed",
    sysDash: "dashed",
    sysDashDot: "dashed",
    sysDot: "dotted",
  };

  function applyShapeEffects(el, shape) {
    const effects = shape.effects;
    if (!effects || !effects.shadow) return;
    const shadow = effects.shadow;
    const distance = emuToPx(shadow.distEmu);
    const angle = ((shadow.dir || 45) * Math.PI) / 180;
    const x = Math.cos(angle) * distance;
    const y = Math.sin(angle) * distance;
    el.style.boxShadow = x.toFixed(1) + "px " + y.toFixed(1) + "px " + emuToPx(shadow.blurEmu).toFixed(1) + "px " + (shadow.color || "rgba(0,0,0,0.35)");
  }

  const PRESET_RADIUS = {
    roundRect: "12%",
    round1Rect: "8%",
    round2SameRect: "10%",
    round2DiagRect: "10%",
    snipRoundRect: "8%",
    plaque: "12%",
    ellipse: "50%",
    circle: "50%",
    flowChartConnector: "50%",
  };

  function applyPresetGeometry(el, shape, slide, o) {
    const geometry = shape.geometry || {};
    const preset = geometry.preset;
    if (!preset || preset === "rect") return;
    const box = mappedBox(shape, o);
    if (preset === "custom") {
      const width = box.cx;
      const height = box.cy;
      const data = customPathData(geometry.custom, width, height);
      if (!data || width <= 0 || height <= 0) return;
      const fill = shape.fill;
      if (!fill || fill.type === "solid") {
        // Draw the silhouette itself. A clipped box can only show its border
        // where the outline happens to touch the rectangle, so a concave
        // freeform would lose its line entirely.
        el.style.background = "none";
        el.style.outline = "none";
        el.style.boxShadow = "none";
        const svg = createSvg(doc, width, height);
        const path = svgEl(doc, "path");
        path.setAttribute("d", data);
        path.setAttribute("fill", fill ? resolveCss(fill.color, slide) || "none" : "none");
        if (shape.line && shape.line.color !== "none") {
          const lineWidth = Math.max(0.5, emuToPx(shape.line.widthEmu || 12700));
          path.setAttribute("stroke", resolveCss(shape.line.color, slide) || "#000000");
          path.setAttribute("stroke-width", lineWidth);
          if (shape.line.dash) path.setAttribute("stroke-dasharray", dashArray(shape.line.dash, lineWidth));
        }
        svg.appendChild(path);
        el.appendChild(svg);
        return;
      }
      // A picture, gradient or pattern fill keeps the element's own background
      // and only borrows the silhouette as a clip.
      el.style.clipPath = 'path("' + data + '")';
      return;
    }
    if (preset === "ellipse" || preset === "circle") {
      el.style.borderRadius = "50%";
      return;
    }
    if (PRESET_RADIUS[preset]) {
      el.style.borderRadius = PRESET_RADIUS[preset];
      return;
    }
    // The preset table covers the silhouettes proper: straight edged shapes
    // come back as points, curved ones as an SVG path, both in local pixels, so
    // one clip path reproduces the shape without an extra element.
    const w = box.cx;
    const h = box.cy;
    if (w > 0 && h > 0) {
      const points = presetPolygon(preset, w, h, geometry.adjustments);
      if (points && points.length >= 6) {
        const parts = [];
        for (let i = 0; i + 1 < points.length; i += 2) {
          parts.push(round1(points[i]) + "px " + round1(points[i + 1]) + "px");
        }
        el.style.clipPath = "polygon(" + parts.join(", ") + ")";
        return;
      }
      const path = presetPath(preset, w, h, geometry.adjustments);
      const data = path ? [path.path, path.fillPath, path.strokePath].filter(Boolean).join(" ") : "";
      if (data) {
        el.style.clipPath = 'path("' + data + '")';
        return;
      }
    }
    if (preset === "triangle" || preset === "rtTriangle") {
      el.style.clipPath = "polygon(50% 0%, 100% 100%, 0% 100%)";
    } else if (preset === "diamond") {
      el.style.clipPath = "polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)";
    } else if (preset === "star5") {
      el.style.clipPath = "polygon(50% 0%, 61% 35%, 98% 35%, 68% 57%, 79% 91%, 50% 70%, 21% 91%, 32% 57%, 2% 35%, 39% 35%)";
    }
  }

  function round1(value) {
    return Math.round(value * 100) / 100;
  }

  // a:custGeom paths carry their own coordinate space; scale it to the shape
  // and turn the DrawingML segments into SVG path data. An arcTo names an angle
  // on an ellipse whose centre follows from the current point, so the endpoint
  // is computed the way the file means it rather than guessed.
  function customPathData(paths, width, height) {
    if (!paths || !paths.length) return "";
    const parts = [];
    for (const path of paths) {
      if (path.fill === "none") continue;
      const sx = path.w > 0 ? width / path.w : 1;
      const sy = path.h > 0 ? height / path.h : 1;
      let cx = 0;
      let cy = 0;
      let startX = 0;
      let startY = 0;
      for (const segment of path.segments || []) {
        if (segment.type === "M") {
          cx = startX = segment.x;
          cy = startY = segment.y;
          parts.push("M" + round1(segment.x * sx) + " " + round1(segment.y * sy));
        } else if (segment.type === "L") {
          cx = segment.x;
          cy = segment.y;
          parts.push("L" + round1(segment.x * sx) + " " + round1(segment.y * sy));
        } else if (segment.type === "C") {
          const control1 = segment.pts[0];
          const control2 = segment.pts[1];
          const end = segment.pts[2];
          parts.push("C" + round1(control1.x * sx) + " " + round1(control1.y * sy) + " "
            + round1(control2.x * sx) + " " + round1(control2.y * sy) + " "
            + round1(end.x * sx) + " " + round1(end.y * sy));
          cx = end.x;
          cy = end.y;
        } else if (segment.type === "Q") {
          const control = segment.pts[0];
          const end = segment.pts[1];
          parts.push("Q" + round1(control.x * sx) + " " + round1(control.y * sy) + " "
            + round1(end.x * sx) + " " + round1(end.y * sy));
          cx = end.x;
          cy = end.y;
        } else if (segment.type === "A") {
          const startAngle = (segment.stAng / 60000) * Math.PI / 180;
          const sweepAngle = (segment.swAng / 60000) * Math.PI / 180;
          const centreX = cx - segment.wR * Math.cos(startAngle);
          const centreY = cy - segment.hR * Math.sin(startAngle);
          const endX = centreX + segment.wR * Math.cos(startAngle + sweepAngle);
          const endY = centreY + segment.hR * Math.sin(startAngle + sweepAngle);
          const large = Math.abs(segment.swAng) > 10800000 ? 1 : 0;
          const sweep = segment.swAng >= 0 ? 1 : 0;
          parts.push("A" + round1(segment.wR * sx) + " " + round1(segment.hR * sy) + " 0 " + large + " " + sweep + " "
            + round1(endX * sx) + " " + round1(endY * sy));
          cx = endX;
          cy = endY;
        } else if (segment.type === "Z") {
          parts.push("Z");
          cx = startX;
          cy = startY;
        }
      }
    }
    return parts.join(" ");
  }

  // ---------- pictures ----------

  // Chromium has no decoder for these, so naming the format is more useful
  // than "Image": the reader can then decide to open the file elsewhere.
  const UNDECODABLE_IMAGE = {
    wdp: "JPEG XR (HD Photo)",
    hdp: "JPEG XR (HD Photo)",
    jxr: "JPEG XR (HD Photo)",
    eps: "EPS",
    ai: "Illustrator",
    psd: "Photoshop",
  };

  function drawPicture(el, shape, slide, o) {
    const url = model.mediaUrl ? model.mediaUrl(shape.path) : null;
    if (!url) {
      el.addClass("ov-pptx-missing");
      if (!o.thumbnail) {
        const note = el.createDiv("ov-pptx-missingtext");
        const ext = String(shape.path || "").split(".").pop().toLowerCase();
        const format = UNDECODABLE_IMAGE[ext];
        note.setText(shape.alt || (format ? format + " image, not decodable here" : "Image"));
      }
      return;
    }
    const img = doc.createElement("img");
    img.src = url;
    img.alt = shape.alt || "";
    img.style.width = "100%";
    img.style.height = "100%";
    img.style.objectFit = "cover";
    if (shape.crop) {
      const crop = shape.crop;
      const scaleX = 1 / Math.max(0.01, 1 - crop.l - crop.r);
      const scaleY = 1 / Math.max(0.01, 1 - crop.t - crop.b);
      img.style.transform = "scale(" + scaleX + ", " + scaleY + ")";
      img.style.transformOrigin = (crop.l > 0 ? "left" : crop.r > 0 ? "right" : "center") + " " + (crop.t > 0 ? "top" : crop.b > 0 ? "bottom" : "center");
    }
    img.setAttribute("draggable", "false");
    el.appendChild(img);
  }

  // ---------- connectors and lines ----------

  function drawConnector(el, shape, slide, o) {
    // The flip is baked into the SVG coordinates below, so the generic flip on
    // the element would apply it a second time and mirror every arrow.
    el.style.scale = "";
    const box = mappedBox(shape, o);
    const width = box.cx;
    const height = box.cy;
    const flipH = shape.flipH;
    const flipV = shape.flipV;
    const preset = shape.geometry ? shape.geometry.preset : "line";
    const lineColor = resolveCss(shape.line && shape.line.color, slide) || "#000000";
    const lineWidth = Math.max(0.5, emuToPx(shape.line ? shape.line.widthEmu : 12700));

    if (preset === "line" || preset === "straightConnector1") {
      // A horizontal or vertical line has a zero extent on one axis. An SVG
      // with a zero width or height collapses and draws nothing; give the
      // viewport one line width on that axis and centre the line in it, in
      // local pixels so a scaled group still scales the stroke with the line.
      const svgWidth = width > 0 ? width : lineWidth;
      const svgHeight = height > 0 ? height : lineWidth;
      const x1 = width > 0 ? (flipH ? width : 0) : lineWidth / 2;
      const x2 = width > 0 ? (flipH ? 0 : width) : lineWidth / 2;
      const y1 = height > 0 ? (flipV ? height : 0) : lineWidth / 2;
      const y2 = height > 0 ? (flipV ? 0 : height) : lineWidth / 2;
      const svg = createSvg(doc, svgWidth, svgHeight);
      if (width <= 0) svg.style.left = (-lineWidth / 2) + "px";
      if (height <= 0) svg.style.top = (-lineWidth / 2) + "px";
      const line = svgEl(doc, "line");
      line.setAttribute("x1", x1);
      line.setAttribute("y1", y1);
      line.setAttribute("x2", x2);
      line.setAttribute("y2", y2);
      line.setAttribute("stroke", lineColor);
      line.setAttribute("stroke-width", lineWidth);
      if (shape.line && shape.line.dash) line.setAttribute("stroke-dasharray", dashArray(shape.line.dash, lineWidth));
      if (shape.line && shape.line.arrowTail) addArrow(doc, svg, line, shape.line.arrowTail, { x: x2, y: y2 }, { x: x1, y: y1 }, lineColor, lineWidth);
      if (shape.line && shape.line.arrowHead) addArrow(doc, svg, line, shape.line.arrowHead, { x: x1, y: y1 }, { x: x2, y: y2 }, lineColor, lineWidth);
      svg.appendChild(line);
      el.appendChild(svg);
      el.style.background = "none";
      el.style.outline = "none";
      return;
    }
    // Bent and curved connectors: a simple elbow is close enough at slide size.
    const svg = createSvg(doc, width, height);
    const path = svgEl(doc, "path");
    const d = preset === "curvedConnector3"
      ? "M " + (flipH ? width : 0) + " " + (flipV ? height : 0) + " Q " + width / 2 + " " + (flipV ? height : 0) + " " + width / 2 + " " + height / 2 + " T " + (flipH ? 0 : width) + " " + (flipV ? 0 : height)
      : "M " + (flipH ? width : 0) + " " + (flipV ? height : 0) + " L " + (flipH ? 0 : width) + " " + (flipV ? height : 0) + " L " + (flipH ? 0 : width) + " " + (flipV ? 0 : height);
    path.setAttribute("d", d);
    path.setAttribute("fill", "none");
    path.setAttribute("stroke", lineColor);
    path.setAttribute("stroke-width", lineWidth);
    if (shape.line && shape.line.dash) path.setAttribute("stroke-dasharray", dashArray(shape.line.dash, lineWidth));
    svg.appendChild(path);
    el.appendChild(svg);
    el.style.background = "none";
    el.style.outline = "none";
  }

  function dashArray(dash, width) {
    switch (dash) {
      case "dash": return String(width * 4) + " " + String(width * 3);
      case "dot": return String(width) + " " + String(width * 2);
      case "dashDot": return String(width * 4) + " " + String(width * 2) + " " + String(width) + " " + String(width * 2);
      case "lgDash": return String(width * 8) + " " + String(width * 3);
      case "sysDash": return String(width * 3) + " " + String(width);
      default: return null;
    }
  }

  function addArrow(docRef, svg, line, spec, tip, from, color, width) {
    const size = width * (spec.w === "lg" ? 5 : spec.w === "sm" ? 2.5 : 3.5);
    const angle = Math.atan2(tip.y - from.y, tip.x - from.x);
    const spread = 0.45;
    const p1 = { x: tip.x - Math.cos(angle - spread) * size, y: tip.y - Math.sin(angle - spread) * size };
    const p2 = { x: tip.x - Math.cos(angle + spread) * size, y: tip.y - Math.sin(angle + spread) * size };
    if (spec.type === "none") return;
    if (spec.type === "oval") {
      const circle = svgEl(docRef, "circle");
      circle.setAttribute("cx", tip.x);
      circle.setAttribute("cy", tip.y);
      circle.setAttribute("r", size / 2);
      circle.setAttribute("fill", color);
      svg.appendChild(circle);
      return;
    }
    const polygon = svgEl(docRef, "polygon");
    polygon.setAttribute("points", tip.x + "," + tip.y + " " + p1.x + "," + p1.y + " " + p2.x + "," + p2.y);
    polygon.setAttribute("fill", color);
    svg.appendChild(polygon);
    void line;
  }

  function createSvg(docRef, width, height) {
    const svg = docRef.createElementNS ? docRef.createElementNS("http://www.w3.org/2000/svg", "svg") : docRef.createElement("svg");
    svg.setAttribute("width", String(width));
    svg.setAttribute("height", String(height));
    svg.setAttribute("viewBox", "0 0 " + width + " " + height);
    svg.setAttribute("preserveAspectRatio", "none");
    svg.style.position = "absolute";
    svg.style.left = "0";
    svg.style.top = "0";
    svg.style.overflow = "visible";
    return svg;
  }

  function svgEl(docRef, name) {
    return docRef.createElementNS ? docRef.createElementNS("http://www.w3.org/2000/svg", name) : docRef.createElement(name);
  }

  // ---------- tables ----------

  function drawTable(frame, shape, slide, o) {
    const table = shape.table;
    if (!table) return;
    const el = doc.createElement("table");
    el.className = "ov-pptx-table";
    el.style.width = "100%";
    el.style.height = "100%";
    el.style.borderCollapse = "collapse";
    el.style.tableLayout = "fixed";
    const totalGrid = table.grid.reduce((a, b) => a + b, 0) || 1;
    const colgroup = doc.createElement("colgroup");
    for (const w of table.grid) {
      const col = doc.createElement("col");
      col.style.width = (w / totalGrid) * 100 + "%";
      colgroup.appendChild(col);
    }
    el.appendChild(colgroup);
    const tbody = doc.createElement("tbody");
    el.appendChild(tbody);
    for (const row of table.rows) {
      const tr = doc.createElement("tr");
      if (row.height) tr.style.height = emuToPx(row.height) + "px";
      for (const cell of row.cells) {
        const td = doc.createElement("td");
        if (cell.gridSpan > 1) td.colSpan = cell.gridSpan;
        if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
        if (cell.fill && cell.fill.type === "solid") td.style.background = cell.fill.color;
        const margins = cell.margins || {};
        td.style.padding = emuToPx(margins.marT != null ? margins.marT : 45720) + "px " +
          emuToPx(margins.marR != null ? margins.marR : 91440) + "px " +
          emuToPx(margins.marB != null ? margins.marB : 45720) + "px " +
          emuToPx(margins.marL != null ? margins.marL : 91440) + "px";
        const borders = cell.borders;
        if (borders) {
          for (const [key, side] of [["lnT", "Top"], ["lnB", "Bottom"], ["lnL", "Left"], ["lnR", "Right"]]) {
            const border = borders[key];
            if (!border || border.color === "none") continue;
            td.style["border" + side] = Math.max(1, emuToPx(border.widthEmu)) + "px solid " + (resolveCss(border.color, slide) || "#000000");
          }
        }
        if (cell.anchor === 1) td.style.verticalAlign = "middle";
        else if (cell.anchor === 2) td.style.verticalAlign = "bottom";
        if (cell.text) drawTextBody(td, cell, slide, Object.assign({}, o, { tableCell: true }));
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    frame.appendChild(el);
  }

  function drawChart(frame, shape) {
    const placeholder = frame.createDiv("ov-pptx-placeholder");
    placeholder.setText(shape.chart && shape.chart.label ? shape.chart.label : "Chart");
    const title = frame.createDiv("ov-pptx-placeholder-sub");
    title.setText("Chart data is kept in the workbook part and is not plotted here.");
  }

  function drawDiagram(frame, shape, slide, o) {
    // SmartArt is stored as shapes in a sibling part; when that part is absent
    // the drawing rectangle still shows the fallback text.
    if (shape.textInfo && shape.textInfo.text) {
      drawTextBody(frame, shape, slide, o);
      return;
    }
    const placeholder = frame.createDiv("ov-pptx-placeholder");
    placeholder.setText(shape.diagram && shape.diagram.label ? shape.diagram.label : "Diagram");
  }

  // ---------- text ----------

  function drawTextBody(frame, shape, slide, o) {
    const body = shape.textInfo || shape.text;
    if (!body || !body.paragraphs || !body.paragraphs.length) return;
    const box = frame.createDiv("ov-pptx-text");
    // Text that does not fit keeps its shape's box but is not hidden: a
    // clipped word is worse than a word that spills, and Office itself spills
    // unless the shape asks to shrink.
    box.style.overflow = "visible";
    const bodyProps = body.bodyProps || {};
    // normAutofit can declare a line spacing reduction; Office applies it to
    // every line, including paragraphs with no spacing of their own.
    const lineScale = bodyProps.autofit && bodyProps.autofit.lnSpcReduction ? 1 - bodyProps.autofit.lnSpcReduction : 1;
    const insets = bodyProps.insets || DEFAULT_INSETS;
    box.style.padding = emuToPx(insets.t) + "px " + emuToPx(insets.r) + "px " + emuToPx(insets.b) + "px " + emuToPx(insets.l) + "px";
    const anchor = bodyProps.anchor || "t";
    box.style.display = "flex";
    box.style.flexDirection = "column";
    box.style.justifyContent = anchor === "ctr" ? "center" : anchor === "b" ? "flex-end" : "flex-start";
    if (bodyProps.vert && bodyProps.vert !== "horz") box.addClass("ov-pptx-vertical");
    if (bodyProps.wrap === "none") box.addClass("ov-pptx-nowrap");
    if (bodyProps.autofit && bodyProps.autofit.type === "shape") box.addClass("ov-pptx-autosize");

    const shapeBox = mappedBox(shape, o);
    const shapeWidth = shapeBox.cx;
    const innerWidth = Math.max(10, shapeWidth - emuToPx(insets.l) - emuToPx(insets.r));
    const defaultSizePt = shapeDefaultFontSize(shape, slide);
    const scale = fitTextToBox(body, innerWidth, shapeBox.cy, defaultSizePt, o, slide);

    const paragraphs = body.paragraphs;
    for (let index = 0; index < paragraphs.length; index++) {
      const paragraph = paragraphs[index];
      const p = box.createDiv("ov-pptx-p");
      const props = paragraph.props || {};
      // marL and indent arrive resolved from the shape's list style and the
      // master's text styles. A hanging bullet sits at marL + indent and the
      // text starts at marL, which is exactly text-indent's model.
      const marginLeft = paragraph.marLEmu != null ? emuToPx(paragraph.marLEmu) : (props.marginLeftEmu != null ? emuToPx(props.marginLeftEmu) : 0);
      const indent = paragraph.indentEmu != null ? emuToPx(paragraph.indentEmu) : (props.indentEmu != null ? emuToPx(props.indentEmu) : 0);
      p.style.paddingLeft = Math.max(0, marginLeft) + "px";
      p.style.textIndent = indent + "px";
      if (props.align) p.style.textAlign = alignCss(props.align);
      if (props.spaceBeforePt && typeof props.spaceBeforePt === "number") p.style.marginTop = ptToPx(props.spaceBeforePt) + "px";
      if (props.spaceAfterPt && typeof props.spaceAfterPt === "number") p.style.marginBottom = ptToPx(props.spaceAfterPt) + "px";
      const lineSpacingPct = props.lineSpacingPct ? props.lineSpacingPct * lineScale : null;
      const lineSpacingPt = props.lineSpacingPt ? props.lineSpacingPt * lineScale : null;
      if (lineSpacingPct) p.style.lineHeight = String(lineSpacingPct);
      else if (lineSpacingPt) p.style.lineHeight = ptToPx(lineSpacingPt) + "px";
      else if (lineScale !== 1) p.style.lineHeight = String(1.18 * lineScale);
      // A list that ends with an empty paragraph shows no marker for it, which
      // is what PowerPoint and the reference renderers all print.
      const trailingEmpty = !paragraph.text && index === paragraphs.length - 1;
      const bullet = trailingEmpty ? null : paragraph.bullet;
      const firstRun = paragraph.runs.find((run) => run.type === "run" && run.props) || null;
      if (bullet) {
        const marker = p.createSpan("ov-pptx-bullet");
        marker.setText(bullet.kind === "auto" ? autoNumber(bullet, paragraph) : bullet.char || "\u2022");
        // The marker has no run of its own, so it takes the size and colour of
        // the paragraph's own level: a second level bullet is smaller because
        // its text is. The shape default gives every bullet one size, which
        // leaves a first level dot too small for a 28pt list.
        const levelPt = firstRun && firstRun.props && firstRun.props.sizePt ? firstRun.props.sizePt : defaultSizePt;
        const bulletPt = levelPt * (props.bulletSize || 1) * (scale || 1);
        marker.style.fontSize = ptToPx(bulletPt) + "px";
        const bulletColour = props.bulletColor || (firstRun && firstRun.props.color) || null;
        marker.style.color = resolveCss(bulletColour, slide) || undefined;
        const hanging = Math.max(0, -indent);
        if (hanging >= 1) {
          // Give the marker the width the negative indent opened up, so the
          // text after it starts at marL as it does in PowerPoint.
          marker.style.display = "inline-block";
          marker.style.width = hanging + "px";
          marker.style.marginRight = "0";
        }
      }
      for (const run of paragraph.runs) {
        if (run.type === "br") {
          p.createEl("br");
          continue;
        }
        if (run.type === "math") {
          // The equation tree and its DOM come from the shared renderer; it
          // takes the paragraph level's size so it sits at the text scale.
          const mathSpan = p.createSpan("ov-pptx-math");
          const mathPt = (firstRun && firstRun.props && firstRun.props.sizePt ? firstRun.props.sizePt : defaultSizePt) * (scale || 1);
          mathSpan.style.fontSize = ptToPx(mathPt) + "px";
          if (firstRun && firstRun.props && firstRun.props.color) mathSpan.style.color = resolveCss(firstRun.props.color, slide) || undefined;
          renderMathDom(run.node, mathSpan, null);
          continue;
        }
        const span = p.createSpan("ov-pptx-r");
        span.setText(fieldText(run, slide));
        applyRunStyle(span, run.props, slide, defaultSizePt, scale);
      }
      if (!paragraph.runs.length) p.createEl("br");
    }
    frame.appendChild(box);
    return scale;
  }

  // Slide number fields carry a placeholder glyph in the file, so the number is
  // substituted at draw time. Other fields keep the text they were given.
  function fieldText(run, slide) {
    if (run.type !== "field") return run.text || "";
    const kind = String(run.kind || "").toLowerCase();
    if (kind.indexOf("slidenum") !== -1) {
      const index = slide && typeof slide.index === "number" ? slide.index + 1 : 1;
      return String(index);
    }
    return run.text || "";
  }

  function shapeDefaultFontSize(shape, slide) {
    const info = shape.textInfo;
    if (info && info.listStyle) {
      const lvl = info.listStyle.levelAt(1);
      if (lvl && lvl.runProps && lvl.runProps.sizePt) return lvl.runProps.sizePt;
    }
    const placeholder = shape.placeholder;
    if (placeholder) {
      if (placeholder.type === "title" || placeholder.type === "ctrTitle") return 28;
      if (placeholder.type === "subTitle") return 20;
      if (placeholder.type === "body") return 18;
    }
    void slide;
    return SHAPE_DEFAULT_FONT_PT;
  }

  function applyRunStyle(span, props, slide, defaultSizePt, scale) {
    // DrawingML baseline is 1000ths of a percent of the font size. The glyphs
    // are also drawn at 58% of the run's size, the proportional height
    // The reference renderers both use; the shift divides by that so it
    // stays the declared fraction of the original em (the usual -25000
    // subscript moves down a quarter of the full size).
    const baseline = props && props.baseline ? props.baseline : 0;
    const size = (props && props.sizePt ? props.sizePt : defaultSizePt) * (scale || 1) * (baseline ? 0.58 : 1);
    span.style.fontSize = ptToPx(size) + "px";
    if (props.bold) span.style.fontWeight = "700";
    if (props.italic) span.style.fontStyle = "italic";
    if (props.underline && props.underline !== "none") span.style.textDecoration = "underline";
    if (props.strike) span.style.textDecoration = (span.style.textDecoration ? span.style.textDecoration + " " : "") + "line-through";
    if (props.color) span.style.color = resolveCss(props.color, slide) || undefined;
    if (props.highlight) span.style.background = resolveCss(props.highlight, slide) || undefined;
    if (props.spacingPt) span.style.letterSpacing = ptToPx(props.spacingPt) + "px";
    if (baseline) span.style.verticalAlign = baseline / 58 + "em";
    if (props.caps === "all") span.style.textTransform = "uppercase";
    if (props.caps === "small") span.style.fontVariant = "small-caps";
    span.style.fontFamily = fontFamilyCss(props.fontFamily, slide, props.fontFamilyTheme);
  }

  // Autofit shrinks text until the laid-out height fits the shape. The height
  // is estimated from measured line counts, which is how PowerPoint decides
  // when to shrink; the declared fontScale is used as the starting point when
  // the file carries one. The search is a bisection, so it lands within a
  // percent of the fit: a five percent step leaves a title one word too wide.
  function fitTextToBox(body, widthPx, heightPx, defaultSizePt, o, slide) {
    if (o.thumbnail) return 0.06;
    if (!body.bodyProps || !body.bodyProps.autofit || body.bodyProps.autofit.type === "none") return 1;
    if (!heightPx || heightPx <= 0 || widthPx <= 0) return 1;
    const declared = body.bodyProps.autofit.fontScale;
    const start = declared && declared < 1 ? Math.max(MIN_AUTOFIT_SCALE, declared) : 1;
    const lineScale = body.bodyProps.autofit.lnSpcReduction ? 1 - body.bodyProps.autofit.lnSpcReduction : 1;
    const estimate = (scale) => {
      let total = 0;
      for (const paragraph of body.paragraphs) {
        const props = paragraph.props || {};
        const runProps = props.runProps || null;
        const sizePt = (runProps && runProps.sizePt ? runProps.sizePt : defaultSizePt) * scale * (runProps && runProps.baseline ? 0.58 : 1);
        const linePx = ptToPx(sizePt) * (props.lineSpacingPct || 1.18) * lineScale;
        // marL is resolved during parsing; the raw paragraph attribute is a
        // fallback only for older shapes.
        const indent = paragraph.marLEmu != null ? emuToPx(paragraph.marLEmu) : (props.marginLeftEmu != null ? emuToPx(props.marginLeftEmu) : 0);
        const bulletRoom = paragraph.bullet ? ptToPx(sizePt) * 1.6 : 0;
        const available = Math.max(20, widthPx - indent - bulletRoom);
        total += Math.max(1, wrappedLineCount(paragraph, available, defaultSizePt, scale, slide)) * linePx;
        total += (props.spaceBeforePt || 0) * scale + (props.spaceAfterPt || 0) * scale;
      }
      return total;
    };
    if (estimate(start) <= heightPx) return start;
    let lo = MIN_AUTOFIT_SCALE;
    let hi = start;
    for (let i = 0; i < 8 && hi - lo > 0.01; i++) {
      const mid = (lo + hi) / 2;
      if (estimate(mid) > heightPx) hi = mid;
      else lo = mid;
    }
    return Math.max(MIN_AUTOFIT_SCALE, lo);
  }

  // Wrapped line count for one paragraph: runs are joined and measured in
  // order, because summing a per-run ceiling counts a line for every run and
  // makes a title shrink when it fits.
  function wrappedLineCount(paragraph, available, basePt, scale, slide) {
    let lines = 1;
    let lineWidth = 0;
    for (const run of paragraph.runs) {
      if (run.type === "br") {
        lines += 1;
        lineWidth = 0;
        continue;
      }
      const runPt = (run.props && run.props.sizePt ? run.props.sizePt : basePt) * (scale || 1) * (run.props && run.props.baseline ? 0.58 : 1);
      // Measure with the family the span is drawn with, or the estimate uses a
      // different face than the layout and the two disagree about wrapping.
      const runFamily = fontFamilyCss(run.props && run.props.fontFamily, slide, run.props && run.props.fontFamilyTheme);
      const font = ptToPx(runPt) + "px " + runFamily;
      const words = String(run.text || "").split(/(\s+)/);
      for (const word of words) {
        if (!word) continue;
        const width = fontMetrics.measure(word, font);
        if (lineWidth > 0 && lineWidth + width > available) {
          lines += 1;
          lineWidth = width;
        } else {
          lineWidth += width;
        }
      }
    }
    return lines;
  }

  function autoNumber(bullet, paragraph) {
    const n = (bullet.startAt || 1) + (paragraph.bulletIndex || 0);
    const type = bullet.type || "arabicPeriod";
    const suffix = type.indexOf("Paren") !== -1 ? ")" : ".";
    const base = type.replace(/(Period|Paren|Plain)$/, "");
    switch (base) {
      case "romanUc": return toRoman(n).toUpperCase() + suffix;
      case "romanLc": return toRoman(n).toLowerCase() + suffix;
      case "alphaUc": return toLetters(n).toUpperCase() + suffix;
      case "alphaLc": return toLetters(n).toLowerCase() + suffix;
      default: return String(n) + suffix;
    }
  }

  function toRoman(n) {
    const table = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
    let out = "";
    let v = Math.max(1, Math.min(3999, n));
    for (const [value, symbol] of table) {
      while (v >= value) {
        out += symbol;
        v -= value;
      }
    }
    return out;
  }

  function toLetters(n) {
    let out = "";
    let v = Math.max(1, n);
    while (v > 0) {
      const rem = (v - 1) % 26;
      out = String.fromCharCode(97 + rem) + out;
      v = Math.floor((v - 1) / 26);
    }
    return out;
  }

  function alignCss(align) {
    switch (align) {
      case "ctr": return "center";
      case "r": return "right";
      case "just": return "justify";
      default: return "left";
    }
  }

  // ---------- colour resolution ----------

  function resolveCss(color, slide) {
    if (!color) return null;
    if (typeof color !== "string") return null;
    if (color.indexOf("scheme:") === 0) {
      const slot = color.slice(7);
      return schemeColor(slide, slot);
    }
    return color;
  }

  function schemeColor(slide, slot) {
    const map = slide.clrMap || (model.masters.size ? null : null);
    const mapped = map && map[slot] ? map[slot] : slot;
    const theme = slide.theme;
    if (!theme) return null;
    if (theme.colorMap && theme.colorMap[mapped]) return theme.colorMap[mapped];
    if (theme.byName && theme.byName[mapped]) return theme.byName[mapped];
    return null;
  }

  // ---------- small helpers ----------

  // Element measurement with the guards needed by the test environment, where
  // getBoundingClientRect exists but reports zeroes outside a layout engine.
  function measureBox(el) {
    if (!el || typeof el.getBoundingClientRect !== "function") return null;
    try {
      const rect = el.getBoundingClientRect();
      if (!rect) return null;
      if (!rect.width && !rect.height && !rect.top && !rect.bottom) return null;
      return rect;
    } catch (err) {
      return null;
    }
  }

  function createFontMetrics(docRef) {
    let measureCtx = null;
    const cache = new Map();
    return {
      measure(text, font) {
        if (!measureCtx && docRef.createElement) {
          const canvas = docRef.createElement("canvas");
          measureCtx = canvas.getContext ? canvas.getContext("2d") : null;
        }
        const key = font + "\u0000" + text;
        if (cache.has(key)) return cache.get(key);
        let width;
        if (measureCtx) {
          measureCtx.font = font;
          width = measureCtx.measureText(text).width;
        } else {
          width = text.length * parseFloat(font) * 0.5;
        }
        cache.set(key, width);
        return width;
      },
    };
  }

  // ---------- public API ----------

  function nextSlide() {
    if (current < model.slides.length - 1) showSlide(current + 1);
  }

  function previousSlide() {
    if (current > 0) showSlide(current - 1);
  }

  function goToSlide(index) {
    showSlide(index);
  }

  function setSettings(patch) {
    Object.assign(settings, patch);
    if (rootEl) rootEl.style.setProperty("--ov-pptx-scale", String(scale));
    applyScale();
    renderNotes(model.slides[current] || {});
  }

  function setZoom(z) {
    scale = Math.max(0.1, Math.min(6, z));
    applyScale();
  }

  function getState() {
    return {
      index: current,
      count: model.slides.length,
      title: model.slides[current] ? model.slides[current].title : "",
    };
  }

  function destroy() {
    destroyed = true;
    thumbnailObserver.disconnect();
    if (resizeObserver) {
      resizeObserver.disconnect();
      resizeObserver = null;
    } else if (typeof window !== "undefined" && window.removeEventListener) {
      window.removeEventListener("resize", scheduleResize);
    }
    if (container) container.textContent = "";
    rootEl = railEl = stageEl = slideEl = notesEl = null;
    thumbnailCache.clear();
  }

  // Lazy thumbnail painting when the platform has IntersectionObserver; the
  // scroll listener above covers the rest.
  const thumbnailObserver = typeof IntersectionObserver === "function"
    ? new IntersectionObserver((entries) => {
      let any = false;
      for (const entry of entries) {
        if (entry.isIntersecting) {
          entry.target.dataset.visible = "1";
          any = true;
        }
      }
      if (any) scheduleRailRender();
    }, { rootMargin: "400px" })
    : { observe() {}, disconnect() {} };

  render();

  return {
    destroy,
    setSettings,
    setZoom,
    nextSlide,
    previousSlide,
    goToSlide,
    getState,
    getSlideCount: () => model.slides.length,
    resolveCss,
    fontMetrics,
  };
}

module.exports = { createPptxRenderer };
