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

// Presentation controller: slides in a rail, one on the stage, notes under it.
// Keyboard navigation is handled by the view so it works with the slide list.

const { parsePptx } = require("../pptx/parse");
const { createPptxRenderer } = require("../pptx/render");
const { readOdp } = require("../odf/parse");

class PresentationController {
  constructor(opts) {
    this.host = opts.host;
    this.plugin = opts.plugin;
    this.format = opts.format;
    this.model = null;
    this.renderer = null;
    this.callbacks = opts.callbacks || {};
    this.slideIndex = 0;
  }

  load(bytes) {
    this.model = this.format === "odp" ? readOdp(bytes) : parsePptx(bytes);
    return this.model;
  }

  mount() {
    if (!this.model) return;
    this.renderer = createPptxRenderer({
      container: this.host,
      model: this.model,
      settings: this.settings(),
      onSlideChange: (info) => {
        this.slideIndex = info.index;
        this.callbacks.onSlideChange ? this.callbacks.onSlideChange(info) : null;
        this.emitStatus();
        this.saveState();
      },
      onExternalLink: (href) => this.openExternal(href),
      onReady: (info) => {
        this.callbacks.onReady ? this.callbacks.onReady(info) : null;
        this.emitStatus();
      },
    });
  }

  settings() {
    return {
      zoom: this.currentZoom() / 100,
      fit: this.plugin.settings.slideFit || "contain",
      showNotes: this.plugin.settings.showSpeakerNotes !== false,
      // Reopen on the slide the reader left off on.
      slideIndex: this.slideIndex,
    };
  }

  currentZoom() {
    const z = this.plugin.settings.slideZoom;
    return typeof z === "number" && z >= 10 && z <= 400 ? z : 100;
  }

  applySettings() {
    if (!this.renderer) return;
    this.renderer.setSettings(this.settings());
  }

  setZoom(value) {
    const zoom = Math.max(10, Math.min(400, Math.round(value / 10) * 10));
    this.plugin.updateSetting("slideZoom", zoom);
    if (this.renderer) this.renderer.setZoom(zoom / 100);
    this.emitStatus();
  }

  setFit(mode) {
    this.plugin.updateSetting("slideFit", mode);
    this.applySettings();
  }

  search(query) {
    if (!this.renderer || !this.model) return 0;
    const needle = String(query || "").toLowerCase();
    if (!needle) return 0;
    let count = 0;
    this.model.slides.forEach((slide, index) => {
      if (slide.text && slide.text.toLowerCase().indexOf(needle) !== -1) count++;
    });
    this.searchHits = [];
    this.model.slides.forEach((slide, index) => {
      if (slide.text && slide.text.toLowerCase().indexOf(needle) !== -1) this.searchHits.push(index);
    });
    this.searchIndex = -1;
    return count;
  }

  // Search walks slides rather than text fragments: a slide is the unit a
  // lecture is reviewed in.
  stepSearch(dir) {
    const hits = this.searchHits || [];
    if (!hits.length) return null;
    this.searchIndex = (this.searchIndex + dir + hits.length * 2) % hits.length;
    const slide = hits[this.searchIndex];
    if (this.renderer) this.renderer.goToSlide(slide);
    return { index: this.searchIndex + 1, count: hits.length };
  }

  nextSlide() {
    if (this.renderer) this.renderer.nextSlide();
  }

  previousSlide() {
    if (this.renderer) this.renderer.previousSlide();
  }

  goToSlide(index) {
    if (this.renderer) this.renderer.goToSlide(index);
  }

  emitStatus() {
    this.callbacks.onStatus ? this.callbacks.onStatus(this.statusText()) : null;
  }

  describe() {
    const rows = [];
    const model = this.model;
    if (!model) return rows;
    const properties = model.properties || {};
    if (properties.title) rows.push({ label: "Title", value: properties.title });
    if (properties.author) rows.push({ label: "Author", value: properties.author });
    if (properties.application) rows.push({ label: "Application", value: properties.application });
    if (properties.created) rows.push({ label: "Created", value: formatDate(properties.created) });
    rows.push({ label: "Slides", value: String(model.slides.length) });
    rows.push({ label: "Slide size", value: Math.round(model.widthPx) + " x " + Math.round(model.heightPx) + " px" });
    if (properties.slides) rows.push({ label: "Slides (as saved)", value: properties.slides });
    let notes = 0, pictures = 0, tables = 0, charts = 0, shapes = 0;
    const walk = (list) => {
      for (const shape of list) {
        if (shape.type === "group") { walk(shape.shapes); continue; }
        shapes++;
        if (shape.type === "picture") pictures++;
        else if (shape.type === "table") tables++;
        else if (shape.type === "chart") charts++;
      }
    };
    for (const slide of model.slides) {
      walk(slide.shapes);
      if (slide.notes) notes++;
    }
    rows.push({ label: "Shapes", value: String(shapes) });
    if (pictures) rows.push({ label: "Pictures", value: String(pictures) });
    if (tables) rows.push({ label: "Tables", value: String(tables) });
    if (charts) rows.push({ label: "Charts", value: String(charts) });
    if (notes) rows.push({ label: "Slides with notes", value: notes + " of " + model.slides.length });
    const media = model.pkg ? model.pkg.list().filter((name) => /^ppt\/media\//.test(name)).length : 0;
    if (media) rows.push({ label: "Media parts", value: String(media) });
    const theme = model.slides[0] && model.slides[0].theme;
    if (theme && theme.fonts && (theme.fonts.major || theme.fonts.minor)) {
      rows.push({ label: "Theme fonts", value: [theme.fonts.major, theme.fonts.minor].filter(Boolean).join(" / ") });
    }
    return rows;
  }

  statusText() {
    if (!this.model || !this.model.slides.length) return "";
    const total = this.model.slides.length;
    const index = this.slideIndex + 1;
    const parts = ["Slide " + index + " of " + total];
    const title = this.model.slides[this.slideIndex] ? this.model.slides[this.slideIndex].title : "";
    if (title) parts.push(title.replace(/\s+/g, " ").slice(0, 70));
    if (this.model.properties && this.model.properties.slides) {
      parts.push("deck reports " + this.model.properties.slides + " slides");
    }
    return parts.join("  |  ");
  }

  saveState() {
    if (!this.filePath) return;
    this.plugin.saveFileState(this.filePath, {
      kind: "presentation",
      slideIndex: this.slideIndex,
    });
  }

  restoreState(saved) {
    if (saved && Number.isInteger(saved.slideIndex)) this.slideIndex = saved.slideIndex;
  }

  openExternal(url) {
    try {
      const electron = require("electron");
      if (electron && electron.shell && electron.shell.openExternal) {
        electron.shell.openExternal(url);
        return;
      }
    } catch (err) {
      // Not on desktop; fall back to window.open below.
    }
    window.open(url, "_blank");
  }

  // Mounts a model that came out of the document cache instead of parsing the
  // bytes again. The model owns its media URLs; the cache keeps them alive.
  adoptModel(model) {
    this.model = model;
    return model;
  }

  destroy() {
    if (this.renderer) {
      this.renderer.destroy();
      this.renderer = null;
    }
    this.model = null;
    this.searchHits = null;
  }
}

function formatDate(value) {
  const text = String(value || "");
  const match = /^(\d{4})-(\d{2})-(\d{2})T?(\d{2}:\d{2})?/.exec(text);
  if (!match) return text;
  return match[1] + "-" + match[2] + "-" + match[3] + (match[4] ? " " + match[4] : "");
}

module.exports = { PresentationController };
