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

// The view. One FileView serves every supported extension: it detects the
// format from the bytes, builds the shared toolbar, hands the file to the
// right controller, and keeps the loading, empty and error states honest.
//
// Viewing only: nothing in this file, or in anything it calls, writes to the
// opened document. The only writes are the plugin's own settings.

const { FileView, Menu, setIcon } = require("obsidian");
const { SpreadsheetController } = require("./spreadsheet-controller");
const { DocumentController } = require("./document-controller");
const { PresentationController } = require("./presentation-controller");
const { detectFormat } = require("./format-router");
const { DocumentCache, releaseModel } = require("./document-cache");
const { ensureJxrDecoder } = require("../media/jxr");
const { ensureEmbeddedFonts } = require("../pptx/fonts");
const { describePackage, sectionsToText } = require("../shared/metadata");

const VIEW_TYPE = "office-viewer";
const DEFAULT_CHIP_LIMIT = 10;

// True when the package lists a JPEG XR part by name; the bytes are only
// inflated if one is actually drawn.
function modelHasJxr(model) {
  const pkg = model && model.pkg;
  if (!pkg || typeof pkg.list !== "function") return false;
  return pkg.list().some((name) => /\.(wdp|hdp|jxr)$/i.test(name));
}

// Resolves after the next paint, so a loading line added to the DOM is on
// screen before a synchronous parse takes the thread.
function nextPaint() {
  return new Promise((resolve) => {
    const raf = typeof requestAnimationFrame === "function" ? requestAnimationFrame : (cb) => setTimeout(cb, 16);
    raf(() => raf(resolve));
  });
}

class OfficeView extends FileView {
  constructor(leaf, plugin) {
    super(leaf);
    this.plugin = plugin;
    this.file = null;
    this.format = null;
    this.controller = null;
    this.savedFileState = null;
    this.built = false;
    this.error = null;
    this.outlinePanelOpen = false;
    this.metaOpen = false;
    // Reopening a document should be instant, so a few parsed models are kept.
    // Every open also takes a token: a slow read that finishes after the reader
    // moved to another file is dropped instead of mounting the wrong document.
    this.documentCache = new DocumentCache({ limit: 3, maxBytes: 32 * 1024 * 1024 });
    this.loadToken = 0;
    // Slideshow mode fills the pane with the slide and takes the arrow keys.
    // It never changes what is drawn: every shape is on the slide already,
    // because the renderer ignores the animation tree.
    this.slideshow = false;
    // Escape is the browser's own key while an element is fullscreen: it leaves
    // fullscreen and the key never reaches the view. Listening for the change is
    // the only reliable way to leave the mode and refit the slide to the pane.
    this.onSlideshowFullscreenChange = () => this.handleFullscreenChange();
  }

  getViewType() {
    return VIEW_TYPE;
  }

  getDisplayText() {
    return this.file ? this.file.name : "Office document";
  }

  getIcon() {
    if (!this.format) return "file-text";
    if (this.format.kind === "spreadsheet") return "table-2";
    if (this.format.kind === "presentation") return "presentation";
    return "file-text";
  }

  async onOpen() {
    this.buildChrome();
    this.renderEmpty();
    if (this.registerDomEvent) {
      this.registerDomEvent(document, "keydown", (ev) => this.onSlideshowKey(ev));
      // Documents and decks get a menu of their own; the grid has had one for a
      // while. Capture phase, so the app's plain text menu does not win.
      this.registerDomEvent(this.hostEl, "contextmenu", (ev) => this.onDocumentContextMenu(ev), true);
      // A touchpad pinch arrives as a wheel with ctrlKey set; a plain wheel
      // stays a scroll. Not passive, because the pinch consumes the event.
      this.registerDomEvent(this.contentEl, "wheel", (ev) => this.onWheelZoom(ev), { passive: false });
    }
    // A file that changed on disk must not be served from the cache. Not every
    // host exposes vault events, so the subscription is optional.
    const vault = this.app && this.app.vault;
    if (vault && typeof vault.on === "function" && this.registerEvent) {
      const drop = (changed) => {
        if (changed && changed.path) this.documentCache.dropPath(changed.path);
      };
      this.registerEvent(vault.on("modify", drop));
      this.registerEvent(vault.on("delete", drop));
      this.registerEvent(vault.on("rename", (changed, oldPath) => {
        drop(changed);
        if (oldPath) this.documentCache.dropPath(oldPath);
      }));
    }
    if (this.register) this.register(() => this.documentCache.clear());
  }

  async onLoadFile(file) {
    this.buildChrome();
    // Name the tab before anything is read, so it never opens blank while the
    // bytes are fetched and the package is parsed.
    this.file = file;
    this.titleEl.setText(file.name);
    this.updateHeader();
    await this.renderFile(file);
  }

  // The tab title comes from getDisplayText; the header has to be told when it
  // becomes known. Not every host exposes the call.
  updateHeader() {
    if (this.leaf && typeof this.leaf.updateHeader === "function") this.leaf.updateHeader();
  }

  async onUnloadFile() {
    this.teardownController();
    this.file = null;
    this.format = null;
    this.renderEmpty();
  }

  async onClose() {
    this.teardownController();
  }

  // ---------- chrome ----------

  buildChrome() {
    if (this.built) return;
    this.built = true;
    const content = this.contentEl;
    content.empty();
    content.addClass("ov-view");

    this.toolbarEl = content.createDiv("ov-toolbar");
    this.titleEl = this.toolbarEl.createDiv("ov-title");
    this.toolsEl = this.toolbarEl.createDiv("ov-tools");

    // The finder chevrons sit left of the box and appear only for a search that
    // has matches; the count sits right of the box.
    this.searchNavEl = this.toolsEl.createSpan("ov-search-nav is-hidden");
    this.searchInput = this.toolsEl.createEl("input", {
      cls: "ov-search-input",
      attr: { type: "text", placeholder: "Search", spellcheck: "false" },
    });
    this.searchCountEl = this.toolsEl.createSpan("ov-search-count");
    this.searchInput.addEventListener("input", () => this.runSearch(this.searchInput.value));
    this.searchInput.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") {
        ev.preventDefault();
        this.stepSearch(ev.shiftKey ? -1 : 1);
      } else if (ev.key === "Escape") {
        this.searchInput.value = "";
        this.runSearch("");
      }
    });

    this.zoomLabel = this.toolsEl.createSpan("ov-zoom-label");
    this.pageLabelEl = this.toolsEl.createSpan("ov-page-label is-hidden");

    this.tabsEl = content.createDiv("ov-tabs");
    this.outlineEl = content.createDiv("ov-outline is-hidden");
    this.bodyEl = content.createDiv("ov-body");
    this.hostEl = this.bodyEl.createDiv("ov-host");
    this.metaEl = this.bodyEl.createDiv("ov-meta is-hidden");
    this.filterBarEl = content.createDiv("ov-filterbar");
    this.statusEl = content.createDiv("ov-status");
    this.applyStatusMode();
  }

  // Rebuilds the toolbar buttons for the format that was just detected.
  buildToolbar() {
    if (this.toolsEl) {
      for (const el of Array.from(this.toolsEl.querySelectorAll(".ov-tool-btn, .ov-toggle"))) el.remove();
    }
    this.searchCountEl.setText("");
    this.searchInput.value = "";
    const kind = this.format ? this.format.kind : "empty";

    this.setSearchNav(0);
    this.addButton("chevron-up", "Previous match", () => this.stepSearch(-1), this.searchNavEl);
    this.addButton("chevron-down", "Next match", () => this.stepSearch(1), this.searchNavEl);
    this.addButton("zoom-out", "Zoom out", () => this.setZoom(this.currentZoom() - 10));
    const zoomIn = this.addButton("zoom-in", "Zoom in", () => this.setZoom(this.currentZoom() + 10));
    // The percentage sits with the magnifying glasses, not among the search
    // controls; the label is created once and moved after the buttons.
    if (zoomIn && zoomIn.parentNode) zoomIn.parentNode.insertBefore(this.zoomLabel, zoomIn.nextSibling);
    this.zoomLabel.setText(this.currentZoom() + "%");

    if (kind === "spreadsheet") {
      this.backgroundBtn = this.addButton("contrast", "Toggle white or theme background", () => {
        if (this.controller) this.controller.toggleBackground();
        this.refreshToolbarState();
      });
      this.filterBtn = this.addButton("filter", "Show or hide the filter chips", () => {
        if (this.controller) this.controller.toggleFilterBar();
        this.refreshFilterBar();
        this.refreshToolbarState();
      });
      this.addButton("list-x", "Clear all filters", () => {
        if (this.controller) this.controller.clearFilters();
        this.refreshFilterBar();
      });
      this.addButton("stretch-horizontal", "Reset the column widths and row heights", () => {
        if (this.controller) this.controller.resetSizes();
      });
    } else if (kind === "presentation") {
      this.addButton("chevron-left", "Previous slide", () => this.controller && this.controller.previousSlide());
      this.addButton("chevron-right", "Next slide", () => this.controller && this.controller.nextSlide());
      this.slideshowBtn = this.addButton("play", "Start or stop the slideshow", () => this.toggleSlideshow());
      this.fitBtn = this.addButton("move-vertical", "Fit the slide to the window or its own size", () => {
        const next = this.plugin.settings.slideFit === "contain" ? "none" : "contain";
        if (this.controller) this.controller.setFit(next);
        this.refreshToolbarState();
      });
    } else if (kind === "document") {
      this.outlineBtn = this.addButton("list", "Show or hide the heading outline", () => this.toggleOutline());
      if (this.pageLabelEl) this.pageLabelEl.removeClass("is-hidden");
    }
    if (this.pageLabelEl && kind !== "document") this.pageLabelEl.addClass("is-hidden");
    this.metaBtn = this.addButton("info", "Show or hide the file details", () => this.toggleMeta());
    this.addButton("refresh-cw", "Reload the file", () => this.reload());
    this.refreshToolbarState();
  }

  addButton(icon, title, onClick, parent) {
    const btn = (parent || this.toolsEl).createEl("button", {
      cls: "ov-tool-btn",
      attr: { "aria-label": title, title },
    });
    setIcon(btn, icon);
    btn.addEventListener("click", onClick);
    return btn;
  }

  refreshToolbarState() {
    if (this.backgroundBtn) {
      const white = this.plugin.settings.sheetBackground !== "theme";
      this.backgroundBtn.toggleClass("is-active", white);
    }
    if (this.filterBtn) {
      this.filterBtn.toggleClass("is-active", this.controller ? this.controller.showFilterBar !== false : true);
    }
    if (this.fitBtn) {
      this.fitBtn.toggleClass("is-active", this.plugin.settings.slideFit !== "none");
    }
    if (this.outlineBtn) {
      this.outlineBtn.toggleClass("is-active", this.outlinePanelOpen);
    }
    if (this.zoomLabel) this.zoomLabel.setText(this.currentZoom() + "%");
  }

  // ---------- file lifecycle ----------

  // Called from onLoadFile and the reload button. Never throws: a file problem
  // is reported inside the view so the leaf stays open.
  async renderFile(file) {
    const token = ++this.loadToken;
    this.exitSlideshow();
    this.buildChrome();
    this.error = null;
    this.file = file;
    this.controller && this.teardownController();
    this.hostEl.empty();
    this.tabsEl.empty();
    this.outlineEl.empty();
    this.outlineEl.addClass("is-hidden");
    this.filterBarEl.empty();
    this.statusEl.setText("");
    this.titleEl.setText(file.name);
    this.updateHeader();
    // Up before the first byte is read, so the pane shows the file is opening.
    const loading = this.hostEl.createDiv("ov-loading");
    loading.setText("Reading " + file.name + "\u2026");

    try {
      const cacheKey = DocumentCache.keyFor(file.path, file.stat || null);
      const cached = this.documentCache.lookup(cacheKey);
      let bytes = null;
      if (cached) {
        this.format = cached.format;
      } else {
        bytes = new Uint8Array(await this.app.vault.readBinary(file));
        // Another file was opened while this one was being read.
        if (token !== this.loadToken) return;
        this.format = detectFormat(file.name, bytes);
      }
      this.buildToolbar();
      if (this.format.kind === "unsupported") {
        this.showUnsupported(this.format.reason);
        return;
      }
      // Let the tab and the loading line paint before the parse takes the
      // thread. The parse is synchronous, and a large deck would otherwise hold
      // the first frame, which is what makes an open feel like a stutter.
      if (!cached) await nextPaint();
      // The JPEG XR codec is WebAssembly, and most packages carry no JPEG XR
      // picture at all. A document resolves its media during the parse, so the
      // module has to be ready before it runs; a deck or a sheet draws later,
      // so it asks for the codec only when the package lists a part. That
      // keeps the wasm off the common open path.
      const jxrBeforeParse = this.format.kind === "document";
      if (jxrBeforeParse) await ensureJxrDecoder();

      const saved = (this.plugin.settings.fileState || {})[file.path] || null;
      this.savedFileState = saved;
      const callbacks = this.controllerCallbacks();
      if (this.format.kind === "spreadsheet") {
        this.controller = new SpreadsheetController({
          host: this.hostEl,
          tabsEl: this.tabsEl,
          filterBarEl: this.filterBarEl,
          plugin: this.plugin,
          format: this.format.ext,
          callbacks,
        });
      } else if (this.format.kind === "document") {
        this.controller = new DocumentController({
          host: this.hostEl,
          plugin: this.plugin,
          format: this.format.ext,
          callbacks,
        });
      } else {
        this.controller = new PresentationController({
          host: this.hostEl,
          plugin: this.plugin,
          format: this.format.ext,
          callbacks,
        });
      }
      this.controller.filePath = file.path;
      this.controller.savedState = saved;
      if (cached && typeof this.controller.adoptModel === "function") {
        this.controller.adoptModel(cached.model);
      } else {
        // Parse first. Restoring the remembered sheet or slide needs the parsed
        // model, so the order here matters.
        this.controller.load(bytes, file.name);
        if (bytes && (this.format.kind === "document" || this.format.kind === "presentation") && typeof this.controller.adoptModel === "function") {
          // The cache owns the model's media from here until it is evicted.
          const stored = this.documentCache.store(cacheKey, { model: this.controller.model, format: this.format }, bytes.byteLength);
          if (stored && this.controller.model) this.controller.model.__ovCached = true;
        }
      }
      // A presentation can carry its own fonts. They register before the mount
      // so autofit measures the text with the face it will be drawn in.
      await ensureEmbeddedFonts(this.controller.model);
      if (!jxrBeforeParse && modelHasJxr(this.controller.model)) await ensureJxrDecoder();
      if (typeof this.controller.restoreState === "function") this.controller.restoreState(saved);
      if (loading.parentNode) loading.parentNode.removeChild(loading);
      this.controller.mount();
      this.refreshFilterBar();
      this.refreshToolbarState();
      this.metaEl.empty();
      if (this.metaOpen) this.renderMeta();
    } catch (err) {
      this.controller = null;
      this.showError(err);
    }
  }

  controllerCallbacks() {
    return {
      onStatus: (text) => {
        this.statusEl.setText(text);
        this.setAppStatus(text);
        // The panel shows counts that change as the reader moves around.
        this.refreshMeta();
      },
      onFiltersChanged: () => this.refreshFilterBar(),
      onOutline: (outline) => this.renderOutline(outline),
      onSlideChange: (info) => this.updateTitleWithSlide(info),
      onPage: (info) => this.updatePageIndicator(info),
      onReady: () => this.refreshToolbarState(),
      showContextMenu: (ev, items) => this.showContextMenu(ev, items),
    };
  }

  updateTitleWithSlide(info) {
    if (!this.file) return;
    const suffix = info && info.title ? " \u2013 " + info.title.replace(/\s+/g, " ").slice(0, 60) : "";
    this.titleEl.setText(this.file.name + (info ? "  (" + (info.index + 1) + "/" + info.count + ")" : "") + suffix);
  }

  // The page the reader is on, in the toolbar beside the zoom, so a long
  // document never leaves them guessing which sheet is on screen.
  updatePageIndicator(info) {
    if (!this.pageLabelEl) return;
    const show = Boolean(this.format && this.format.kind === "document" && info && info.count);
    this.pageLabelEl.toggleClass("is-hidden", !show);
    if (show) this.pageLabelEl.setText("Page " + info.page + " / " + info.count);
  }

  refreshFilterBar() {
    if (!this.filterBarEl || !this.controller || this.format.kind !== "spreadsheet") return;
    const limit = Number(this.plugin.settings.chipLimit) > 0
      ? Number(this.plugin.settings.chipLimit)
      : DEFAULT_CHIP_LIMIT;
    this.controller.refreshFilterChips(this.filterBarEl, limit);
  }

  // ---------- outline ----------

  renderOutline(outline) {
    this.outlineEl.empty();
    if (!outline || !outline.length) {
      this.outlineEl.addClass("is-hidden");
      this.outlinePanelOpen = false;
      return;
    }
    for (const item of outline) {
      const row = this.outlineEl.createDiv("ov-outline-item");
      row.style.paddingLeft = (item.level - 1) * 14 + "px";
      row.setText(item.text);
      row.addEventListener("click", () => {
        if (this.controller && typeof this.controller.scrollToHeading === "function") {
          this.controller.scrollToHeading(item);
        }
      });
    }
    this.outlineEl.toggleClass("is-hidden", !this.outlinePanelOpen);
  }

  // ---------- file details ----------

  // The app's status bar is where a reader looks for counts, and it costs the
  // document no height. The in view line stays for hosts without one, hidden by
  // the stylesheet. The setting moves the line back into the view, bottom left.
  setAppStatus(text) {
    if (this.statusInView()) text = "";
    const el = this.plugin && this.plugin.statusBarEl;
    if (!el || !el.setText) return;
    el.setText(text || "");
    if (el.toggleClass) el.toggleClass("is-hidden", !text);
  }

  statusInView() {
    return Boolean(this.plugin && this.plugin.settings && this.plugin.settings.statusInView);
  }

  applyStatusMode() {
    const inView = this.statusInView();
    this.contentEl.toggleClass("ov-status-inline", inView);
    this.setAppStatus(this.statusEl ? this.statusEl.textContent || "" : "");
  }

  // ---------- slideshow ----------

  toggleSlideshow() {
    if (this.slideshow) this.exitSlideshow();
    else this.startSlideshow();
  }

  startSlideshow() {
    if (!this.format || this.format.kind !== "presentation" || !this.controller) return;
    this.slideshow = true;
    this.contentEl.addClass("ov-slideshow");
    if (this.slideshowBtn) this.slideshowBtn.addClass("is-active");
    const hint = "Slideshow: arrows or space to move, Escape to leave";
    if (this.statusEl) this.statusEl.setText(hint);
    this.setAppStatus(hint);
    if (this.controller.applySettings) this.controller.applySettings();
    if (typeof document !== "undefined" && document.addEventListener) {
      document.addEventListener("fullscreenchange", this.onSlideshowFullscreenChange);
    }
    if (this.contentEl.requestFullscreen) {
      const request = this.contentEl.requestFullscreen();
      if (request && request.catch) request.catch(() => {});
    }
  }

  // The browser handles Escape itself while an element is fullscreen: it leaves
  // fullscreen and the key never reaches the view, so the way out has to hang
  // on this event. Every fullscreen change also resizes the pane, so the slide
  // is refitted for whatever size it has now.
  handleFullscreenChange() {
    if (typeof document === "undefined") return;
    if (this.slideshow && !document.fullscreenElement) {
      this.exitSlideshow();
      return;
    }
    this.applySettings();
  }

  exitSlideshow() {
    if (!this.slideshow) return;
    this.slideshow = false;
    if (typeof document !== "undefined" && document.removeEventListener) {
      document.removeEventListener("fullscreenchange", this.onSlideshowFullscreenChange);
    }
    this.contentEl.removeClass("ov-slideshow");
    if (this.slideshowBtn) this.slideshowBtn.removeClass("is-active");
    if (this.statusEl) this.statusEl.empty();
    if (this.controller && this.controller.emitStatus) this.controller.emitStatus();
    if (document.fullscreenElement && document.exitFullscreen) {
      const exit = document.exitFullscreen();
      if (exit && exit.catch) exit.catch(() => {});
    }
    // Back in the pane: recompute the fit now, and again after the fullscreen
    // exit lands, which the browser does asynchronously.
    this.applySettings();
    const raf = typeof requestAnimationFrame === "function" ? requestAnimationFrame : (cb) => setTimeout(cb, 16);
    raf(() => {
      if (!this.slideshow) this.applySettings();
    });
  }

  onSlideshowKey(ev) {
    if (!this.slideshow || !this.controller) return;
    const target = ev.target;
    if (target && (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)) return;
    switch (ev.key) {
      case "ArrowRight":
      case "ArrowDown":
      case "PageDown":
      case " ":
      case "Enter":
        ev.preventDefault();
        this.controller.nextSlide();
        break;
      case "ArrowLeft":
      case "ArrowUp":
      case "PageUp":
        ev.preventDefault();
        this.controller.previousSlide();
        break;
      case "Home":
        ev.preventDefault();
        if (this.controller.goToSlide) this.controller.goToSlide(0);
        break;
      case "End":
        ev.preventDefault();
        if (this.renderer || this.controller.setFit) {
          const count = this.controller.slideCount ? this.controller.slideCount() : null;
          if (this.controller.goToSlide && count) this.controller.goToSlide(count - 1);
        }
        break;
      case "Escape":
        ev.preventDefault();
        this.exitSlideshow();
        break;
      default:
        break;
    }
  }

  toggleMeta() {
    this.metaOpen = !this.metaOpen;
    this.metaEl.toggleClass("is-hidden", !this.metaOpen);
    this.metaBtn && this.metaBtn.toggleClass("is-active", this.metaOpen);
    if (this.metaOpen) this.renderMeta();
  }

  renderMeta() {
    if (!this.metaEl) return;
    this.metaEl.empty();
    const head = this.metaEl.createDiv("ov-meta-head");
    head.createSpan({ cls: "ov-meta-title", text: "File details" });
    // The buttons sit together at the right edge, in one group.
    const actions = head.createDiv("ov-meta-actions");
    const external = actions.createEl("button", { cls: "ov-tool-btn", attr: { "aria-label": "Open in the system editor", title: "Open in the system editor" } });
    setIcon(external, "external-link");
    external.addEventListener("click", () => this.openInDefaultApp());
    const copy = actions.createEl("button", { cls: "ov-tool-btn", attr: { "aria-label": "Copy the details", title: "Copy the details" } });
    setIcon(copy, "copy");
    copy.addEventListener("click", () => this.copyMeta());
    const close = actions.createEl("button", { cls: "ov-tool-btn", attr: { "aria-label": "Close", title: "Close" } });
    setIcon(close, "x");
    close.addEventListener("click", () => this.toggleMeta());

    const sections = [];
    const fileRows = [];
    if (this.file) fileRows.push({ label: "Name", value: this.file.name });
    if (this.format) fileRows.push({ label: "Format", value: this.format.label || this.format.ext });
    if (this.file) fileRows.push({ label: "Path", value: this.file.path });
    const stat = this.file && this.file.stat ? this.file.stat : null;
    if (stat && stat.size != null) fileRows.push({ label: "Size", value: formatSize(stat.size) });
    if (stat && stat.mtime) fileRows.push({ label: "File modified", value: new Date(stat.mtime).toLocaleString() });
    if (fileRows.length) sections.push({ title: "File", rows: fileRows });

    try {
      if (this.controller && typeof this.controller.describe === "function") {
        const contents = this.controller.describe();
        if (contents && contents.length) sections.push({ title: "Contents", rows: contents });
      }
    } catch (err) {
      // A broken model still gets the file details below.
    }

    // Everything the package itself says: core and application properties,
    // custom properties, language, comments authors, protection, media sizes.
    try {
      const model = this.controller ? (this.controller.model || this.controller.book) : null;
      if (model && model.pkg) {
        const kind = this.format && this.format.kind ? this.format.kind : "";
        const ext = this.format && this.format.ext ? this.format.ext : "";
        for (const entry of describePackage(model.pkg, { kind, ext, model })) sections.push(entry);
      }
    } catch (err) {
      // Metadata is a bonus; it must never break the panel.
    }
    this.metaSections = sections;

    for (const entry of sections) {
      this.metaEl.createDiv("ov-meta-section").setText(entry.title);
      const list = this.metaEl.createDiv("ov-meta-list");
      for (const row of entry.rows) {
        const line = list.createDiv("ov-meta-row");
        line.createDiv("ov-meta-label").setText(row.label);
        const value = line.createDiv("ov-meta-value");
        value.setText(row.value == null || row.value === "" ? "\u2013" : String(row.value));
        value.title = String(row.value == null ? "" : row.value);
      }
    }
    if (!sections.length) this.metaEl.createDiv("ov-meta-empty").setText("No details available for this file.");
  }

  // Copies every section as plain text, which is what pasting into a note or a
  // bug report wants. The panel and the viewer stay read only.
  copyMeta() {
    const text = sectionsToText(this.metaSections || []);
    if (!text) return;
    try {
      if (navigator && navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text);
        return;
      }
    } catch (err) {
      // Fall through to the element based copy.
    }
    try {
      const area = document.createElement("textarea");
      area.value = text;
      document.body.appendChild(area);
      area.select();
      document.execCommand("copy");
      document.body.removeChild(area);
    } catch (err) {
      // The panel still shows everything.
    }
  }

  refreshMeta() {
    if (this.metaOpen) this.renderMeta();
  }

  toggleOutline() {
    this.outlinePanelOpen = !this.outlinePanelOpen;
    this.outlineEl.toggleClass("is-hidden", !this.outlinePanelOpen);
    this.outlineBtn && this.outlineBtn.toggleClass("is-active", this.outlinePanelOpen);
  }

  // ---------- shared actions ----------

  currentZoom() {
    if (!this.format) return 100;
    if (this.format.kind === "spreadsheet") {
      const z = this.plugin.settings.zoom;
      return typeof z === "number" ? z : 100;
    }
    if (this.format.kind === "presentation") {
      const z = this.plugin.settings.slideZoom;
      return typeof z === "number" ? z : 100;
    }
    const z = this.plugin.settings.documentZoom;
    return typeof z === "number" ? z : 100;
  }

  setZoom(value) {
    if (this.controller && typeof this.controller.setZoom === "function") {
      this.controller.setZoom(value);
      this.refreshToolbarState();
    }
  }

  // Touchpad pinch: Chromium reports it as a wheel event with ctrlKey. Steps in
  // tens like the buttons, once per accumulated gesture so a fast pinch does not
  // write the settings on every event. The inversions setting flips the sign for
  // a reader who expects the other direction.
  onWheelZoom(ev) {
    if (!ev || !ev.ctrlKey) return;
    if (!(this.plugin && this.plugin.settings && this.plugin.settings.pinchZoom !== false)) return;
    if (!this.controller || typeof this.controller.setZoom !== "function") return;
    if (typeof ev.preventDefault === "function") ev.preventDefault();
    const invert = Boolean(this.plugin.settings.pinchInvert);
    const delta = (ev.deltaY || 0) * (invert ? -1 : 1);
    this.pinchAccum = (this.pinchAccum || 0) + delta;
    if (Math.abs(this.pinchAccum) < 16) return;
    const step = this.pinchAccum < 0 ? 10 : -10;
    this.pinchAccum = 0;
    this.setZoom(this.currentZoom() + step);
  }

  applySettings() {
    if (this.controller && typeof this.controller.applySettings === "function") {
      this.controller.applySettings();
    }
    this.applyStatusMode();
    this.refreshFilterBar();
    this.refreshToolbarState();
  }

  // True when the open file is a spreadsheet with a live grid.
  canCopy() {
    return Boolean(this.controller && this.format && this.format.kind === "spreadsheet"
      && typeof this.controller.copySelection === "function");
  }

  copySelection(mode) {
    if (!this.canCopy()) return false;
    return this.controller.copySelection(mode);
  }

  resetSizes() {
    if (!this.canCopy() || typeof this.controller.resetSizes !== "function") return false;
    this.controller.resetSizes();
    return true;
  }

  selectAllCells() {
    if (!this.canCopy() || typeof this.controller.selectAll !== "function") return false;
    this.controller.selectAll();
    return true;
  }

  runSearch(query) {
    if (!this.controller) return;
    const count = this.controller.search(query);
    this.searchCountEl.setText(query ? count + (count === 1 ? " match" : " matches") : "");
    this.setSearchNav(count);
  }

  stepSearch(dir) {
    if (!this.controller) return;
    const result = this.controller.stepSearch(dir);
    if (result) this.searchCountEl.setText(result.index + " of " + result.count);
    this.setSearchNav(result ? result.count : 0);
  }

  // The finder chevrons only earn their place when the search found something.
  setSearchNav(count) {
    if (!this.searchNavEl) return;
    this.searchNavEl.toggleClass("is-hidden", !(count > 0));
  }

  // The selection aware menu for documents and presentations. The grid asks the
  // spreadsheet controller for its items instead.
  onDocumentContextMenu(ev) {
    if (!this.controller || !this.format) return;
    if (this.format.kind === "spreadsheet") return;
    const target = ev.target;
    if (target && (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)) return;
    ev.preventDefault();
    ev.stopPropagation();
    const selected = String(window.getSelection ? window.getSelection().toString() : "").trim();
    const items = [];
    if (selected) {
      const preview = selected.length > 24 ? selected.slice(0, 24) + "\u2026" : selected;
      items.push({
        title: "Copy \u201c" + preview.replace(/\s+/g, " ") + "\u201d",
        icon: "copy",
        action: () => this.copyText(selected),
      });
      items.push({
        title: "Search this document for the selection",
        icon: "search",
        action: () => {
          if (!this.searchInput) return;
          this.searchInput.value = selected.replace(/\s+/g, " ").slice(0, 60);
          this.runSearch(this.searchInput.value);
        },
      });
    }
    items.push({ title: "Select all text", icon: "text", action: () => this.selectAllText() });
    items.push({ title: "Show or hide the file details", icon: "info", action: () => this.toggleMeta() });
    items.push({ title: "Reload the file", icon: "refresh-cw", action: () => this.reload() });
    items.push({ title: "Open in the system editor", icon: "external-link", action: () => this.openInDefaultApp() });
    this.showContextMenu(ev, items.filter((item) => item.title !== ""));
  }

  copyText(text) {
    try {
      if (navigator && navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text);
        return;
      }
    } catch (err) {
      // Fall through to the clipboard command.
    }
    try {
      const area = document.createElement("textarea");
      area.value = text;
      document.body.appendChild(area);
      area.select();
      document.execCommand("copy");
      document.body.removeChild(area);
    } catch (err) {
      // Nothing else to try.
    }
  }

  selectAllText() {
    const host = this.hostEl;
    if (!host) return;
    try {
      const range = document.createRange();
      range.selectNodeContents(host);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
    } catch (err) {
      // Selection is a convenience; ignore a host that refuses it.
    }
  }

  showContextMenu(ev, items) {
    const menu = new Menu();
    for (const item of items) {
      menu.addItem((mi) => mi.setTitle(item.title).setIcon(item.icon).onClick(item.action));
    }
    menu.showAtMouseEvent(ev);
  }

  teardownController() {
    if (this.slideshow) this.exitSlideshow();
    if (!this.controller) return;
    const model = this.controller.model;
    this.controller.destroy();
    this.controller = null;
    // A cached model keeps its media alive for the next open; anything else
    // gives its object URLs back now.
    if (model && !model.__ovCached) releaseModel(model);
  }

  // ---------- states ----------

  renderEmpty() {
    if (!this.hostEl) return;
    this.hostEl.empty();
    const box = this.hostEl.createDiv("ov-message");
    box.createEl("h3", { text: "No document open" });
    box.createEl("p", { text: "Open an .xlsx, .docx, .pptx or one of the related formats to read it here." });
  }

  showUnsupported(reason) {
    this.hostEl.empty();
    const box = this.hostEl.createDiv("ov-message ov-message-warning");
    box.createEl("h3", { text: "This format is not supported yet" });
    box.createEl("p", {
      text: reason || (this.format && this.format.label) || "The viewer cannot read this file.",
    });
    const actions = box.createDiv("ov-message-actions");
    const external = actions.createEl("button", { cls: "ov-button", text: "Open in the default app" });
    external.addEventListener("click", () => this.openInDefaultApp());
    this.statusEl.setText("Unsupported format");
  }

  showError(err) {
    this.hostEl.empty();
    this.tabsEl.empty();
    this.filterBarEl.empty();
    const box = this.hostEl.createDiv("ov-message ov-message-error");
    box.createEl("h3", { text: "Could not read this file" });
    box.createEl("p", { text: err && err.message ? err.message : String(err) });
    box.createEl("p", {
      cls: "ov-message-hint",
      text: "The file may be damaged, password protected, or saved in a format the viewer does not read.",
    });
    const actions = box.createDiv("ov-message-actions");
    const retry = actions.createEl("button", { cls: "ov-button", text: "Try again" });
    retry.addEventListener("click", () => this.reload());
    const external = actions.createEl("button", { cls: "ov-button", text: "Open in the default app" });
    external.addEventListener("click", () => this.openInDefaultApp());
    this.statusEl.setText("Error");
  }

  openInDefaultApp() {
    if (!this.file) return;
    const adapter = this.app.vault.adapter;
    if (adapter && typeof adapter.getFullPath === "function") {
      const fullPath = adapter.getFullPath(this.file.path);
      try {
        const electron = require("electron");
        if (electron && electron.shell && electron.shell.openPath) {
          electron.shell.openPath(fullPath);
          return;
        }
      } catch (err) {
        // Not on desktop; fall through to the notice below.
      }
    }
    this.suggestEditor(null);
  }

  // The viewer never writes to a document, so anything that needs saving is
  // handed to an editor. This is also the fallback when the system cannot open
  // the file itself.
  suggestEditor(detail) {
    const { Notice } = require("obsidian");
    const hint = "A free desktop office editor can open and save this file.";
    new Notice((detail ? detail + " " : "This viewer reads files only. ") + hint, 8000);
  }

  async reload() {
    if (this.file) await this.renderFile(this.file);
  }
}

function formatSize(bytes) {
  const value = Number(bytes) || 0;
  if (value < 1024) return value + " bytes";
  if (value < 1024 * 1024) return (value / 1024).toFixed(1) + " KB";
  return (value / 1024 / 1024).toFixed(1) + " MB";
}

module.exports = { OfficeView, VIEW_TYPE };
