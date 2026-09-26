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

// Document controller: parses a docx, odt or rtf file and drives the document
// renderer, the outline panel and the page navigation.

const { parseDocx } = require("../docx/parse");
const { createDocxRenderer } = require("../docx/render");
const { readOdt } = require("../odf/parse");
const { readRtf } = require("../rtf/parse");
const { readDoc } = require("../legacy/doc");

class DocumentController {
  constructor(opts) {
    this.host = opts.host;
    this.plugin = opts.plugin;
    this.format = opts.format;
    this.model = null;
    this.renderer = null;
    this.outline = [];
    this.callbacks = opts.callbacks || {};
  }

  load(bytes, fileName) {
    if (this.format === "odt") {
      this.model = readOdt(bytes);
      // ODT images are referenced by href; resolving them needs the package.
      this.model.mediaUrl = (href) => resolveOdfHref(this.model.pkg, href);
      return;
    }
    if (this.format === "rtf") {
      this.model = readRtf(new TextDecoder("utf-8").decode(bytes));
      return;
    }
    if (this.format === "doc") {
      // Word 97 text: paragraphs and words out of the compound file.
      this.model = readDoc(bytes);
      return;
    }
    this.model = parseDocx(bytes);
  }

  mount() {
    if (!this.model) return;
    // Selection is a DOM affair, so it is watched here and folded into the
    // status line: words, characters, headings and pages.
    this.selectionHandler = () => this.emitStatus();
    const selectionDoc = (this.host.ownerDocument || document);
    selectionDoc.addEventListener("selectionchange", this.selectionHandler);
    this.renderer = createDocxRenderer({
      container: this.host,
      model: this.model,
      settings: this.settings(),
      onInternalLink: (anchor) => {
        if (this.renderer) this.renderer.scrollToAnchor(anchor);
      },
      onExternalLink: (href) => this.openExternal(href),
      onReady: (info) => {
        this.outline = info.outline || [];
        this.renderedPageCount = info.pageCount || 0;
        this.callbacks.onOutline ? this.callbacks.onOutline(this.outline) : null;
        this.emitStatus();
      },
      onPageChange: (info) => {
        this.currentPage = info.page || 1;
        this.pageTotal = info.count || 0;
        this.callbacks.onPage ? this.callbacks.onPage(info) : null;
      },
    });
  }

  settings() {
    return {
      zoom: this.currentZoom() / 100,
      pageBackground: this.plugin.settings.pageBackground === "theme" ? "theme" : "white",
      showHeaders: this.plugin.settings.showHeaders !== false,
      showPageNumbers: this.plugin.settings.showPageNumbers !== false,
    };
  }

  currentZoom() {
    const z = this.plugin.settings.documentZoom;
    return typeof z === "number" && z >= 40 && z <= 400 ? z : 100;
  }

  applySettings() {
    if (!this.renderer) return;
    this.renderer.setSettings(this.settings());
  }

  search(query) {
    return this.renderer ? this.renderer.search(query) : 0;
  }

  stepSearch(dir) {
    return this.renderer ? this.renderer.searchNext(dir) : null;
  }

  setZoom(value) {
    const zoom = Math.max(40, Math.min(400, Math.round(value / 10) * 10));
    this.plugin.updateSetting("documentZoom", zoom);
    this.renderer ? this.renderer.setSettings({ zoom: zoom / 100 }) : null;
    this.emitStatus();
  }

  toggleBackground() {
    const next = this.plugin.settings.pageBackground === "white" ? "theme" : "white";
    this.plugin.updateSetting("pageBackground", next);
    this.applySettings();
  }

  emitStatus() {
    this.callbacks.onStatus ? this.callbacks.onStatus(this.statusText()) : null;
  }

  statusText() {
    if (!this.model) return "";
    const parts = [];
    const selected = this.selectedText();
    const words = countWords(this.model);
    const characters = countCharacters(this.model);
    if (selected) {
      const selectedWords = countWordsIn(selected);
      parts.push(formatCount(selectedWords, "word") + ", " + formatCount(selected.length, "character") + " selected");
      if (words) parts.push(formatCount(words, "word") + " in the document");
    } else {
      if (words) parts.push(formatCount(words, "word"));
      if (characters) parts.push(formatCount(characters, "character"));
    }
    if (this.outline && this.outline.length) parts.push(this.outline.length + " headings");
    if (this.renderedPageCount) parts.push(formatCount(this.renderedPageCount, "page"));
    const properties = this.model.properties || {};
    if (properties.application) parts.push(properties.application);
    return parts.join("  |  ");
  }

  // What the reader has highlighted, if anything.
  selectedText() {
    try {
      const selection = (this.host.ownerDocument || document).getSelection();
      if (!selection || selection.isCollapsed) return "";
      const text = selection.toString();
      return text && text.trim() ? text : "";
    } catch (err) {
      return "";
    }
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

  // Everything the metadata panel shows for this document.
  describe() {
    const rows = [];
    const properties = (this.model && this.model.properties) || {};
    if (properties.title) rows.push({ label: "Title", value: properties.title });
    if (properties.author) rows.push({ label: "Author", value: properties.author });
    if (properties.application) rows.push({ label: "Application", value: properties.application });
    if (properties.created) rows.push({ label: "Created", value: formatDate(properties.created) });
    if (properties.modified) rows.push({ label: "Modified", value: formatDate(properties.modified) });
    const words = this.model ? countWords(this.model) : 0;
    const characters = this.model ? countCharacters(this.model) : 0;
    if (words) rows.push({ label: "Words", value: words.toLocaleString("en-GB") });
    if (characters) rows.push({ label: "Characters", value: characters.toLocaleString("en-GB") });
    if (this.renderedPageCount) rows.push({ label: "Pages (rendered)", value: String(this.renderedPageCount) });
    if (properties.pages) rows.push({ label: "Pages (as saved)", value: properties.pages });
    if (this.outline && this.outline.length) rows.push({ label: "Headings", value: String(this.outline.length) });
    const section = (this.model && this.model.section) || null;
    if (section && section.pageWidthTw && section.pageHeightTw) {
      rows.push({ label: "Page size", value: Math.round(section.pageWidthTw / 20) + " x " + Math.round(section.pageHeightTw / 20) + " pt" });
    }
    if (this.model) {
      const footnotes = this.model.footnotes ? this.model.footnotes.size : 0;
      const endnotes = this.model.endnotes ? this.model.endnotes.size : 0;
      if (footnotes || endnotes) rows.push({ label: "Notes", value: footnotes + " footnotes, " + endnotes + " endnotes" });
      const media = this.model.pkg ? this.model.pkg.list().filter((name) => /^word\/media\//.test(name)).length : 0;
      if (media) rows.push({ label: "Media parts", value: String(media) });
    }
    return rows;
  }

  getOutline() {
    return this.outline;
  }

  scrollToAnchor(anchor) {
    if (this.renderer) this.renderer.scrollToAnchor(anchor);
  }

  // Mounts a model that came out of the document cache instead of parsing the
  // bytes again. The model owns its media URLs; the cache keeps them alive.
  adoptModel(model) {
    this.model = model;
    return model;
  }

  destroy() {
    if (this.selectionHandler) {
      const selectionDoc = (this.host.ownerDocument || document);
      selectionDoc.removeEventListener("selectionchange", this.selectionHandler);
      this.selectionHandler = null;
    }
    if (this.renderer) {
      this.renderer.destroy();
      this.renderer = null;
    }
    this.model = null;
  }
}

// Word writes ISO timestamps; show just the date and time.
function formatDate(value) {
  const text = String(value || "");
  const match = /^(\d{4})-(\d{2})-(\d{2})T?(\d{2}:\d{2})?/.exec(text);
  if (!match) return text;
  return match[1] + "-" + match[2] + "-" + match[3] + (match[4] ? " " + match[4] : "");
}

function formatCount(value, noun) {
  const shown = Number(value).toLocaleString("en-GB");
  return shown + " " + noun + (Number(value) === 1 ? "" : "s");
}

function countWordsIn(text) {
  const trimmed = String(text || "").trim();
  if (!trimmed) return 0;
  return trimmed.split(/\s+/).length;
}

function countCharacters(model) {
  let total = 0;
  const walk = (blocks) => {
    for (const block of blocks || []) {
      if (block.type === "p") total += plainText(block.runs).length;
      else if (block.type === "table") {
        for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks);
      }
    }
  };
  walk(model.body);
  return total;
}

function countWords(model) {
  let n = 0;
  const walk = (blocks) => {
    for (const block of blocks || []) {
      if (block.type === "p") {
        const text = plainText(block.runs);
        if (text.trim()) n += text.trim().split(/\s+/).length;
      } else if (block.type === "table") {
        for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks);
      }
    }
  };
  walk(model.body);
  return n;
}

function plainText(runs) {
  let out = "";
  for (const run of runs || []) {
    if (!run) continue;
    if (run.type === "text") out += run.text + " ";
    else if (run.type === "run") out += plainText(run.runs);
    else if (run.type === "link") out += plainText(run.link.runs);
    else if (run.type === "shapegroup" && run.texts) out += run.texts.join(" ") + " ";
    else if (run.type === "vml" && run.texts) out += run.texts.join(" ") + " ";
  }
  return out;
}

function resolveOdfHref(pkg, href) {
  if (!href) return null;
  const path = href.replace(/^\.\//, "");
  const bytes = pkg.bytes(path);
  if (!bytes) return null;
  const mime = pkg.mimeOf(path);
  try {
    let binary = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return "data:" + mime + ";base64," + btoa(binary);
  } catch (err) {
    return null;
  }
}

module.exports = { DocumentController };
