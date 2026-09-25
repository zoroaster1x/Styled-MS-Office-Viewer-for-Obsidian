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

// Spreadsheet controller: sheets become tabs, autofilter columns become chips,
// and the grid renderer does the drawing. This is the original XLSX Styled
// Viewer behaviour, now shared with csv, tsv and ods files.

const { readWorkbook } = require("../spreadsheet/read");
const { createRenderer } = require("../spreadsheet/render");
const { parseDelimited, buildSheetModel } = require("../delimited/csv");
const { readOds } = require("../odf/parse");

function isDelimitedFormat(format) {
  return format === "csv" || format === "tsv";
}

class SpreadsheetController {
  constructor(opts) {
    this.host = opts.host;
    this.tabsEl = opts.tabsEl;
    this.filterBarEl = opts.filterBarEl;
    this.plugin = opts.plugin;
    this.format = opts.format;
    this.book = null;
    this.sheetIndex = 0;
    this.filterStates = new Map();
    this.outlineStates = new Map();
    this.scrollPositions = new Map();
    this.renderer = null;
    this.savedFilters = null;
    this.showFilterBar = true;
    this.selection = null;
    this.callbacks = opts.callbacks || {};
  }

  // The sizes remembered for the sheet on screen.
  savedSizesFor(sheetName) {
    const saved = this.savedState || {};
    const all = saved.sizes || {};
    return all[sheetName] || null;
  }

  rememberSizes(sizes) {
    if (!this.filePath || !this.book) return;
    const sheetName = this.book.sheets[this.sheetIndex] ? this.book.sheets[this.sheetIndex].name : String(this.sheetIndex);
    const existing = (this.savedState || {}).sizes || {};
    existing[sheetName] = sizes;
    this.savedState = Object.assign({}, this.savedState, { sizes: existing });
    this.plugin.saveFileState(this.filePath, this.savedState);
  }

  load(bytes, fileName) {
    this.book = null;
    if (isDelimitedFormat(this.format)) {
      const text = new TextDecoder("utf-8").decode(bytes);
      const delimiter = this.format === "tsv" ? "\t" : null;
      const rows = parseDelimited(text, delimiter);
      const model = buildSheetModel(rows, trimExtension(fileName));
      this.book = {
        kind: "delimited",
        sheets: [{ name: model.name, index: 0, state: "visible", sheetId: 1, path: model.name }],
        loadSheet: () => model,
        sheetIndexByName: () => 0,
        styles: model.styles,
        theme: null,
      };
      return;
    }
    if (this.format === "ods") {
      this.book = readOds(bytes);
      return;
    }
    this.book = readWorkbook(bytes);
  }

  restoreState(saved) {
    if (!saved) return;
    this.showFilterBar = saved.showFilterBar !== false;
    this.savedFilters = saved.filters || null;
    this.sheetIndex = this.firstVisibleSheet();
    if (Number.isInteger(saved.sheetIndex) && this.book.sheets[saved.sheetIndex]
      && this.book.sheets[saved.sheetIndex].state === "visible") {
      this.sheetIndex = saved.sheetIndex;
    }
  }

  saveState() {
    if (!this.book || !this.filePath) return;
    const sheetName = this.book.sheets[this.sheetIndex] ? this.book.sheets[this.sheetIndex].name : String(this.sheetIndex);
    const existing = this.savedState || {};
    const filters = Object.assign({}, existing.filters || {});
    filters[sheetName] = serializeFilters(this.filterStateFor(this.sheetIndex));
    this.savedState = {
      showFilterBar: this.showFilterBar !== false,
      filters,
      sheetIndex: this.sheetIndex,
    };
    this.plugin.saveFileState(this.filePath, this.savedState);
  }

  firstVisibleSheet() {
    if (!this.book) return 0;
    for (let i = 0; i < this.book.sheets.length; i++) {
      if (this.book.sheets[i].state === "visible") return i;
    }
    return 0;
  }

  visibleSheets() {
    if (!this.book) return [];
    return this.book.sheets.filter((sheet) => sheet.state === "visible");
  }

  mount() {
    this.buildTabs();
    this.renderSheet();
  }

  buildTabs() {
    if (!this.tabsEl) return;
    this.tabsEl.empty();
    const sheets = this.visibleSheets();
    this.tabsEl.toggleClass("is-hidden", sheets.length <= 1);
    for (const sheet of sheets) {
      const tab = this.tabsEl.createEl("button", { cls: "ov-tab" });
      tab.setText(sheet.name);
      tab.toggleClass("is-active", sheet.index === this.sheetIndex);
      tab.addEventListener("click", () => this.switchSheet(sheet.index));
    }
  }

  updateTabs() {
    if (!this.tabsEl) return;
    const tabs = this.tabsEl.querySelectorAll(".ov-tab");
    const sheets = this.visibleSheets();
    for (let i = 0; i < tabs.length && i < sheets.length; i++) {
      tabs[i].toggleClass("is-active", sheets[i].index === this.sheetIndex);
    }
  }

  switchSheet(index, ref) {
    if (!this.book || index === this.sheetIndex) {
      if (ref && this.renderer) this.renderer.scrollToCell(ref);
      return;
    }
    if (this.renderer) this.scrollPositions.set(this.sheetIndex, this.renderer.getScrollPosition());
    this.saveState();
    this.sheetIndex = index;
    this.renderSheet();
    if (ref && this.renderer) this.renderer.scrollToCell(ref);
    const saved = this.scrollPositions.get(index);
    if (!ref && saved && this.renderer) this.renderer.setScrollPosition(saved);
    this.saveState();
  }

  filterStateFor(index) {
    if (!this.filterStates.has(index)) this.filterStates.set(index, new Map());
    return this.filterStates.get(index);
  }

  outlineStateFor(index) {
    if (!this.outlineStates.has(index)) this.outlineStates.set(index, new Map());
    return this.outlineStates.get(index);
  }

  renderSheet() {
    this.teardown();
    const model = this.book.loadSheet(this.sheetIndex);
    const filterState = this.filterStateFor(this.sheetIndex);
    const savedForSheet = this.savedFilters ? this.savedFilters[model.name] : null;
    if (savedForSheet) {
      filterState.clear();
      for (const [colId, sel] of deserializeFilters(savedForSheet)) filterState.set(colId, sel);
    } else if (filterState.size === 0 && model.autoFilter) {
      for (const [colId, def] of model.autoFilter.columns) {
        if (def.values) {
          const set = new Set(def.values);
          if (def.blank) set.add("");
          filterState.set(colId, set.size ? set : null);
        } else if (def.custom) {
          filterState.set(colId, { custom: def.custom });
        }
      }
    }
    const styles = this.book.styles || model.styles;
    this.renderer = createRenderer({
      container: this.host,
      sheet: model,
      styles,
      theme: this.book.theme,
      date1904: this.book.date1904,
      settings: this.settings(),
      filterState,
      outlineCollapsed: this.outlineStateFor(this.sheetIndex),
      onNavigate: (sheetName, ref) => {
        const target = this.book.sheetIndexByName(sheetName);
        if (target === -1) return;
        this.switchSheet(target, ref);
      },
      onOpenExternal: (url) => this.openExternal(url),
      showContextMenu: this.callbacks.showContextMenu || null,
      onSelect: (sel) => {
        this.selection = sel;
        this.emitStatus();
      },
      onFilterChange: () => {
        this.callbacks.onFilterChange ? this.callbacks.onFilterChange() : null;
        this.emitStatus();
        this.saveState();
      },
      sizeOverrides: this.savedSizesFor(model.name),
      onSizesChanged: (sizes) => this.rememberSizes(sizes),
    });
    this.callbacks.onFiltersChanged ? this.callbacks.onFiltersChanged() : null;
    this.updateTabs();
    this.emitStatus();
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

  settings() {
    const pluginSettings = this.plugin.settings;
    return {
      sheetBackground: pluginSettings.sheetBackground === "theme" ? "theme" : "white",
      minColumnWidth: Number(pluginSettings.minColumnWidth) > 0 ? Number(pluginSettings.minColumnWidth) : 0,
      minRowHeight: Number(pluginSettings.minRowHeight) > 0 ? Number(pluginSettings.minRowHeight) : 0,
      showGridlines: pluginSettings.showGridlines !== false,
      showHeaders: pluginSettings.showHeaders !== false,
      maxRows: Number(pluginSettings.maxRows) > 0 ? Number(pluginSettings.maxRows) : 0,
      scale: this.currentZoom() / 100,
    };
  }

  currentZoom() {
    const z = this.plugin.settings.zoom;
    return typeof z === "number" && z >= 40 && z <= 400 ? z : 100;
  }

  applySettings() {
    if (!this.renderer) return;
    this.renderer.setSettings(this.settings());
    this.renderer.render(true);
    this.emitStatus();
  }

  // Filter chips live under the tabs, exactly as in the original plugin.
  refreshFilterChips(filterBarEl, limit) {
    if (!filterBarEl) return;
    filterBarEl.empty();
    if (!this.renderer) return;
    const state = this.renderer.getState();
    let groups = 0;
    for (const filter of state.filters) {
      const valueCount = filter.values.length + (filter.blanks ? 1 : 0);
      if (valueCount > limit) continue;
      if (filter.values.length < 2) continue;
      groups++;
      const group = filterBarEl.createDiv("ov-chip-group");
      const label = group.createSpan("ov-chip-label");
      label.setText(filter.letter + (filter.headerText ? ": " + filter.headerText : ""));
      const selected = filter.selected instanceof Array ? new Set(filter.selected) : null;
      const isAll = filter.all;
      const allValues = filter.values.map((v) => v.value);
      if (filter.blanks) allValues.push("");
      const toggleValue = (value) => {
        let set;
        if (selected) set = new Set(selected);
        else set = new Set(allValues);
        if (set.has(value)) set.delete(value);
        else set.add(value);
        this.renderer.setFilter(filter.colId, set.size === allValues.length ? null : set);
        this.callbacks.onFiltersChanged ? this.callbacks.onFiltersChanged() : null;
        this.emitStatus();
        this.saveState();
      };
      for (const value of filter.values) {
        const chip = group.createEl("button", { cls: "ov-chip" });
        chip.setText(value.value);
        chip.toggleClass("is-off", !(isAll || (selected && selected.has(value.value))));
        chip.addEventListener("click", () => toggleValue(value.value));
      }
      if (filter.blanks) {
        const chip = group.createEl("button", { cls: "ov-chip" });
        chip.setText("(Blanks)");
        chip.toggleClass("is-off", !(isAll || (selected && selected.has(""))));
        chip.addEventListener("click", () => toggleValue(""));
      }
      const allBtn = group.createEl("button", { cls: "ov-chip ov-chip-all" });
      allBtn.setText("all");
      allBtn.addEventListener("click", () => {
        this.renderer.setFilter(filter.colId, null);
        this.callbacks.onFiltersChanged ? this.callbacks.onFiltersChanged() : null;
        this.emitStatus();
        this.saveState();
      });
      const noneBtn = group.createEl("button", { cls: "ov-chip ov-chip-all" });
      noneBtn.setText("none");
      noneBtn.addEventListener("click", () => {
        this.renderer.setFilter(filter.colId, new Set());
        this.callbacks.onFiltersChanged ? this.callbacks.onFiltersChanged() : null;
        this.emitStatus();
        this.saveState();
      });
    }
    const hidden = this.showFilterBar === false || groups === 0;
    filterBarEl.toggleClass("is-hidden", hidden);
  }

  // ---------- common controller surface ----------

  search(query) {
    return this.renderer ? this.renderer.search(query) : 0;
  }

  stepSearch(dir) {
    return this.renderer ? this.renderer.searchNext(dir) : null;
  }

  setZoom(value) {
    const zoom = Math.max(40, Math.min(400, Math.round(value / 10) * 10));
    this.plugin.updateSetting("zoom", zoom);
    this.applySettings();
  }

  toggleBackground() {
    const next = this.plugin.settings.sheetBackground === "white" ? "theme" : "white";
    this.plugin.updateSetting("sheetBackground", next);
    this.applySettings();
  }

  toggleFilterBar() {
    this.showFilterBar = !this.showFilterBar;
    this.saveState();
    this.callbacks.onFiltersChanged ? this.callbacks.onFiltersChanged() : null;
  }

  // Clipboard actions, used by the commands and by the grid's own keybindings.
  copySelection(mode) {
    if (!this.renderer || typeof this.renderer.copySelection !== "function") return false;
    return this.renderer.copySelection(mode);
  }

  selectAll() {
    if (this.renderer && typeof this.renderer.selectAll === "function") this.renderer.selectAll();
  }

  resetSizes() {
    if (this.renderer && typeof this.renderer.resetSizes === "function") this.renderer.resetSizes();
  }

  clearFilters() {
    if (!this.renderer) return;
    this.renderer.clearFilters();
    this.callbacks.onFiltersChanged ? this.callbacks.onFiltersChanged() : null;
    this.emitStatus();
    this.saveState();
  }

  emitStatus() {
    this.callbacks.onStatus ? this.callbacks.onStatus(this.statusText()) : null;
  }

  describe() {
    const rows = [];
    const book = this.book;
    if (!book) return rows;
    const properties = book.properties || {};
    if (properties.title) rows.push({ label: "Title", value: properties.title });
    if (properties.author) rows.push({ label: "Author", value: properties.author });
    if (properties.application) rows.push({ label: "Application", value: properties.application });
    if (properties.company) rows.push({ label: "Company", value: properties.company });
    if (properties.created) rows.push({ label: "Created", value: formatDate(properties.created) });
    if (properties.modified) rows.push({ label: "Modified", value: formatDate(properties.modified) });
    const visible = this.visibleSheets();
    rows.push({ label: "Sheets", value: visible.map((sheet) => sheet.name).join(", ") });
    if (book.sheets.length !== visible.length) {
      rows.push({ label: "Hidden sheets", value: String(book.sheets.length - visible.length) });
    }
    if (this.renderer) {
      const state = this.renderer.getState();
      rows.push({ label: "Active sheet size", value: state.totalRows + " rows x " + state.totalCols + " columns" });
      if (state.frozenRows || state.frozenCols) {
        rows.push({ label: "Frozen panes", value: state.frozenRows + " rows, " + state.frozenCols + " columns" });
      }
      const active = state.filters.filter((filter) => !filter.all);
      if (active.length) rows.push({ label: "Filters active", value: String(active.length) });
    }
    const sizes = this.renderer && typeof this.renderer.getSizes === "function" ? this.renderer.getSizes() : null;
    if (sizes) {
      const columns = Object.keys(sizes.cols).length;
      const rowCount = Object.keys(sizes.rows).length;
      if (columns || rowCount) {
        rows.push({ label: "Resized by hand", value: columns + " columns, " + rowCount + " rows" });
      }
    }
    return rows;
  }

  statusText() {
    if (!this.renderer) return "";
    const state = this.renderer.getState();
    const parts = [];
    parts.push("Sheet: " + state.sheetName);
    parts.push(state.totalRows + " rows x " + state.totalCols + " cols");
    if (state.visibleRows !== state.totalRows || state.visibleCols !== state.totalCols) {
      parts.push(state.visibleRows + " x " + state.visibleCols + " visible");
    }
    const active = state.filters.filter((f) => !f.all);
    if (active.length) parts.push(active.length + " filter" + (active.length === 1 ? "" : "s") + " active");
    if (this.selection && this.selection.ref) {
      parts.push("Selected " + this.selection.ref + " (" + this.selection.text.replace(/\s+/g, " ").slice(0, 60) + ")");
    }
    return parts.join("  |  ");
  }

  getScrollPosition() {
    return this.renderer ? this.renderer.getScrollPosition() : null;
  }

  teardown() {
    if (this.renderer) {
      this.renderer.destroy();
      this.renderer = null;
    }
  }

  destroy() {
    this.teardown();
    this.book = null;
    this.filterStates = new Map();
    this.outlineStates = new Map();
    this.scrollPositions = new Map();
  }
}

function serializeFilters(map) {
  const out = {};
  for (const [colId, sel] of map) {
    if (sel instanceof Set) out[colId] = { values: Array.from(sel) };
    else if (sel && sel.custom) out[colId] = { custom: sel.custom };
    else out[colId] = { all: true };
  }
  return out;
}

function deserializeFilters(data) {
  const map = new Map();
  if (!data) return map;
  for (const key of Object.keys(data)) {
    const colId = Number(key);
    const entry = data[key];
    if (!entry || entry.all) map.set(colId, null);
    else if (entry.values) map.set(colId, new Set(entry.values));
    else if (entry.custom) map.set(colId, { custom: entry.custom });
  }
  return map;
}

function formatDate(value) {
  const text = String(value || "");
  const match = /^(\d{4})-(\d{2})-(\d{2})T?(\d{2}:\d{2})?/.exec(text);
  if (!match) return text;
  return match[1] + "-" + match[2] + "-" + match[3] + (match[4] ? " " + match[4] : "");
}

function trimExtension(name) {
  const dot = name.lastIndexOf(".");
  return dot === -1 ? name : name.slice(0, dot);
}

module.exports = { SpreadsheetController, isDelimitedFormat };
