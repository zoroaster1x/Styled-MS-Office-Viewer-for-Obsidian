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

const { Plugin, PluginSettingTab, Setting } = require("obsidian");
const { OfficeView, VIEW_TYPE } = require("./app/view");
const { OFFICE_EXTENSIONS } = require("./app/format-router");

const DEFAULT_SETTINGS = {
  // Spreadsheets
  sheetBackground: "white",
  showGridlines: true,
  showHeaders: true,
  zoom: 100,
  maxRows: 2000,
  minColumnWidth: 0,
  minRowHeight: 0,
  chipLimit: 10,
  // Documents
  documentZoom: 100,
  pageBackground: "white",
  showPageNumbers: true,
  // Presentations
  slideZoom: 100,
  slideFit: "contain",
  showSpeakerNotes: true,
  // View
  statusInView: false,
  pinchZoom: true,
  pinchInvert: false,
  // Per file state: active sheet, filters, slide position.
  fileState: {},
};

class OfficeViewerPlugin extends Plugin {
  async onload() {
    // One item in the app's own status bar, so the view keeps its full height
    // for the document instead of reserving a row for a counter.
    this.statusBarEl = this.addStatusBarItem ? this.addStatusBarItem() : null;
    await this.loadSettings();

    this.registerView(VIEW_TYPE, (leaf) => new OfficeView(leaf, this));
    this.registerExtensions(OFFICE_EXTENSIONS, VIEW_TYPE);
    this.addSettingTab(new OfficeViewerSettingTab(this.app, this));

    this.addCommand({
      id: "reload-document",
      name: "Reload the current document",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(OfficeView);
        if (!view) return false;
        if (!checking) view.reload();
        return true;
      },
    });

    this.addCommand({
      id: "copy-selection-with-formatting",
      name: "Copy the selected cells with formatting",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(OfficeView);
        if (!view || !view.canCopy()) return false;
        if (!checking) view.copySelection("rich");
        return true;
      },
    });

    this.addCommand({
      id: "copy-selection-plain",
      name: "Copy the selected cells as plain values",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(OfficeView);
        if (!view || !view.canCopy()) return false;
        if (!checking) view.copySelection("plain");
        return true;
      },
    });

    this.addCommand({
      id: "select-all-cells",
      name: "Select every visible cell",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(OfficeView);
        if (!view || !view.canCopy()) return false;
        if (!checking) view.selectAllCells();
        return true;
      },
    });

    this.addCommand({
      id: "toggle-slideshow",
      name: "Start or stop the slideshow",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(OfficeView);
        if (!view || !view.format || view.format.kind !== "presentation") return false;
        if (!checking) view.toggleSlideshow();
        return true;
      },
    });

    this.addCommand({
      id: "open-in-system-editor",
      name: "Open the current file in the system editor",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(OfficeView);
        if (!view || !view.file) return false;
        if (!checking) view.openInDefaultApp();
        return true;
      },
    });

    this.addCommand({
      id: "toggle-file-details",
      name: "Show or hide the file details",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(OfficeView);
        if (!view) return false;
        if (!checking) view.toggleMeta();
        return true;
      },
    });

    this.addCommand({
      id: "reset-column-row-sizes",
      name: "Reset the column widths and row heights for this file",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(OfficeView);
        if (!view || !view.canCopy()) return false;
        if (!checking) view.resetSizes();
        return true;
      },
    });

    this.addCommand({
      id: "toggle-background",
      name: "Toggle the page background between white and the theme",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(OfficeView);
        if (!view) return false;
        if (!checking) {
          const next = this.settings.sheetBackground === "white" ? "theme" : "white";
          this.updateSetting("sheetBackground", next);
          this.updateSetting("pageBackground", next);
        }
        return true;
      },
    });
  }

  onunload() {
    // Leaves are closed by Obsidian.
    if (this.statusBarEl && this.statusBarEl.setText) this.statusBarEl.setText("");
  }

  async loadSettings() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
  }

  async saveSettings() {
    await this.saveData(this.settings);
  }

  // Updates a setting and refreshes every open view so the change is visible
  // without reopening the file.
  async updateSetting(key, value) {
    this.settings[key] = value;
    await this.saveSettings();
    for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE)) {
      const view = leaf.view;
      if (view instanceof OfficeView && typeof view.applySettings === "function") {
        view.applySettings();
      }
    }
  }

  // Per file view state (active sheet, filters, slide). Saved quietly, so the
  // other open views are not re-rendered on every chip click.
  async saveFileState(path, state) {
    if (!this.settings.fileState) this.settings.fileState = {};
    this.settings.fileState[path] = state;
    await this.saveSettings();
  }
}

class OfficeViewerSettingTab extends PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display() {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.createEl("h2", { text: "Office Viewer" });

    containerEl.createEl("h3", { text: "Spreadsheets" });

    new Setting(containerEl)
      .setName("Sheet background")
      .setDesc("White keeps a sheet looking like Excel even in a dark theme. Theme follows the Obsidian colours for cells that have no fill of their own.")
      .addDropdown((dropdown) =>
        dropdown
          .addOption("white", "Always white (recommended)")
          .addOption("theme", "Follow the Obsidian theme")
          .setValue(this.plugin.settings.sheetBackground)
          .onChange((value) => this.plugin.updateSetting("sheetBackground", value))
      );

    new Setting(containerEl)
      .setName("Show gridlines")
      .setDesc("Draw light gridlines in cells that have no border of their own.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.showGridlines !== false)
          .onChange((value) => this.plugin.updateSetting("showGridlines", value))
      );

    new Setting(containerEl)
      .setName("Show row and column headers")
      .setDesc("Show the A B C column letters and the row numbers around a sheet.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.showHeaders !== false)
          .onChange((value) => this.plugin.updateSetting("showHeaders", value))
      );

    new Setting(containerEl)
      .setName("Filter chips limit")
      .setDesc("Show chips above the grid for autofilter columns with at most this many values. The full dropdown stays available on every filter column.")
      .addSlider((slider) =>
        slider
          .setLimits(4, 40, 2)
          .setValue(Number(this.plugin.settings.chipLimit) || 10)
          .setDynamicTooltip()
          .onChange((value) => this.plugin.updateSetting("chipLimit", value))
      );

    new Setting(containerEl)
      .setName("Minimum column width")
      .setDesc("Widen every column to at least this many pixels, so long values are readable without dragging. Off at 0.")
      .addText((text) =>
        text
          .setPlaceholder("0")
          .setValue(String(this.plugin.settings.minColumnWidth || 0))
          .onChange((value) => {
            const n = parseInt(value, 10);
            this.plugin.updateSetting("minColumnWidth", isNaN(n) || n < 0 ? 0 : n);
          })
      );

    new Setting(containerEl)
      .setName("Minimum row height")
      .setDesc("Make every row at least this many pixels tall. Off at 0.")
      .addText((text) =>
        text
          .setPlaceholder("0")
          .setValue(String(this.plugin.settings.minRowHeight || 0))
          .onChange((value) => {
            const n = parseInt(value, 10);
            this.plugin.updateSetting("minRowHeight", isNaN(n) || n < 0 ? 0 : n);
          })
      );

    new Setting(containerEl)
      .setName("Row render limit")
      .setDesc("Maximum rows to draw at once, to keep very large sheets responsive. Set 0 for no limit.")
      .addText((text) =>
        text
          .setPlaceholder("2000")
          .setValue(String(this.plugin.settings.maxRows))
          .onChange((value) => {
            const n = parseInt(value, 10);
            this.plugin.updateSetting("maxRows", isNaN(n) || n < 0 ? 0 : n);
          })
      );

    containerEl.createEl("h3", { text: "Documents" });

    new Setting(containerEl)
      .setName("Page background")
      .setDesc("White draws a white page, like Word. Theme tints the page with the Obsidian background.")
      .addDropdown((dropdown) =>
        dropdown
          .addOption("white", "Always white (recommended)")
          .addOption("theme", "Follow the Obsidian theme")
          .setValue(this.plugin.settings.pageBackground)
          .onChange((value) => this.plugin.updateSetting("pageBackground", value))
      );

    new Setting(containerEl)
      .setName("Show page numbers")
      .setDesc("Replace PAGE and NUMPAGES fields with the number they stood for when the document was last saved.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.showPageNumbers !== false)
          .onChange((value) => this.plugin.updateSetting("showPageNumbers", value))
      );

    containerEl.createEl("h3", { text: "Presentations" });

    new Setting(containerEl)
      .setName("Slide fit")
      .setDesc("Contain fits the whole slide in the pane. Actual size keeps the slide at its saved size, which needs scrolling.")
      .addDropdown((dropdown) =>
        dropdown
          .addOption("contain", "Fit the whole slide")
          .addOption("none", "Actual size")
          .setValue(this.plugin.settings.slideFit)
          .onChange((value) => this.plugin.updateSetting("slideFit", value))
      );

    new Setting(containerEl)
      .setName("Show speaker notes")
      .setDesc("Show the notes pane under the slide when the deck carries notes.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.showSpeakerNotes !== false)
          .onChange((value) => this.plugin.updateSetting("showSpeakerNotes", value))
      );

    containerEl.createEl("h3", { text: "View" });

    new Setting(containerEl)
      .setName("Show the counts inside the view")
      .setDesc("Put the word, character and page counts at the bottom left of the view. Off, they live in the app's status bar and cost the document no height.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.statusInView === true)
          .onChange((value) => this.plugin.updateSetting("statusInView", value))
      );

    new Setting(containerEl)
      .setName("Pinch to zoom")
      .setDesc("A touchpad pinch over an open document, sheet or slide steps the zoom in and out.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.pinchZoom !== false)
          .onChange((value) => this.plugin.updateSetting("pinchZoom", value))
      );

    new Setting(containerEl)
      .setName("Invert pinch zoom")
      .setDesc("Swap the pinch direction for a reader who expects the other one.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.pinchInvert === true)
          .onChange((value) => this.plugin.updateSetting("pinchInvert", value))
      );

    containerEl.createEl("p", {
      cls: "setting-item-description",
      text: "Everything opens read only. The viewer never writes to an Office file; it only remembers which sheet, filters and slide you were on.",
    });
  }
}

module.exports = OfficeViewerPlugin;
