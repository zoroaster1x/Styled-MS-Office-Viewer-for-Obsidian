/*
 * A stub of the Obsidian API for tests. Only the surface the plugin touches is
 * implemented; every method records its call so tests can assert on it.
 */

class Events {
  on() { return this; }
  off() { return this; }
  trigger() { return this; }
}

class Component extends Events {
  load() {}
  onload() {}
  unload() {}
  register() {}
  registerEvent() {}
  registerDomEvent() {}
  registerInterval() {}
}

class WorkspaceLeaf extends Component {
  constructor() {
    super();
    this.view = null;
  }
  setViewState() { return Promise.resolve(); }
  detach() { return Promise.resolve(); }
}

class TFile {
  constructor(path) {
    this.path = path;
    this.name = path.split("/").pop();
    this.basename = this.name.replace(/\.[^.]+$/, "");
    this.extension = (this.name.split(".").pop() || "").toLowerCase();
  }
}

class Plugin extends Component {
  constructor(app, manifest) {
    super();
    this.app = app || appStub;
    this.manifest = manifest || { id: "styled-ms-office-viewer", version: "2.0.0" };
    this.registered = { views: [], extensions: null, commands: [], settingTabs: [] };
  }
  registerView(type, factory) {
    this.registered.views.push({ type, factory });
  }
  registerExtensions(extensions, viewType) {
    this.registered.extensions = { extensions, viewType };
  }
  addSettingTab(tab) {
    this.registered.settingTabs.push(tab);
  }
  addCommand(command) {
    this.registered.commands.push(command);
  }
  addRibbonIcon() {}
  async loadData() { return {}; }
  async saveData() {}
}

class PluginSettingTab {
  constructor(app, plugin) {
    this.app = app;
    this.plugin = plugin;
    this.containerEl = globalThis.document.createElement("div");
  }
  display() {}
  hide() {}
}

class Setting {
  constructor(containerEl) {
    this.containerEl = containerEl;
  }
  setName() { return this; }
  setDesc() { return this; }
  setHeading() { return this; }
  addText(cb) { cb(new TextComponent()); return this; }
  addToggle(cb) { cb(new ToggleComponent()); return this; }
  addDropdown(cb) { cb(new DropdownComponent()); return this; }
  addSlider(cb) { cb(new SliderComponent()); return this; }
  addButton(cb) { cb(new ButtonComponent()); return this; }
}

class BaseComponent {
  setValue() { return this; }
  setPlaceholder() { return this; }
  onChange() { return this; }
  setLimits() { return this; }
  setDynamicTooltip() { return this; }
  addOption() { return this; }
  setDisabled() { return this; }
  setCta() { return this; }
  setButtonText() { return this; }
  setTooltip() { return this; }
}
class TextComponent extends BaseComponent {}
class ToggleComponent extends BaseComponent {}
class DropdownComponent extends BaseComponent {}
class SliderComponent extends BaseComponent {}
class ButtonComponent extends BaseComponent {}

class ItemView extends Component {
  constructor(leaf) {
    super();
    this.leaf = leaf;
    // Obsidian's View carries the app; anything that loads files expects it.
    this.app = appStub;
    this.containerEl = globalThis.document.createElement("div");
    this.contentEl = globalThis.document.createElement("div");
    this.containerEl.appendChild(this.contentEl);
  }
  getViewType() { return "stub"; }
  getDisplayText() { return "stub"; }
  getIcon() { return "file"; }
}

class FileView extends ItemView {}

class Modal extends Component {
  constructor(app) {
    super();
    this.app = app;
    this.contentEl = globalThis.document.createElement("div");
  }
  open() {}
  close() {}
}

class Menu {
  constructor() {
    this.items = [];
  }
  addItem(cb) {
    const item = {
      setTitle() { return item; },
      setIcon() { return item; },
      onClick(handler) { item.handler = handler; return item; },
    };
    this.items.push(item);
    cb(item);
    return this;
  }
  showAtMouseEvent() {}
  showAtPosition() {}
}

class Notice {
  constructor(message) {
    this.message = message;
    Notice.last = this;
  }
}

function setIcon(el, icon) {
  if (el) el.setAttribute("data-icon", icon);
}

const appStub = {
  workspace: {
    getLeavesOfType: () => [],
    getActiveViewOfType: () => null,
    on: () => ({}),
  },
  vault: {
    adapter: { getFullPath: (path) => path },
    readBinary: async () => new ArrayBuffer(0),
    on: () => ({}),
  },
  metadataCache: { getFileCache: () => null },
};

module.exports = {
  Plugin,
  PluginSettingTab,
  Setting,
  ItemView,
  FileView,
  Modal,
  Menu,
  Notice,
  TFile,
  WorkspaceLeaf,
  Component,
  Events,
  setIcon,
  App: class {},
  MarkdownView: class {},
  normalizePath: (p) => p,
  requestUrl: async () => ({}),
};
