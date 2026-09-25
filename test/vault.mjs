/*
 * Opens real files through the built bundle's view, with a stubbed vault that
 * serves bytes from disk. This is the closest a CLI test gets to clicking a
 * file in the app: the same view code, the same controllers, the same
 * renderers, only the file IO is simulated.
 *
 *   bun test/vault.mjs <file> [<file> ...]
 */

import { createRequire } from "node:module";
import { readFileSync, existsSync } from "node:fs";
import { basename, join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { setupDom } from "./harness.mjs";
import { testPath, skipNotice } from "./env.mjs";

setupDom();
const require = createRequire(import.meta.url);
const root = join(dirname(fileURLToPath(import.meta.url)), "..");

let files = process.argv.slice(2).filter((a) => !a.startsWith("--"));
if (!files.length) {
  const single = testPath("OV_TEST_XLSX") || testPath("OV_TEST_DOCX") || testPath("OV_TEST_PPTX");
  if (!single) {
    skipNotice("OV_TEST_XLSX", "vault run");
    process.exit(0);
  }
  files = [single];
}

const PluginClass = require(join(root, "main.js"));
const { WorkspaceLeaf, TFile } = require("obsidian");

const opened = [];
const app = {
  workspace: { getLeavesOfType: () => [], getActiveViewOfType: () => null },
  vault: {
    on: () => ({}),
    adapter: { getFullPath: (p) => p },
    readBinary: async (file) => {
      const bytes = readFileSync(opened[opened.length - 1].diskPath);
      return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    },
  },
};

const plugin = new PluginClass();
plugin.app = app;
plugin.settings = { fileState: {} };
plugin.saveFileState = async (path, state) => {
  plugin.settings.fileState[path] = state;
};
plugin.updateSetting = async (key, value) => {
  plugin.settings[key] = value;
};
await plugin.onload();

const viewFactory = plugin.registered.views[0].factory;
let pass = 0;
let fail = 0;

for (const diskPath of files) {
  if (!existsSync(diskPath)) {
    console.log("FAIL missing file", diskPath);
    fail++;
    continue;
  }
  const file = new TFile(basename(diskPath));
  file.path = basename(diskPath);
  opened.push({ diskPath, file });
  const view = viewFactory(new WorkspaceLeaf());
  view.app = app;
  const started = performance.now();
  try {
    await view.onOpen();
    // The tab must carry the name, and the pane a loading line, before the
    // bytes are read: the read plus the synchronous parse is what takes time.
    const opening = view.onLoadFile(file);
    const earlyTitle = view.getDisplayText() === file.name;
    const earlyLoading = !!view.hostEl.querySelector(".ov-loading");
    await opening;
    const ms = performance.now() - started;
    const loadingAfter = !!view.hostEl.querySelector(".ov-loading");
    const text = (view.contentEl.textContent || "").replace(/\s+/g, " ");
    const errorBox = view.contentEl.querySelector(".ov-message-error");
    const unsupported = view.contentEl.querySelector(".ov-message-warning");
    if (errorBox) {
      console.log("FAIL", basename(diskPath), "->", errorBox.textContent.slice(0, 120));
      fail++;
    } else if (unsupported) {
      console.log("WARN", basename(diskPath), "->", unsupported.textContent.slice(0, 90));
      fail++;
    } else {
      const status = view.statusEl ? view.statusEl.textContent : "";
      console.log("ok  ", basename(diskPath).slice(0, 60).padEnd(60), Math.round(ms) + " ms", "|", text.length + " chars |", status.slice(0, 70));
      // The counts live in the app status bar by default; the option puts the
      // line inside the view at the bottom left instead.
      plugin.settings.statusInView = true;
      view.applySettings();
      const statusInView = view.contentEl.classList.contains("ov-status-inline")
        && Boolean(view.statusEl && view.statusEl.textContent.trim());
      plugin.settings.statusInView = false;
      view.applySettings();
      // Toolbar: the finder chevrons sit left of the search box and only show
      // for a search with matches; the zoom percentage sits with the magnifiers.
      const zoomInBtn = view.toolsEl.querySelector('button[data-icon="zoom-in"]');
      const navLeft = view.searchNavEl && view.searchNavEl.nextElementSibling === view.searchInput;
      const labelWithZoom = Boolean(zoomInBtn && zoomInBtn.nextElementSibling === view.zoomLabel);
      const navHidden = view.searchNavEl.classList.contains("is-hidden");
      view.runSearch("e");
      const navShown = !view.searchNavEl.classList.contains("is-hidden");
      view.runSearch("");
      const navGone = view.searchNavEl.classList.contains("is-hidden");
      // Pinch: a ctrl+wheel steps the zoom, and the invert setting flips it.
      let zoomDelta = 0;
      const realSetZoom = view.controller.setZoom.bind(view.controller);
      view.controller.setZoom = (value) => { zoomDelta = value - view.controller.currentZoom(); };
      plugin.settings.pinchZoom = true;
      plugin.settings.pinchInvert = false;
      view.pinchAccum = 0;
      view.onWheelZoom({ ctrlKey: true, deltaY: -40, preventDefault() {} });
      const pinchIn = zoomDelta > 0;
      view.pinchAccum = 0;
      plugin.settings.pinchInvert = true;
      view.onWheelZoom({ ctrlKey: true, deltaY: -40, preventDefault() {} });
      const pinchOut = zoomDelta < 0;
      plugin.settings.pinchInvert = false;
      view.controller.setZoom = realSetZoom;
      const chromeOk = navLeft && labelWithZoom && navHidden && navShown && navGone && pinchIn && pinchOut;
      if (earlyTitle && earlyLoading && !loadingAfter && statusInView && chromeOk) {
        pass++;
      } else {
        console.log("FAIL", basename(diskPath), "-> early title " + earlyTitle + ", loading line " + earlyLoading + ", leftover " + loadingAfter + ", status in view " + statusInView + ", toolbar " + chromeOk + " (" + navLeft + "/" + labelWithZoom + "/" + navHidden + "/" + navShown + "/" + navGone + "/" + pinchIn + "/" + pinchOut + ")");
        fail++;
      }
      // A deck also has to survive slideshow mode: the class on, the arrow keys
      // moving slides, an input field left alone, the browser's own fullscreen
      // exit (Escape) taking the mode down, and a pane resize refitting the slide.
      if (view.format && view.format.kind === "presentation" && view.controller) {
        const moves = { next: 0, previous: 0, first: 0 };
        const controller = view.controller;
        const original = {
          next: controller.nextSlide,
          previous: controller.previousSlide,
          goTo: controller.goToSlide,
          applySettings: controller.applySettings ? controller.applySettings.bind(controller) : null,
        };
        let refits = 0;
        controller.nextSlide = () => { moves.next++; };
        controller.previousSlide = () => { moves.previous++; };
        controller.goToSlide = () => { moves.first++; };
        if (original.applySettings) controller.applySettings = () => { refits++; original.applySettings(); };
        view.toggleSlideshow();
        const entered = view.slideshow === true && view.contentEl.classList.contains("ov-slideshow");
        // Escape while an element is fullscreen is the browser's key: it leaves
        // fullscreen and no keydown reaches the view. The change event is the
        // only signal, and it must leave the mode and refit the slide.
        document.dispatchEvent(new globalThis.Event("fullscreenchange"));
        const escaped = view.slideshow === false && !view.contentEl.classList.contains("ov-slideshow") && refits > 0;
        view.toggleSlideshow();
        const key = (k, target) => view.onSlideshowKey({ key: k, target, preventDefault() {} });
        key("ArrowRight");
        key(" ");
        key("ArrowLeft");
        key("Home");
        key("ArrowRight", { tagName: "INPUT" });
        key("Escape");
        const left = view.slideshow === false && !view.contentEl.classList.contains("ov-slideshow");
        // A pane resize must recompute the fit without a settings change.
        const stage = view.contentEl.querySelector(".ov-pptx-stage");
        const frame = view.contentEl.querySelector(".ov-pptx-frame");
        let resized = false;
        if (stage && frame && view.controller.model && typeof globalThis.ResizeObserver !== "function") {
          const slideWidth = view.controller.model.widthPx;
          const slideHeight = view.controller.model.heightPx;
          stage.getBoundingClientRect = () => ({ width: 1000, height: 600, top: 0, bottom: 600, left: 0, right: 1000 });
          window.dispatchEvent(new globalThis.Event("resize"));
          await new Promise((resolve) => setTimeout(resolve, 40));
          const expected = Math.max(0.05, Math.min((1000 - 24) / slideWidth, (600 - 24) / slideHeight));
          resized = Math.abs(parseFloat(frame.style.width) - slideWidth * expected) < 1;
        }
        const good = entered && escaped && left && resized && moves.next === 2 && moves.previous === 1 && moves.first === 1;
        console.log(good ? "ok  " : "FAIL", "slideshow mode                 next " + moves.next + ", previous " + moves.previous + ", home " + moves.first + ", input ignored " + (moves.next === 2) + ", fullscreen exit " + escaped + ", refit " + resized);
        if (good) pass++;
        else fail++;
        controller.nextSlide = original.next;
        controller.previousSlide = original.previous;
        controller.goToSlide = original.goTo;
        if (original.applySettings) controller.applySettings = original.applySettings;
      }
    }
    view.onClose();
  } catch (err) {
    console.log("ERR ", basename(diskPath), err && err.message ? err.message : err);
    fail++;
  }
}

console.log("");
console.log("vault:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
