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

// Renders a document with the real controllers into a standalone HTML file, so
// the result can be opened in any browser and compared against a reference
// render side by side.
//
//   bun test/render-html.mjs <file> [--slide N] [--sheet N] [--out out.html]
//
// Pictures come out as data URLs, so the file is self contained. The app's
// stylesheet is inlined and a small theme shim stands in for Obsidian's light
// theme, which is the only thing the renderers need from the host.

import { readFileSync, writeFileSync } from "node:fs";
import { basename } from "node:path";
import { createRequire } from "node:module";
import { setupDom, createContainer } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { detectFormat } = require("../src/app/format-router.js");
const { DocumentController } = require("../src/app/document-controller.js");
const { PresentationController } = require("../src/app/presentation-controller.js");
const { SpreadsheetController } = require("../src/app/spreadsheet-controller.js");
const { ensureEmbeddedFonts } = require("../src/pptx/fonts.js");

const args = process.argv.slice(2);
const file = args.find((a) => !a.startsWith("--"));
function valueOf(flag, fallback) {
  const i = args.indexOf(flag);
  return i === -1 ? fallback : args[i + 1];
}
const out = valueOf("--out", "/tmp/opencode/out/render.html");
const slide = Number(valueOf("--slide", 1));
const sheet = Number(valueOf("--sheet", 1));
const bare = args.includes("--bare");

if (!file) {
  console.error("usage: bun test/render-html.mjs <file> [--slide N] [--sheet N] [--out out.html]");
  process.exit(2);
}

const bytes = new Uint8Array(readFileSync(file));
const name = basename(file);
const format = detectFormat(name, bytes);

// Every blob URL the renderers make is remembered so it can be rewritten as a
// data URL once the document has been laid out.
const blobs = new Map();
let blobCount = 0;
const realCreate = URL.createObjectURL ? URL.createObjectURL.bind(URL) : null;
URL.createObjectURL = (blob) => {
  void realCreate;
  const url = "blob:office-viewer/" + blobCount++;
  blobs.set(url, blob);
  return url;
};

const settings = {
  zoom: 100,
  slideZoom: 100,
  sheetBackground: "white",
  showGridlines: true,
  showHeaders: true,
  showSpeakerNotes: true,
  slideFit: "contain",
  minColumnWidth: 0,
  minRowHeight: 0,
  maxRows: 0,
  fileState: {},
};
const plugin = {
  settings,
  savePluginData() {},
  async savePluginData2() {},
  saveFileState() {},
  updateSetting() {},
  registerEvent() {},
};

const host = createContainer();
const tabsEl = createContainer();
const filterBarEl = createContainer();
const callbacks = {};
const saved = format.kind === "presentation" ? { slide: slide - 1 } : format.kind === "spreadsheet" ? { sheet: sheet - 1 } : {};

const controller = format.kind === "spreadsheet"
  ? new SpreadsheetController({ host, tabsEl, filterBarEl, plugin, format: format.ext, callbacks })
  : format.kind === "presentation"
    ? new PresentationController({ host, plugin, format: format.ext, callbacks })
    : new DocumentController({ host, plugin, format: format.ext, callbacks });

controller.filePath = file;
controller.savedState = saved;
controller.load(bytes, name);
await ensureEmbeddedFonts(controller.model);
if (typeof controller.restoreState === "function") controller.restoreState(saved);
controller.mount();
if (format.kind === "presentation" && controller.goToSlide) controller.goToSlide(slide - 1);

async function dataUrlFor(blob) {
  const buffer = Buffer.from(await blob.arrayBuffer());
  return "data:" + (blob.type || "application/octet-stream") + ";base64," + buffer.toString("base64");
}

// --bare drops the rail, the notes and the fit transform so the page can be
// screenshotted at its natural size and diffed against a reference render.
if (bare) {
  for (const el of Array.from(host.querySelectorAll(".ov-pptx-rail, .ov-pptx-notes, .ov-pptx-status"))) el.remove();
  const root = host.querySelector(".ov-pptx");
  if (root) root.style.setProperty("--ov-pptx-scale", "1");
  const page = host.querySelector(".ov-pptx-page");
  if (page) {
    page.style.transform = "scale(1)";
    const frame = host.querySelector(".ov-pptx-frame");
    if (frame) {
      frame.style.width = page.style.width;
      frame.style.height = page.style.height;
    }
  }
}

let html = host.innerHTML;
// One pass over every blob URL. Replacing them one by one corrupts the URLs
// that share a prefix ("blob:x/1" also matches inside "blob:x/12").
const dataUrls = new Map();
for (const [url, blob] of blobs) dataUrls.set(url, await dataUrlFor(blob));
html = html.replace(/blob:[A-Za-z0-9._\-/]+/g, (match) => dataUrls.get(match) || match);

const theme = `
:root {
  --background-primary: #ffffff;
  --background-primary-alt: #ffffff;
  --background-secondary: #f2f3f5;
  --background-secondary-alt: #f2f3f5;
  --background-modifier-border: #dddddd;
  --background-modifier-hover: rgba(0, 0, 0, 0.05);
  --background-modifier-active-hover: rgba(0, 0, 0, 0.08);
  --text-normal: #2e3338;
  --text-muted: #888888;
  --text-faint: #aaaaaa;
  --text-accent: #705dcf;
  --text-on-accent: #ffffff;
  --interactive-accent: #7f6df2;
  --interactive-accent-hover: #8875ff;
  --interactive-normal: #f2f3f5;
  --interactive-hover: #e9e9e9;
  --color-accent: #7f6df2;
  --font-interface: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  --font-text: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  --font-monospace: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace;
}
html, body { margin: 0; padding: 0; background: #f6f6f6; }
body.ov-preview > .ov-host { padding: 12px; }
body.ov-bare { background: #ffffff; }
body.ov-bare > .ov-host { padding: 0; }
body.ov-bare .ov-pptx, body.ov-bare .ov-pptx-body, body.ov-bare .ov-pptx-stagewrap,
body.ov-bare .ov-pptx-stage, body.ov-bare .ov-docx-scroll { display: block; margin: 0; padding: 0; overflow: visible; }
body.ov-bare .ov-pptx-page, body.ov-bare .ov-docx-page { box-shadow: none; margin: 0; }
`;

const page = `<!doctype html>
<html><head><meta charset="utf-8"><title>${name}</title>
<style>
${readFileSync(new URL("../styles.css", import.meta.url), "utf8")}
${theme}
</style>
</head>
<body class="ov-preview${bare ? " ov-bare" : ""}"><div class="ov-host">${html}</div></body></html>`;

writeFileSync(out, page);
console.log("wrote", out, "bytes", page.length, "kind", format.kind, "blobs", blobs.size);
