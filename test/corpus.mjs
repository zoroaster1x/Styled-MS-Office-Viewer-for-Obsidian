/*
 * Styled MS Office Viewer, an Obsidian plugin that renders office documents
 * (xlsx, docx, pptx and their relatives) with their real styling, read only.
 *
 * Copyright (C) 2026 Zoroaster1x
 *
 * This program is free software: you can redistribute it and/or modify it under
 * the terms of the GNU General Public License as published by the Free Software
 * Foundation, either version 3 of the License, or (at your option) any later
 * version.
 *
 * This program is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
 * FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
 * details.
 *
 * You should have received a copy of the GNU General Public License along with
 * this program. If not, see <https://www.gnu.org/licenses/>.
 */

// Sweeps a folder of real documents through the built bundle and reports what
// failed and what was slow. Nothing is written anywhere but the report.
//
//   bun test/corpus.mjs ~/Documents /tmp/corpus-report.json
//
// Every file the plugin claims is opened the way a click opens it. A file that
// shows the error box is a failure with its message; a file that opens is
// timed. The report is sorted by open time, so the slow outliers come first.

import { createRequire } from "node:module";
import { readFileSync, readdirSync, statSync, writeFileSync } from "node:fs";
import { basename, join } from "node:path";
import { setupDom } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const root = join(import.meta.dir, "..");

const folder = process.argv[2];
const reportPath = process.argv[3] || join("/tmp", "corpus-report.json");
if (!folder) {
  console.log("usage: bun test/corpus.mjs <folder> [report.json]");
  process.exit(2);
}

const CLAIMED = new Set([
  "xlsx", "xlsm", "xltx", "xltm", "csv", "tsv", "ods",
  "docx", "docm", "dotx", "dotm", "odt", "rtf", "doc",
  "pptx", "pptm", "ppsx", "ppsm", "potx", "potm", "odp",
]);
const KNOWN_UNSUPPORTED = new Set(["xls", "ppt"]);

function walk(dir, out) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      if (entry.name.startsWith(".")) continue;
      walk(full, out);
      continue;
    }
    const ext = entry.name.slice(entry.name.lastIndexOf(".") + 1).toLowerCase();
    if (CLAIMED.has(ext) || KNOWN_UNSUPPORTED.has(ext)) out.push({ path: full, ext });
  }
}

const files = [];
walk(folder, files);
console.log("found", files.length, "office files in", folder);

const PluginClass = require(join(root, "main.js"));
const { WorkspaceLeaf, TFile } = require("obsidian");
let currentPath = null;
const app = {
  workspace: { getLeavesOfType: () => [], getActiveViewOfType: () => null },
  vault: {
    on: () => ({}),
    adapter: { getFullPath: (p) => p },
    readBinary: async () => {
      const bytes = readFileSync(currentPath);
      return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    },
  },
};
const plugin = new PluginClass();
plugin.app = app;
plugin.settings = { fileState: {} };
plugin.saveFileState = async () => {};
plugin.updateSetting = async () => {};
await plugin.onload();
const viewFactory = plugin.registered.views[0].factory;

const rows = [];
for (const [i, file] of files.entries()) {
  if (KNOWN_UNSUPPORTED.has(file.ext)) continue;
  currentPath = file.path;
  const size = statSync(file.path).size;
  const view = viewFactory(new WorkspaceLeaf());
  view.app = app;
  const started = performance.now();
  let error = null;
  let failure = null;
  try {
    await view.onOpen();
    await view.onLoadFile(new TFile(basename(file.path)));
    const box = view.contentEl.querySelector(".ov-message-error");
    if (box) error = (box.querySelector("p") || box).textContent.trim().slice(0, 200);
  } catch (err) {
    failure = (err && err.message) || String(err);
  }
  const ms = Math.round(performance.now() - started);
  try {
    view.onClose();
  } catch (err) {
    // A close that throws is worth knowing about too.
    failure = failure || "close: " + ((err && err.message) || String(err));
  }
  rows.push({ path: file.path, ext: file.ext, bytes: size, ms, error, failure });
  // One line per file, so a document that hangs names itself instead of
  // leaving the sweep looking stuck between two progress marks.
  console.log(
    String(i + 1).padStart(4) + "/" + files.length +
    " " + String(ms).padStart(6) + " ms " + (size / 1048576).toFixed(1).padStart(6) + " MB " +
    (error || failure ? "FAIL " + (error || failure).slice(0, 60) : "ok   ") +
    " " + file.path.replace(folder, ""),
  );
}

const failed = rows.filter((r) => r.error || r.failure);
const ok = rows.filter((r) => !r.error && !r.failure);
ok.sort((a, b) => b.ms - a.ms);
failed.sort((a, b) => b.bytes - a.bytes);

console.log("");
console.log("opened", ok.length, "failed", failed.length);
if (failed.length) {
  console.log("");
  console.log("failures, biggest first:");
  for (const row of failed.slice(0, 40)) {
    console.log("  " + (row.bytes / 1048576).toFixed(1) + " MB  " + row.ext + "  " + (row.error || row.failure));
    console.log("      " + row.path);
  }
}
console.log("");
console.log("slowest that did open:");
for (const row of ok.slice(0, 15)) {
  console.log("  " + String(row.ms).padStart(6) + " ms  " + (row.bytes / 1048576).toFixed(1) + " MB  " + row.path);
}

writeFileSync(reportPath, JSON.stringify({ folder, total: rows.length, failed, rows }, null, 1));
console.log("");
console.log("report written to", reportPath);
