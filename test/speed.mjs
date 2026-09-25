// Speed pass: opens the largest real files through the built bundle's view and
// reports parse and mount time, so regressions in the hot paths are visible.
import { createRequire } from "node:module";
import { readFileSync, statSync } from "node:fs";
import { basename, join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { setupDom } from "./harness.mjs";
import { testEnv, testPath, skipNotice } from "./env.mjs";

setupDom();
const require = createRequire(import.meta.url);
const root = join(dirname(fileURLToPath(import.meta.url)), "..");
let files = process.argv.slice(2).filter((a) => !a.startsWith("--"));
if (!files.length) {
  // With nothing named, time the configured sample document.
  const single = testPath("OV_TEST_PPTX") || testPath("OV_TEST_DOCX");
  if (!single) {
    skipNotice("OV_TEST_PPTX", "speed run");
    process.exit(0);
  }
  files = [single];
}
void testEnv;

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
for (const file of files) {
  currentPath = file;
  const sizeMb = statSync(file).size / 1024 / 1024;
  const view = viewFactory(new WorkspaceLeaf());
  view.app = app;
  const started = performance.now();
  await view.onOpen();
  await view.onLoadFile(new TFile(basename(file)));
  const total = performance.now() - started;
  const error = view.contentEl.querySelector(".ov-message-error");
  rows.push({ name: basename(file).slice(0, 52), sizeMb, total, error: Boolean(error) });
  view.onClose();
}
rows.sort((a, b) => b.total - a.total);
console.log("file".padEnd(54), "size".padStart(8), "open".padStart(10));
for (const row of rows) {
  console.log(row.name.padEnd(54), (row.sizeMb.toFixed(1) + " MB").padStart(8), (Math.round(row.total) + " ms").padStart(10), row.error ? "ERROR" : "");
}
const total = rows.reduce((a, b) => a + b.total, 0);
console.log("");
console.log(rows.length + " files, total " + Math.round(total) + " ms, average " + Math.round(total / rows.length) + " ms");
