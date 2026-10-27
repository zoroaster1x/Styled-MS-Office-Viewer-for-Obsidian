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

// OpenDocument files that a real office suite wrote, opened through the built
// bundle. The synthetic packages in test/odf.mjs prove the routing; this proves
// the readers against LibreOffice's own output, which is what a user sends.
//
// LibreOffice has to be installed. The source documents come from the same
// .testenv entries the other real document suites use, or from the command
// line. Anything missing is skipped with a notice rather than failed, so CI
// (which has no LibreOffice) still runs the rest of the plan.
//
//   bun test/odf-live.mjs
//   bun test/odf-live.mjs path/to/deck.pptx path/to/book.xlsx

import { createRequire } from "node:module";
import { readFileSync, mkdtempSync, existsSync, rmSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { basename, extname, join } from "node:path";
import { tmpdir } from "node:os";
import { setupDom } from "./harness.mjs";
import { testPath, skipNotice } from "./env.mjs";

setupDom();
const require = createRequire(import.meta.url);
const root = join(import.meta.dir, "..");
const dir = mkdtempSync(join(tmpdir(), "ov-odf-"));

let pass = 0;
let fail = 0;
function check(name, condition, detail) {
  if (condition) {
    pass++;
    console.log("ok   " + name);
  } else {
    fail++;
    console.log("FAIL " + name + (detail !== undefined ? "  -> " + detail : ""));
  }
}

const soffice = spawnSync("sh", ["-c", "command -v soffice || command -v libreoffice"], { encoding: "utf8" });
const binary = (soffice.stdout || "").trim();
if (!binary) {
  skipNotice("LibreOffice", "OpenDocument round trip");
  process.exit(0);
}

// One source per format, so each reader is exercised. .ods comes from a
// workbook, .odp from a deck, .odt from a text document.
const sources = process.argv.length > 2
  ? process.argv.slice(2).filter((a) => !a.startsWith("--"))
  : [testPath("OV_TEST_XLSX"), testPath("OV_TEST_DOCX"), testPath("OV_TEST_PPTX")].filter(Boolean);
const present = sources.filter((file) => existsSync(file));
if (!present.length) {
  skipNotice("OV_TEST_XLSX, OV_TEST_DOCX or OV_TEST_PPTX", "OpenDocument round trip");
  process.exit(0);
}

// ---------- convert with LibreOffice ----------

function convert(file) {
  const ext = extname(file).slice(1).toLowerCase();
  const target = ext === "xlsx" ? "ods" : ext === "pptx" ? "odp" : "odt";
  const run = spawnSync(binary, ["--headless", "--convert-to", target, "--outdir", dir, file], { encoding: "utf8", timeout: 300000 });
  const out = join(dir, basename(file, extname(file)) + "." + target);
  return { target, out, ok: existsSync(out), log: (run.stderr || "").trim() };
}

const converted = [];
for (const file of present) {
  const result = convert(file);
  check("LibreOffice wrote " + basename(result.out), result.ok, result.log || "no output file");
  if (result.ok) converted.push(result);
}
if (!converted.length) {
  console.log("");
  console.log("odf-live: 0 pass, " + fail + " fail");
  process.exit(fail ? 1 : 0);
}

// ---------- open each one through the bundle ----------

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

for (const { out } of converted) {
  currentPath = out;
  const view = viewFactory(new WorkspaceLeaf());
  view.app = app;
  const started = performance.now();
  await view.onOpen();
  await view.onLoadFile(new TFile(basename(out)));
  const total = Math.round(performance.now() - started);
  const error = view.contentEl.querySelector(".ov-message-error");
  check(basename(out) + " opens", !error, error ? error.textContent.trim().slice(0, 120) : "");
  // A rendered document puts something in the pane beyond the error box.
  const drawn = view.contentEl.querySelectorAll("*").length;
  check(basename(out) + " draws content", drawn > 20, drawn + " elements");
  console.log("     " + basename(out) + " opened in " + total + " ms");
  view.onClose();
}

rmSync(dir, { recursive: true, force: true });
console.log("");
console.log("odf-live: " + pass + " pass, " + fail + " fail");
if (fail) process.exit(1);
