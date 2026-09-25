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

// Loads the built bundle under the Obsidian stub and checks that the plugin
// wires itself up: view, extensions, commands and settings tab. This catches
// a broken bundle, a bad import path or a typo in onload before Obsidian ever
// sees the plugin.
//
//   bun test/smoke.mjs

import { createRequire } from "node:module";
import { existsSync, readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { setupDom } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const root = join(dirname(fileURLToPath(import.meta.url)), "..");

let pass = 0;
let fail = 0;
function check(name, condition, detail) {
  if (condition) {
    pass++;
    console.log("ok   " + name);
  } else {
    fail++;
    console.log("FAIL " + name + (detail ? "  -> " + detail : ""));
  }
}

const bundlePath = join(root, "main.js");
if (!existsSync(bundlePath)) {
  console.log("FAIL main.js is missing; run: bun esbuild.config.mjs production");
  process.exit(1);
}

const PluginClass = require(bundlePath);
const plugin = new PluginClass();
plugin.app = require("obsidian").default ? require("obsidian").default : plugin.app;
if (!plugin.app || !plugin.app.workspace) {
  plugin.app = {
    workspace: { getLeavesOfType: () => [], getActiveViewOfType: () => null },
    vault: { readBinary: async () => new ArrayBuffer(0), adapter: { getFullPath: (p) => p } },
  };
}
await plugin.onload();

const registered = plugin.registered || { views: [], extensions: null, commands: [], settingTabs: [] };
check("view registered", registered.views.length === 1, JSON.stringify(registered.views.map((v) => v.type)));
check("view type is office-viewer", registered.views[0] && registered.views[0].type === "office-viewer", registered.views[0] && registered.views[0].type);

// The manifest is what Obsidian installs from. The id has to be lowercase and
// hyphenated, the folder is named after it, and versions.json has to carry the
// version the manifest declares, or Obsidian shows the plugin as broken.
const manifest = JSON.parse(readFileSync(join(root, "manifest.json"), "utf8"));
const versions = JSON.parse(readFileSync(join(root, "versions.json"), "utf8"));
check("manifest id is valid for Obsidian", /^[a-z0-9-]+$/.test(manifest.id), manifest.id);
check("manifest id matches the plugin", manifest.id === "styled-ms-office-viewer", manifest.id);
check("manifest name is set", manifest.name === "Styled MS Office Viewer", manifest.name);
check("manifest version is semver", /^\d+\.\d+\.\d+$/.test(manifest.version), manifest.version);
check("versions.json lists the manifest version", Boolean(versions[manifest.version]), Object.keys(versions).join(","));
check("manifest description names the families", /xlsx/.test(manifest.description) && /docx/.test(manifest.description) && /pptx/.test(manifest.description), manifest.description);

const extensions = registered.extensions ? registered.extensions.extensions : [];
const expected = ["xlsx", "xlsm", "csv", "tsv", "ods", "docx", "docm", "odt", "rtf", "pptx", "pptm", "odp"];
for (const ext of expected) {
  check("claims ." + ext, extensions.indexOf(ext) !== -1);
}
check("legacy .doc is claimed as text", extensions.indexOf("doc") !== -1);
check("legacy .xls is not claimed", extensions.indexOf("xls") === -1);
check("legacy .ppt is not claimed", extensions.indexOf("ppt") === -1);
const fullSet = [
  "xlsx", "xlsm", "xltx", "xltm", "csv", "tsv", "ods",
  "docx", "docm", "dotx", "dotm", "odt", "rtf", "doc",
  "pptx", "pptm", "ppsx", "ppsm", "potx", "potm", "odp",
];
check("extension set matches the router", fullSet.every((ext) => extensions.indexOf(ext) !== -1) && extensions.length === fullSet.length, extensions.join(" "));

check("commands registered", registered.commands.length >= 5, JSON.stringify(registered.commands.map((c) => c.id)));
for (const id of ["copy-selection-with-formatting", "copy-selection-plain", "select-all-cells", "reload-document"]) {
  check("command " + id, registered.commands.some((c) => c.id === id), registered.commands.map((c) => c.id).join(","));
}
check("settings tab registered", registered.settingTabs.length === 1);
check("file details command", registered.commands.some((c) => c.id === "toggle-file-details"), registered.commands.map((c) => c.id).join(","));

// The view must construct and paint its chrome without a real workspace.
const viewFactory = registered.views[0] && registered.views[0].factory;
const { WorkspaceLeaf } = require("obsidian");
const leaf = new WorkspaceLeaf();
const view = viewFactory ? viewFactory(leaf) : null;
check("view constructs", Boolean(view));
if (view) {
  await view.onOpen();
  check("view builds chrome", Boolean(view.contentEl.querySelector(".ov-toolbar")));
  check("empty state shown", (view.contentEl.textContent || "").indexOf("No document open") !== -1);
}

console.log("");
console.log("smoke:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
