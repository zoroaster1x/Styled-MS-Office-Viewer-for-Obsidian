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

// Document cache: the LRU rules, the invalidation key, media release on
// eviction, and the loader's stale token when two files are opened in quick
// succession. The unit half needs nothing; the loader half opens a real
// document, so it skips with a notice when .testenv is unset.
//
//   bun test/cache.mjs

import { createRequire } from "node:module";
import { existsSync, readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { setupDom } from "./harness.mjs";
import { testPath, skipNotice } from "./env.mjs";
import { DocumentCache, releaseModel } from "../src/app/document-cache.js";

setupDom();
const require = createRequire(import.meta.url);
const root = join(dirname(fileURLToPath(import.meta.url)), "..");

let pass = 0;
let fail = 0;
function check(name, condition, detail) {
  if (condition) {
    pass++;
    return;
  }
  fail++;
  console.log("FAIL " + name + (detail ? "  -> " + detail : ""));
}

// ---------- the cache itself ----------

const released = [];
function fakeModel(id) {
  return { id, mediaCache: { release() { released.push(id); } } };
}

const cache = new DocumentCache({ limit: 2 });
const key = (path, size, mtime) => DocumentCache.keyFor(path, { size, mtime });

cache.store(key("a.docx", 10, 1), { model: fakeModel("a"), format: { ext: "docx" } }, 100);
check("a stored model is found by its key", cache.lookup(key("a.docx", 10, 1)).model.id === "a");
check("a changed size misses", cache.lookup(key("a.docx", 11, 1)) === null);
check("a changed mtime misses", cache.lookup(key("a.docx", 10, 2)) === null);

cache.store(key("b.docx", 10, 1), { model: fakeModel("b") }, 100);
cache.store(key("c.docx", 10, 1), { model: fakeModel("c") }, 100);
check("the LRU evicts the oldest", cache.lookup(key("a.docx", 10, 1)) === null, "a should be gone");
check("eviction releases the media", released.indexOf("a") !== -1);
check("the two most recent stay", cache.lookup(key("b.docx", 10, 1)) !== null && cache.lookup(key("c.docx", 10, 1)) !== null);

check("an oversized document is refused", cache.store(key("d.docx", 10, 1), { model: fakeModel("d") }, 999 * 1024 * 1024) === false);
check("a refused model is not released by the cache", released.indexOf("d") === -1);
releaseModel({ mediaCache: { release() { released.push("d"); } } });
check("releaseModel can be called directly", released.indexOf("d") !== -1);

cache.dropPath("b.docx");
check("dropPath removes the entry", cache.lookup(key("b.docx", 10, 1)) === null);
check("dropPath releases the media", released.indexOf("b") !== -1);

cache.clear();
check("clear releases everything left", released.indexOf("c") !== -1);

// ---------- the loader ----------

const file = testPath("OV_TEST_DOCX") || testPath("OV_TEST_PPTX") || testPath("OV_TEST_XLSX");
if (!file || !existsSync(join(root, "main.js"))) {
  skipNotice("OV_TEST_DOCX", "cache loader checks");
  console.log(`cache: ${pass} pass, ${fail} fail`);
  process.exit(fail ? 1 : 0);
}

const bytes = new Uint8Array(readFileSync(file));
const PluginClass = require(join(root, "main.js"));
const obsidian = require("obsidian");
const plugin = new PluginClass();

let reads = 0;
let delayMs = 0;
const vaultEvents = new Map();
plugin.app = {
  workspace: { getLeavesOfType: () => [], getActiveViewOfType: () => null, onLayoutReady: (fn) => fn() },
  vault: {
    readBinary: async () => {
      reads++;
      if (delayMs) await new Promise((resolve) => setTimeout(resolve, delayMs));
      return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    },
    on: (event, handler) => {
      vaultEvents.set(event, handler);
      return { event };
    },
    adapter: { getFullPath: (path) => path },
  },
  metadataCache: { getFileCache: () => null },
};
await plugin.onload();

const view = plugin.registered.views[0].factory(new obsidian.WorkspaceLeaf());
// Obsidian gives every view the app; the stub's views carry a shared default,
// so point this one at the counting vault.
view.app = plugin.app;
check("the view factory builds a view", Boolean(view && typeof view.renderFile === "function"));
if (view && view.onOpen) await view.onOpen();

function makeFile(path, mtime) {
  const file = new obsidian.TFile(path);
  file.stat = { size: bytes.byteLength, mtime };
  return file;
}

const target = makeFile("cache-target.docx", 1000);
await view.renderFile(target);
check("the first open reads the file", reads === 1, "reads " + reads);
check("the first open parses a model", Boolean(view.controller && view.controller.model));
check("the first open stays cached", Boolean(view.controller && view.controller.model && view.controller.model.__ovCached));
const model = view.controller.model;
const controller = view.controller;

await view.renderFile(target);
check("the second open is served from the cache", reads === 1, "reads " + reads);
check("the cached model is reused", view.controller !== controller && view.controller.model === model);
check("no error from a cached open", !view.error, view.error && String(view.error));

target.stat = { size: bytes.byteLength, mtime: 2000 };
await view.renderFile(target);
check("a modified file is parsed again", reads === 2, "reads " + reads);

// The vault event must drop the entry even before the next open.
const before = reads;
const modified = makeFile("cache-target.docx", 2000);
if (vaultEvents.has("modify")) vaultEvents.get("modify")(modified);
await view.renderFile(modified);
check("the modify event invalidates the cache", reads === before + 1, "reads " + reads);

// Two opens in flight: the earlier one must not mount over the later one.
const slow = makeFile("cache-slow.docx", 3000);
const fast = makeFile("cache-fast.docx", 3000);
delayMs = 40;
const slowOpen = view.renderFile(slow);
delayMs = 0;
const fastOpen = view.renderFile(fast);
await Promise.all([slowOpen, fastOpen]);
check("the later open is the one on screen", view.file === fast, view.file && view.file.path);
check("the stale open left no error", !view.error, view.error && String(view.error));

view.onClose();
console.log(`cache: ${pass} pass, ${fail} fail`);
if (fail) process.exit(1);
