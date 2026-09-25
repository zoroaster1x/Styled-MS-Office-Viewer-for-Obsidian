/*
 * The WebAssembly JPEG XR codec, on a package that really carries a .wdp part.
 *
 * The lecture decks store these as a14:imgLayer effect layers over a rendered
 * base picture, so the part is not a slide shape and the base is what a slide
 * draws; this suite proves the decoder itself works through the same media
 * layer the renderers use: the part resolves to a real PNG, and the named
 * undecodable placeholder is gone.
 *
 * Reads a path from .testenv (OV_TEST_JXR) and skips with a notice when it is
 * missing. The file is read only, like every other suite.
 *
 *   bun test/jxr.mjs [<file>]
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

const arg = process.argv.slice(2).find((a) => !a.startsWith("--"));
const diskPath = arg || testPath("OV_TEST_JXR");
if (!diskPath || !existsSync(diskPath)) {
  skipNotice("OV_TEST_JXR", "JPEG XR run");
  process.exit(0);
}

const PluginClass = require(join(root, "main.js"));
const { WorkspaceLeaf, TFile } = require("obsidian");

const opened = [];
const app = {
  workspace: { getLeavesOfType: () => [], getActiveViewOfType: () => null },
  vault: {
    on: () => ({}),
    adapter: { getFullPath: (p) => p },
    readBinary: async () => {
      const bytes = readFileSync(opened[opened.length - 1].diskPath);
      return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    },
  },
};

const plugin = new PluginClass();
plugin.app = app;
plugin.settings = { fileState: {} };
plugin.saveFileState = async () => {};
plugin.updateSetting = async (key, value) => {
  plugin.settings[key] = value;
};
await plugin.onload();

const file = new TFile(basename(diskPath));
file.path = basename(diskPath);
opened.push({ diskPath, file });

const view = plugin.registered.views[0].factory(new WorkspaceLeaf());
view.app = app;
await view.onOpen();
await view.onLoadFile(file);

let pass = 0;
let fail = 0;
const check = (ok, label, detail) => {
  console.log(ok ? "ok  " : "FAIL", label.padEnd(44), detail || "");
  if (ok) pass++;
  else fail++;
};

const model = view.controller && view.controller.model;
if (!model || !model.pkg) {
  console.log("FAIL the view did not mount a presentation");
  process.exit(1);
}

const jxrParts = (model.pkg.list ? model.pkg.list() : []).filter((name) => /\.(wdp|hdp|jxr)$/i.test(name));
check(jxrParts.length > 0, "package carries JPEG XR parts", jxrParts.join(", "));

for (const path of jxrParts.slice(0, 2)) {
  const url = model.mediaUrl(path);
  const okUrl = typeof url === "string" && url.startsWith("data:image/png;base64,");
  let size = "";
  if (okUrl) {
    const png = Buffer.from(url.slice("data:image/png;base64,".length), "base64");
    const signature = png.subarray(0, 8).toString("hex") === "89504e470d0a1a0a";
    const width = png.readUInt32BE(16);
    const height = png.readUInt32BE(20);
    if (signature && width > 0 && height > 0) size = width + "x" + height;
  }
  check(Boolean(size), "decodes " + path, size || "no PNG");
}

// The renderers must not fall back to the named placeholder for a picture they
// can draw; the layer is not a slide shape, so a mounted slide keeps its base
// picture and the media layer draws it as a PNG.
const placeholder = /not decodable here/.test(view.contentEl.textContent || "");
check(!placeholder, "no undecodable placeholder for a decodable part");

view.onClose();
console.log("");
console.log("jxr:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
