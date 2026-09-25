/*
 * A presentation's own fonts. The fixture embeds Be Vietnam Pro as MicroType
 * Express compressed EOT parts; the parser has to name the faces and the loader
 * has to turn them into registrable TrueType faces, so text renders in the
 * author's font even where nothing is installed. The fixture and its OFL
 * licence are described in test/fixtures/embedded-font/PROVENANCE.md.
 *
 *   bun test/embed-font.mjs
 */

import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { setupDom } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { parsePptx } = require("../src/pptx/parse.js");
const { ensureEmbeddedFonts, releaseEmbeddedFonts } = require("../src/pptx/fonts.js");

let pass = 0;
let fail = 0;
function check(name, condition, detail) {
  if (condition) { pass++; console.log("ok   " + name); }
  else { fail++; console.log("FAIL " + name + (detail !== undefined ? "  -> " + detail : "")); }
}

const fixture = new URL("./fixtures/embedded-font/source.pptx", import.meta.url);
const model = parsePptx(new Uint8Array(readFileSync(fixture)));

const faces = model.embeddedFonts || [];
check("the faces are listed", faces.length === 2, String(faces.length));
check("the family is named", faces.every((face) => face.family === "Be Vietnam Pro"), faces.map((f) => f.family).join(","));
check("the weights are named", faces.some((face) => face.weight === "400") && faces.some((face) => face.weight === "700"), faces.map((f) => f.weight).join(","));

// Register into a recording font set; the harness stub normally swallows them.
const added = [];
const deleted = [];
document.fonts = {
  add(face) { added.push(face); },
  delete(face) { deleted.push(face); },
};

await ensureEmbeddedFonts(model);
check("both faces register", added.length === 2, String(added.length));
check("they register under the family", added.every((face) => face.family === "Be Vietnam Pro"));
check("the source is a TrueType font", added.every((face) => {
  const bytes = new Uint8Array(face.source);
  return bytes[0] === 0 && bytes[1] === 1 && bytes[2] === 0 && bytes[3] === 0;
}));
const utf16be = (text) => Buffer.from([...text].flatMap((ch) => [0, ch.charCodeAt(0)]));
check("the decompressed font carries its name", added.every((face) => Buffer.from(face.source).includes(utf16be("Be Vietnam Pro"))));

// A second ensure is a no-op: the model already owns its faces.
await ensureEmbeddedFonts(model);
check("registering is idempotent", added.length === 2, String(added.length));

releaseEmbeddedFonts(model);
check("release hands the faces back", deleted.length === 2, String(deleted.length));

// A face that will not decode is skipped without throwing.
const broken = Object.assign({}, model, {
  embeddedFonts: [{ family: "Broken", weight: "400", style: "normal", path: "ppt/fonts/missing.fntdata" }],
  __ovFontFaces: null,
  __ovFontsReady: false,
});
await ensureEmbeddedFonts(broken);
check("a broken face falls back to the host", broken.__ovFontFaces.length === 0);

console.log("");
console.log("embed-font:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
