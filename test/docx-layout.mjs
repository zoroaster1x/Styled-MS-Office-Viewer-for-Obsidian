/*
 * The document layout rules that a form depends on, pinned on a model built
 * here: an empty paragraph is still a line at the paragraph mark's size, blank
 * paragraphs count toward the pages, the renderer does not invent spacing the
 * file does not state, and the zoom the buttons set is actually wired to CSS.
 *
 *   bun test/docx-layout.mjs
 */

import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { setupDom, createContainer } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { createDocxRenderer } = require("../src/docx/render.js");

let pass = 0;
let fail = 0;
function check(name, condition, detail) {
  if (condition) { pass++; console.log("ok   " + name); }
  else { fail++; console.log("FAIL " + name + (detail !== undefined ? "  -> " + detail : "")); }
}

const para = (text) => ({
  type: "p",
  props: {},
  style: null,
  markRunProps: {},
  runs: text ? [{ type: "text", props: {}, text }] : [],
});

const body = [para("Heading A")];
for (let i = 0; i < 40; i++) body.push(para(""));
body.push(para("Heading B"));
for (let i = 0; i < 40; i++) body.push(para(""));

const model = {
  kind: "docx",
  body,
  section: { pageWidthTw: 11906, pageHeightTw: 16838, marginTopTw: 1440, marginRightTw: 1440, marginBottomTw: 1440, marginLeftTw: 1440 },
  // The document default is 12pt (sz 24); the empty paragraphs carry no run of
  // their own, so this is the size their line box has to use.
  styles: { paragraph: new Map(), character: new Map(), table: new Map(), docDefaults: { pPr: {}, rPr: { sizeHalfPt: 24 } } },
  footnotes: new Map(),
  endnotes: new Map(),
  headers: new Map(),
  footers: new Map(),
  mediaCache: null,
  properties: {},
  mediaUrl() { return null; },
};

let pageCount = 0;
const host = createContainer();
const renderer = createDocxRenderer({
  container: host,
  model,
  settings: { zoom: 1 },
  onReady: (info) => { pageCount = info.pageCount || 0; },
});

const paras = Array.from(host.querySelectorAll(".ov-docx-p"));
const empties = paras.filter((el) => !(el.textContent || "").length);
check("the empty paragraphs are drawn", empties.length === 80, String(empties.length));
check("an empty paragraph carries the mark's size", empties[0] && empties[0].style.fontSize === "16px", empties[0] && empties[0].style.fontSize);
check("the renderer invents no spacing", paras.every((el) => !el.style.marginBottom), paras.find((el) => el.style.marginBottom) ? paras.find((el) => el.style.marginBottom).style.marginBottom : "");
check("blank lines count toward the pages", pageCount >= 2, String(pageCount));

// The zoom the buttons write has to reach the stylesheet: the renderer only
// sets the variable, the CSS scales the pages with it (a transform would leave
// the scroll extent behind).
const css = readFileSync(new URL("../styles.css", import.meta.url), "utf8");
check("document zoom is wired to the stylesheet", /zoom:\s*var\(--ov-docx-zoom/.test(css));
check("an empty paragraph produces a line box", /\.ov-docx-p:empty::after/.test(css));

// setSettings carries a new zoom to the element.
renderer.setSettings({ zoom: 1.5 });
const rootEl = host.querySelector(".ov-docx");
check("setSettings writes the zoom variable", rootEl && rootEl.style.getPropertyValue("--ov-docx-zoom") === "1.5", rootEl && rootEl.style.getPropertyValue("--ov-docx-zoom"));

renderer.destroy();
console.log("");
console.log("docx-layout:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
