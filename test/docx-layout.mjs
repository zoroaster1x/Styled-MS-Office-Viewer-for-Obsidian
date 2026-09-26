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

// CSS zoom scales the pages, so the scroll offset has to follow the same ratio
// or zooming out walks the reader down the document.
const scrollEl = host.querySelector(".ov-docx-scroll");
scrollEl.scrollTop = 900;
renderer.setSettings({ zoom: 0.75 });
check("zooming keeps the scroll position in view", scrollEl.scrollTop === 450, scrollEl.scrollTop);
renderer.destroy();

// The page callback follows the scroll position. Rect geometry is patched
// because linkedom lays nothing out.
const pageHost = createContainer();
let lastPage = null;
const pageRenderer = createDocxRenderer({
  container: pageHost,
  model,
  settings: { zoom: 1 },
  onPageChange: (info) => { lastPage = info; },
});
check("the page callback reports a page and the count", lastPage && lastPage.page >= 1 && lastPage.count === 2, lastPage && JSON.stringify(lastPage));
const pageScroll = pageHost.querySelector(".ov-docx-scroll");
const pageEls = Array.from(pageHost.querySelectorAll(".ov-docx-page"));
pageEls.forEach((p, i) => {
  p.getBoundingClientRect = () => ({
    top: i * 1000 - pageScroll.scrollTop,
    bottom: (i + 1) * 1000 - pageScroll.scrollTop,
    height: 1000, width: 800, left: 0, right: 800,
  });
});
pageScroll.getBoundingClientRect = () => ({ top: 0, bottom: 800, height: 800, width: 800, left: 0, right: 800 });
pageScroll.scrollTop = 1200;
pageScroll.dispatchEvent(new globalThis.Event("scroll"));
await new Promise((resolve) => setTimeout(resolve, 80));
check("the page callback follows the scroll", lastPage && lastPage.page === 2, lastPage && JSON.stringify(lastPage));
pageRenderer.destroy();

// A table with a gridSpan and a vMerge pair keeps its grid: the span becomes a
// real attribute, the continuation cell is not emitted, the restart cell gets
// the rowspan, and the declared cell margins replace the stylesheet padding.
const tableBlock = {
  type: "table",
  props: { cellMargin: { left: 0, right: 0 } },
  grid: [1000, 1000, 1000],
  rows: [
    { props: {}, cells: [{ props: { gridSpan: 3 }, blocks: [para("wide")], gridSpan: 3, vMerge: null }] },
    {
      props: {},
      cells: [
        { props: { vMerge: "restart" }, blocks: [para("tall")], gridSpan: 1, vMerge: "restart" },
        { props: {}, blocks: [para("right")], gridSpan: 1, vMerge: null },
      ],
    },
    {
      props: {},
      cells: [
        { props: { vMerge: "continue" }, blocks: [], gridSpan: 1, vMerge: "continue" },
        { props: {}, blocks: [para("right2")], gridSpan: 1, vMerge: null },
      ],
    },
  ],
};
const tableModel = Object.assign({}, model, { body: [tableBlock] });
const tableHost = createContainer();
createDocxRenderer({ container: tableHost, model: tableModel, settings: {} });
const tableRows = Array.from(tableHost.querySelectorAll("tr"));
const wideCell = tableRows[0] && tableRows[0].children[0];
check("a gridSpan cell carries a colspan attribute", wideCell && wideCell.getAttribute("colspan") === "3", wideCell && wideCell.getAttribute("colspan"));
check("a vMerge restart cell carries a rowspan", tableRows[1] && tableRows[1].children[0].getAttribute("rowspan") === "2", tableRows[1] && tableRows[1].children[0].getAttribute("rowspan"));
check("a vMerge continuation cell is not emitted", tableRows[2] && tableRows[2].children.length === 1, tableRows[2] && tableRows[2].children.length);
check("the table cell margins are applied", wideCell && wideCell.style.paddingLeft === "0px" && wideCell.style.paddingTop === "0px", wideCell && wideCell.style.paddingLeft);

// A PAGE field must show the live page, not the saved result. One footer in a
// real manual carries 53 while the document has 31 pages.
const fieldBlock = {
  type: "p", props: {}, style: null, markRunProps: {},
  runs: [
    { type: "text", props: {}, text: "Page " },
    { type: "fieldChar", stage: "begin" },
    { type: "instr", text: " PAGE   \\* MERGEFORMAT " },
    { type: "fieldChar", stage: "separate" },
    { type: "text", props: {}, text: "53" },
    { type: "fieldChar", stage: "end" },
  ],
};
const fieldModel = Object.assign({}, model, { body: [fieldBlock] });
const fieldHost = createContainer();
createDocxRenderer({ container: fieldHost, model: fieldModel, settings: {} });
const fieldText = (fieldHost.textContent || "").replace(/\s+/g, " ").trim();
const fieldMarkers = Array.from(fieldHost.querySelectorAll(".ov-docx-field")).map((el) => el.textContent).join(",");
check("a PAGE field shows the live number", fieldMarkers === "1", JSON.stringify(fieldMarkers));
check("the saved field result is dropped", !fieldText.includes("53"), fieldText);

console.log("");
console.log("docx-layout:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
