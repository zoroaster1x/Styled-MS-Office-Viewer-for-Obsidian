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

// Windowed page rendering: a long document draws only the pages near the
// viewport, renders a page when the reader scrolls to it, gives the DOM back
// when they leave, and still finds text and headings on pages that were never
// drawn.
//
//   bun test/virtual-docx.mjs

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
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const para = (text) => ({
  type: "p",
  props: {},
  style: null,
  markRunProps: {},
  runs: [{ type: "text", props: {}, text }],
});

const body = [];
for (let i = 0; i < 900; i++) {
  body.push(para(i === 500 ? "UNIQUE_MARKER" : "Paragraph " + i));
}

const model = {
  kind: "docx",
  body,
  section: { pageWidthTw: 11906, pageHeightTw: 16838, marginTopTw: 1440, marginRightTw: 1440, marginBottomTw: 1440, marginLeftTw: 1440 },
  styles: { paragraph: new Map(), character: new Map(), table: new Map(), docDefaults: { pPr: {}, rPr: { sizeHalfPt: 24 } } },
  footnotes: new Map(),
  endnotes: new Map(),
  headers: new Map(),
  footers: new Map(),
  mediaCache: null,
  properties: {},
  mediaUrl() { return null; },
};

const host = createContainer();
let outline = null;
const renderer = createDocxRenderer({
  container: host,
  model,
  settings: { zoom: 1, virtualize: true },
  onReady: (info) => { outline = info.outline; },
});

const pages = Array.from(host.querySelectorAll(".ov-docx-page"));
check("every page shell exists", pages.length > 8, String(pages.length));

const drawn = () => pages.filter((el) => el.querySelector(".ov-docx-p"));
check("only the first pages are drawn at mount", drawn().length === 4, String(drawn().length));
check("page 1 carries its paragraphs", Boolean(pages[0].querySelector(".ov-docx-p")));
check("a far page is still empty", !pages[pages.length - 1].querySelector(".ov-docx-p"));

// Scroll so the middle of the document is on screen.
const scrollEl = host.querySelector(".ov-docx-scroll");
scrollEl.getBoundingClientRect = () => ({ top: 0, bottom: 800, height: 800, width: 800, left: 0, right: 800 });
pages.forEach((el, i) => {
  el.getBoundingClientRect = () => ({
    top: i * 1100 - scrollEl.scrollTop,
    bottom: (i + 1) * 1100 - scrollEl.scrollTop,
    height: 1100, width: 794, left: 0, right: 794,
  });
});
const middle = Math.floor(pages.length / 2);
scrollEl.scrollTop = middle * 1100;
scrollEl.dispatchEvent(new globalThis.Event("scroll"));
await wait(100);
check("scrolling draws the page it lands on", Boolean(pages[middle].querySelector(".ov-docx-p")));
check("pages left behind give their DOM back", !pages[0].querySelector(".ov-docx-p"));

// Search works over the model, not the drawn DOM: the marker is on an undrawn
// page, and stepping to the hit draws that page.
scrollEl.scrollTop = 0;
scrollEl.dispatchEvent(new globalThis.Event("scroll"));
await wait(100);
const count = renderer.search("UNIQUE_MARKER");
check("search finds text on an undrawn page", count === 1, String(count));
const step = renderer.searchNext(1);
check("stepping to a hit returns its place", step && step.count === 1, JSON.stringify(step));
const marked = host.querySelector(".ov-search-current");
check("the hit's page is drawn and marked", Boolean(marked && marked.textContent.indexOf("UNIQUE_MARKER") !== -1),
  marked ? marked.textContent : "none");

renderer.clearSearch();
check("clearing the search drops the mark", !host.querySelector(".ov-search-current"));

// An internal link to a bookmark on a far page draws that page first.
renderer.destroy();
console.log("");
console.log("virtual-docx:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
