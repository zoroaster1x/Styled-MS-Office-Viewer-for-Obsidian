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

// Legacy VML drawings: w:pict shapes, their coordinate spaces and their
// textboxes. A file that carries only VML used to render as empty gaps.
//
//   bun test/vml.mjs

import { createRequire } from "node:module";
import { setupDom, createContainer } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { parseVmlRun, renderVml } = require("../src/docx/vml.js");

let pass = 0;
let fail = 0;
function check(name, condition, detail) {
  if (condition) { pass++; console.log("ok   " + name); }
  else { fail++; console.log("FAIL " + name + (detail !== undefined ? "  -> " + detail : "")); }
}

const PICT = `<?xml version="1.0"?>
<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:v="urn:schemas-microsoft-com:vml"
        xmlns:w10="urn:schemas-microsoft-com:office:word"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <v:shape type="#_x0000_t202" style="position:absolute;margin-left:100pt;margin-top:20pt;width:200pt;height:50pt;mso-wrap-style:square" strokecolor="#0070c0" strokeweight="2.25pt">
    <v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>y2 = 2r0 x</w:t></w:r></w:p></w:txbxContent></v:textbox>
  </v:shape>
</w:pict>`;

const doc = new DOMParser().parseFromString(PICT, "application/xml");
const run = parseVmlRun(doc.documentElement);
check("the pict parses as a VML run", run && run.type === "vml", run && run.type);
check("the size comes from the style in points", run && Math.round(run.widthPx) === 267 && Math.round(run.heightPx) === 67, run && `${run.widthPx}x${run.heightPx}`);
check("a square wrap stays in flow", run && run.anchor === null, JSON.stringify(run && run.anchor));
check("the textbox text is collected", run && run.texts.join(" ") === "y2 = 2r0 x", run && run.texts.join(" "));

const host = createContainer();
renderVml(host, run, {
  theme: null,
  contentWidthPx: 558,
  pageMarginLeftPx: 120,
  pageMarginTopPx: 95,
  mediaUrl: () => null,
  parseParagraph: () => ({ type: "p", props: {}, style: null, markRunProps: {}, runs: [] }),
  drawParagraph: (block, box) => { box.createDiv("test-p").textContent = "para"; },
});
const wrapper = host.querySelector(".ov-docx-vml");
check("the pict wraps in a vml box", Boolean(wrapper), wrapper && wrapper.className);
const wrapperWidth = wrapper ? Math.round(parseFloat(wrapper.style.width) * 100) / 100 : NaN;
const wrapperMargin = wrapper ? Math.round(parseFloat(wrapper.style.marginLeft) * 100) / 100 : NaN;
check("the box takes its size and margins", wrapperWidth === 266.67 && wrapperMargin === 133.33,
  wrapper && wrapper.style.cssText);
check("the shape draws its outline", Boolean(wrapper && wrapper.querySelector("svg rect")), wrapper && wrapper.innerHTML.slice(0, 120));
check("the text runs through the document renderer", Boolean(wrapper && wrapper.querySelector(".test-p")), wrapper && wrapper.innerHTML.slice(-160));

// A shape with wrap=none floats: positioned from the style offsets and taking
// no room in the line.
const FLOAT = PICT.replace("mso-wrap-style:square", "mso-wrap-style:none");
const floatDoc = new DOMParser().parseFromString(FLOAT, "application/xml");
const floatRun = parseVmlRun(floatDoc.documentElement);
check("wrap none floats the shape", floatRun && floatRun.anchor && floatRun.anchor.outOfFlow === true, JSON.stringify(floatRun && floatRun.anchor));
const floatHost = createContainer();
renderVml(floatHost, floatRun, {
  theme: null, contentWidthPx: 558, pageMarginLeftPx: 120, pageMarginTopPx: 95,
  mediaUrl: () => null, parseParagraph: () => null, drawParagraph: () => {},
});
const floated = floatHost.querySelector(".ov-docx-vml");
const floatedLeft = floated ? Math.round(parseFloat(floated.style.left) * 100) / 100 : NaN;
const floatedTop = floated ? Math.round(parseFloat(floated.style.top) * 100) / 100 : NaN;
check("a floating pict is absolute", floated && floated.style.position === "absolute", floated && floated.style.position);
check("a floating pict uses its offsets", floatedLeft === 133.33 && floatedTop === 26.67, floated && floated.style.cssText);

// A group scales its children from the group coordinate space. The child at
// coordinate 1000 with coordsize 2000 over a 100px box lands at 50px.
const GROUP = `<?xml version="1.0"?>
<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:v="urn:schemas-microsoft-com:vml"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <v:group style="position:absolute;margin-left:10pt;margin-top:10pt;width:100pt;height:50pt;mso-wrap-style:square" coordorigin="0,0" coordsize="2000,1000">
    <v:rect style="position:absolute;left:1000;top:500;width:1000;height:500" fillcolor="#ff0000" stroked="f"/>
  </v:group>
</w:pict>`;
const groupDoc = new DOMParser().parseFromString(GROUP, "application/xml");
const groupRun = parseVmlRun(groupDoc.documentElement);
const groupHost = createContainer();
renderVml(groupHost, groupRun, {
  theme: null, contentWidthPx: 558, pageMarginLeftPx: 120, pageMarginTopPx: 95,
  mediaUrl: () => null, parseParagraph: () => null, drawParagraph: () => {},
});
const groupRect = groupHost.querySelector(".ov-docx-vml-group .ov-docx-shape svg rect");
const groupBox = groupHost.querySelector(".ov-docx-vml-group .ov-docx-shape");
const groupLeft = groupBox ? Math.round(parseFloat(groupBox.style.left) * 100) / 100 : NaN;
const groupWidth = groupBox ? Math.round(parseFloat(groupBox.style.width) * 100) / 100 : NaN;
check("a group child is scaled from the coordinate space", groupLeft === 66.67 && groupWidth === 66.67,
  groupBox && groupBox.style.cssText);
check("the child keeps its fill", groupRect && groupRect.getAttribute("fill") === "#ff0000", groupRect && groupRect.getAttribute("fill"));

console.log("");
console.log("vml:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
