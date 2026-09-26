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

// Inline Word drawing groups: the chOff/chExt mapping, a preset rectangle with
// fill and line, a custom dashed connector, a picture and a text box. A group
// is what makes a document's diagrams and photographs appear at all.
//
//   bun test/drawing.mjs

import { createRequire } from "node:module";
import { setupDom, createContainer } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { renderShapeGroup } = require("../src/docx/drawing.js");

let pass = 0;
let fail = 0;
function check(name, condition, detail) {
  if (condition) { pass++; console.log("ok   " + name); }
  else { fail++; console.log("FAIL " + name + (detail !== undefined ? "  -> " + detail : "")); }
}

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  + ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
  + ' xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"'
  + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
  + ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
  + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

// The group maps a 500000 EMU child space onto a 1000000 EMU box, so every
// child coordinate doubles.
const GROUP_XML = `<wpg:wgp ${NS}>`
  + '<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/>'
  + '<a:chOff x="0" y="0"/><a:chExt cx="500000" cy="500000"/></a:xfrm></wpg:grpSpPr>'
  + '<wps:wsp><wps:cNvPr id="1" name="Box 1"/><wps:spPr>'
  + '<a:xfrm><a:off x="250000" y="250000"/><a:ext cx="100000" cy="100000"/></a:xfrm>'
  + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
  + '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'
  + '<a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></wps:spPr>'
  + '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Label</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
  + '<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/></wps:wsp>'
  + '<wps:wsp><wps:cNvPr id="2" name="Dashed 2"/><wps:spPr>'
  + '<a:xfrm><a:off x="0" y="0"/><a:ext cx="500000" cy="0"/></a:xfrm>'
  + '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
  + '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
  + '<a:custDash><a:ds d="300000" sp="225000"/></a:custDash></a:ln></wps:spPr></wps:wsp>'
  + '<pic:pic><pic:nvPicPr><pic:cNvPr id="3" name="Pic 3"/><pic:cNvPicPr/></pic:nvPicPr>'
  + '<pic:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
  + '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="200000" cy="200000"/></a:xfrm>'
  + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
  + '</wpg:wgp>';

const doc = new DOMParser().parseFromString("<?xml version=\"1.0\"?>" + GROUP_XML, "application/xml");
const node = doc.documentElement;
check("the group XML parses", node && node.tagName.indexOf("wgp") !== -1, node && node.tagName);

const host = createContainer();
renderShapeGroup(host, { node, widthPx: 105, heightPx: 105, name: "Group 1" }, {
  theme: null,
  mediaUrl: (rid) => (rid === "rId9" ? "data:image/png;base64,AAAA" : null),
  parseParagraph: (p) => {
    const text = p.textContent || "";
    return { type: "p", props: {}, style: null, markRunProps: {}, runs: text ? [{ type: "text", props: {}, text }] : [] };
  },
  drawParagraph: (block, box) => {
    const span = box.createSpan("test-p");
    span.textContent = (block.runs[0] && block.runs[0].text) || "";
  },
});

const wrapper = host.querySelector(".ov-docx-shapegroup");
check("the group box takes the wp:extent size", wrapper && wrapper.style.width === "105px" && wrapper.style.height === "105px", wrapper && wrapper.style.cssText);
const shapes = wrapper ? Array.from(wrapper.querySelectorAll(".ov-docx-shape")) : [];
check("every child becomes a shape", shapes.length === 3, String(shapes.length));
const box = shapes[0];
const boxLeft = box ? Math.round(parseFloat(box.style.left) * 100) / 100 : NaN;
const boxWidth = box ? Math.round(parseFloat(box.style.width) * 100) / 100 : NaN;
check("chOff/chExt scales the child box", boxLeft === 52.49 && boxWidth === 21, boxLeft + "x" + boxWidth);
const rect = box && box.querySelector("rect");
const rectFill = rect ? rect.getAttribute("fill").toLowerCase() : "";
const rectStroke = rect ? rect.getAttribute("stroke").toLowerCase() : "";
check("a rect gets its fill and stroke", rectFill === "#ff0000" && rectStroke === "#0000ff", rectFill + "/" + rectStroke);
check("the text box text is drawn", box && (box.textContent || "").indexOf("Label") !== -1, box && box.textContent);
const dashed = shapes[1] && shapes[1].querySelector("polyline");
check("a custDash line becomes a dash array", dashed && dashed.getAttribute("stroke-dasharray") === "3 2.25", dashed && dashed.getAttribute("stroke-dasharray"));
const img = wrapper && wrapper.querySelector("img");
check("a picture becomes an image", img && img.getAttribute("src") === "data:image/png;base64,AAAA", img && img.getAttribute("src"));

// A group whose own ext is zero is an identity transform: its box is chExt and
// its children keep their coordinates.
const identityXml = `<wpg:wgp ${NS}>`
  + '<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
  + '<a:chOff x="0" y="0"/><a:chExt cx="500000" cy="500000"/></a:xfrm></wpg:grpSpPr>'
  + '<wps:wsp><wps:cNvPr id="4" name="Box 4"/><wps:spPr>'
  + '<a:xfrm><a:off x="100000" y="100000"/><a:ext cx="50000" cy="50000"/></a:xfrm>'
  + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr></wps:wsp></wpg:wgp>';
const identityDoc = new DOMParser().parseFromString("<?xml version=\"1.0\"?>" + identityXml, "application/xml");
const identityHost = createContainer();
renderShapeGroup(identityHost, { node: identityDoc.documentElement, widthPx: 52.49, heightPx: 52.49 }, {
  theme: null,
  mediaUrl: () => null,
  parseParagraph: () => null,
  drawParagraph: () => {},
});
const identityBox = identityHost.querySelector(".ov-docx-shape");
const identityLeft = identityBox ? Math.round(parseFloat(identityBox.style.left) * 100) / 100 : NaN;
check("an identity group keeps child coordinates", identityLeft === 10.5, String(identityLeft));

console.log("");
console.log("drawing:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
