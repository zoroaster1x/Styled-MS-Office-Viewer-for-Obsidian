/*
 * The presentation layout rules that only a real file can teach, pinned on a
 * package built here: subscript runs scale and shift, a flipped connector is
 * not flipped twice, a placeholder level carries its line spacing and spacing
 * before, normAutofit's line spacing reduction applies, a layout's date
 * placeholder is a prototype and does not print, and a slide's own slide
 * number placeholder resolves to the live number.
 *
 *   bun test/pptx-layout.mjs
 */

import { createRequire } from "node:module";
import { setupDom, createContainer } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { zipSync } = require("fflate");
const { parsePptx } = require("../src/pptx/parse.js");
const { createPptxRenderer } = require("../src/pptx/render.js");

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const textBody = (inner) => "<p:txBody><a:bodyPr/><a:lstStyle/>" + inner + "</p:txBody>";
const shape = (id, name, ph, box, inner) =>
  "<p:sp><p:nvSpPr><p:cNvPr id=\"" + id + "\" name=\"" + name + "\"/><p:cNvSpPr/><p:nvPr>" + ph + "</p:nvPr></p:nvSpPr>" +
  "<p:spPr><a:xfrm><a:off x=\"" + box[0] + "\" y=\"" + box[1] + "\"/><a:ext cx=\"" + box[2] + "\" cy=\"" + box[3] + "\"/></a:xfrm></p:spPr>" +
  inner + "</p:sp>";

const parts = {
  "_rels/.rels": "<Relationships xmlns=\"" + REL + "\"><Relationship Id=\"rId1\" Type=\"" + RT + "/officeDocument\" Target=\"ppt/presentation.xml\"/></Relationships>",
  "ppt/presentation.xml":
    "<p:presentation xmlns:p=\"" + P + "\" xmlns:r=\"" + R + "\"><p:sldIdLst><p:sldId id=\"256\" r:id=\"rId1\"/></p:sldIdLst>" +
    "<p:sldSz cx=\"9144000\" cy=\"6858000\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/></p:presentation>",
  "ppt/_rels/presentation.xml.rels":
    "<Relationships xmlns=\"" + REL + "\"><Relationship Id=\"rId1\" Type=\"" + RT + "/slide\" Target=\"slides/slide1.xml\"/></Relationships>",
  "ppt/slideMasters/slideMaster1.xml":
    "<p:sldMaster xmlns:p=\"" + P + "\" xmlns:a=\"" + A + "\" xmlns:r=\"" + R + "\"><p:cSld><p:spTree>" +
    "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>" +
    shape(2, "Master Date", "<p:ph type=\"dt\" idx=\"10\"/>", [0, 6000000, 2000000, 400000], textBody("<a:p><a:r><a:t>1/2/03</a:t></a:r></a:p>")) +
    "</p:spTree></p:cSld><p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>" +
    "<p:txStyles>" +
    "<p:titleStyle><a:lvl1pPr><a:lnSpc><a:spcPct val=\"90000\"/></a:lnSpc><a:spcBef><a:spcPct val=\"0\"/></a:spcBef><a:defRPr sz=\"4400\"/></a:lvl1pPr></p:titleStyle>" +
    "<p:bodyStyle><a:lvl1pPr marL=\"228600\" indent=\"-228600\"><a:lnSpc><a:spcPct val=\"90000\"/></a:lnSpc><a:spcBef><a:spcPts val=\"1000\"/></a:spcBef><a:buChar char=\"&#8226;\"/><a:defRPr sz=\"2800\"/></a:lvl1pPr></p:bodyStyle>" +
    "<p:otherStyle><a:lvl1pPr><a:defRPr sz=\"1800\"/></a:lvl1pPr></p:otherStyle>" +
    "</p:txStyles></p:sldMaster>",
  "ppt/slideMasters/_rels/slideMaster1.xml.rels":
    "<Relationships xmlns=\"" + REL + "\"><Relationship Id=\"rId1\" Type=\"" + RT + "/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/></Relationships>",
  "ppt/slideLayouts/slideLayout1.xml":
    "<p:sldLayout xmlns:p=\"" + P + "\" xmlns:a=\"" + A + "\" xmlns:r=\"" + R + "\" type=\"title\"><p:cSld><p:spTree>" +
    "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>" +
    shape(2, "Layout Date", "<p:ph type=\"dt\" idx=\"10\"/>", [0, 6000000, 2000000, 400000], textBody("<a:p><a:r><a:t>1/2/03</a:t></a:r></a:p>")) +
    "</p:spTree></p:cSld></p:sldLayout>",
  "ppt/slideLayouts/_rels/slideLayout1.xml.rels":
    "<Relationships xmlns=\"" + REL + "\"><Relationship Id=\"rId1\" Type=\"" + RT + "/slideMaster\" Target=\"../slideMasters/slideMaster1.xml\"/></Relationships>",
  "ppt/slides/slide1.xml":
    "<p:sld xmlns:p=\"" + P + "\" xmlns:a=\"" + A + "\" xmlns:r=\"" + R + "\"><p:cSld><p:spTree>" +
    "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>" +
    shape(2, "Title 1", "<p:ph type=\"title\"/>", [0, 0, 9144000, 1000000], textBody("<a:p><a:r><a:rPr lang=\"en-US\"/><a:t>Plain title</a:t></a:r></a:p>")) +
    shape(3, "Body 2", "<p:ph type=\"body\" idx=\"1\"/>", [0, 1000000, 4000000, 2000000],
      textBody("<a:p><a:r><a:rPr lang=\"en-US\"/><a:t>First point</a:t></a:r></a:p><a:p><a:r><a:rPr lang=\"en-US\"/><a:t>Second point</a:t></a:r></a:p>")) +
    shape(4, "Math 3", "", [5000000, 0, 3000000, 800000],
      textBody("<a:p><a:r><a:rPr lang=\"en-US\"/><a:t>Sens</a:t></a:r><a:r><a:rPr lang=\"en-US\" baseline=\"-25000\"/><a:t>BIN</a:t></a:r></a:p>")) +
    "<p:cxnSp><p:nvCxnSpPr><p:cNvPr id=\"5\" name=\"Arrow 4\"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>" +
    "<p:spPr><a:xfrm flipV=\"1\"><a:off x=\"5000000\" y=\"1000000\"/><a:ext cx=\"1000000\" cy=\"2000000\"/></a:xfrm>" +
    "<a:prstGeom prst=\"straightConnector1\"><a:avLst/></a:prstGeom>" +
    "<a:ln w=\"28575\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill><a:tailEnd type=\"triangle\"/></a:ln></p:spPr></p:cxnSp>" +
    shape(6, "Reduced 5", "", [0, 3000000, 3000000, 800000],
      "<p:txBody><a:bodyPr><a:normAutofit fontScale=\"90000\" lnSpcReduction=\"20000\"/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"/><a:t>Reduced</a:t></a:r></a:p></p:txBody>") +
    shape(7, "Slide Number 6", "<p:ph type=\"sldNum\" idx=\"12\"/>", [8000000, 6000000, 1000000, 400000],
      textBody("<a:p><a:fld id=\"{1}\" type=\"slidenum\"><a:rPr lang=\"en-US\"/><a:t>&#8249;#&#8250;</a:t></a:fld></a:p>")) +
    // A freeform: 200x100px, path space 100x100, an arc through 90 degrees.
    "<p:sp><p:nvSpPr><p:cNvPr id=\"8\" name=\"Freeform 7\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
    "<p:spPr><a:xfrm><a:off x=\"0\" y=\"4000000\"/><a:ext cx=\"1905000\" cy=\"952500\"/></a:xfrm>" +
    "<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l=\"0\" t=\"0\" r=\"r\" b=\"b\"/><a:pathLst><a:path w=\"100\" h=\"100\">" +
    "<a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo><a:lnTo><a:pt x=\"100\" y=\"0\"/></a:lnTo>" +
    "<a:arcTo wR=\"50\" hR=\"50\" stAng=\"0\" swAng=\"5400000\"/><a:close/>" +
    "</a:path></a:pathLst></a:custGeom>" +
    "<a:solidFill><a:srgbClr val=\"00B050\"/></a:solidFill>" +
    "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln></p:spPr>" +
    "<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>" +
    // A group whose child coordinate space is 1000 units mapped to 2000000
    // EMU: the child scales with the group, in plain arithmetic, with no CSS
    // transform and no font or stroke multiplication.
    "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"20\" name=\"Scaled Group\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
    "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"2000000\" cy=\"2000000\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"1000\" cy=\"1000\"/></a:xfrm></p:grpSpPr>" +
    "<p:sp><p:nvSpPr><p:cNvPr id=\"21\" name=\"Scaled Child\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
    "<p:spPr><a:xfrm><a:off x=\"500\" y=\"500\"/><a:ext cx=\"200\" cy=\"200\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>" +
    "<a:solidFill><a:srgbClr val=\"123456\"/></a:solidFill></p:spPr>" +
    "<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></p:grpSp>" +
    "</p:spTree></p:cSld></p:sld>",
  "ppt/slides/_rels/slide1.xml.rels":
    "<Relationships xmlns=\"" + REL + "\"><Relationship Id=\"rId1\" Type=\"" + RT + "/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/></Relationships>",
};

const encoder = new TextEncoder();
for (const [name, xml] of Object.entries(parts)) parts[name] = encoder.encode(xml);
const model = parsePptx(zipSync(parts, { level: 6 }));
const host = createContainer();
const renderer = createPptxRenderer({ container: host, model, settings: { fit: "none" } });

let pass = 0;
let fail = 0;
function check(name, condition, detail) {
  if (condition) { pass++; console.log("ok   " + name); }
  else { fail++; console.log("FAIL " + name + (detail !== undefined ? "  -> " + detail : "")); }
}

const shapeEls = Array.from(host.querySelectorAll(".ov-pptx-shape"));
const byText = (needle) => shapeEls.find((el) => (el.textContent || "").indexOf(needle) !== -1);
const px = (value) => parseFloat(String(value).replace("px", ""));

// Subscript: 58% of the run's size and shifted down by a quarter of the full em.
const math = byText("Sens");
const sens = math && Array.from(math.querySelectorAll(".ov-pptx-r")).find((el) => el.textContent === "Sens");
const bin = math && Array.from(math.querySelectorAll(".ov-pptx-r")).find((el) => el.textContent === "BIN");
check("subscript run exists", Boolean(sens && bin));
if (sens && bin) {
  const ratio = px(bin.style.fontSize) / px(sens.style.fontSize);
  check("subscript draws at 58 percent", Math.abs(ratio - 0.58) < 0.01, String(ratio));
  check("subscript shifts down", parseFloat(bin.style.verticalAlign) < 0, bin.style.verticalAlign);
}

// A flipped straight connector keeps the flip once, in the SVG coordinates.
const arrow = shapeEls.find((el) => el.querySelector("line"));
check("connector has an svg line", Boolean(arrow));
if (arrow) {
  const line = arrow.querySelector("line");
  check("connector element is not flipped twice", !arrow.style.scale, arrow.style.scale);
  check("flipV puts the line end at the top", line.getAttribute("y2") === "0" && line.getAttribute("y1") !== "0", line.getAttribute("y1") + " -> " + line.getAttribute("y2"));
}

// Paragraph levels carry their spacing: the title's 90% line spacing, the
// body list's 10pt before, and normAutofit's 20% line spacing reduction.
const title = byText("Plain title");
const titleP = title && title.querySelector(".ov-pptx-p");
check("title inherits 90 percent line spacing", titleP && titleP.style.lineHeight === "0.9", titleP && titleP.style.lineHeight);
const body = byText("Second point");
const second = body && Array.from(body.querySelectorAll(".ov-pptx-p")).find((el) => el.textContent.indexOf("Second") !== -1);
check("body list inherits its spacing before", second && Math.abs(px(second.style.marginTop) - 13.333) < 0.01, second && second.style.marginTop);
const reduced = byText("Reduced");
const reducedP = reduced && reduced.querySelector(".ov-pptx-p");
check("normAutofit reduces the line spacing", reducedP && Math.abs(px(reducedP.style.lineHeight) - 0.944) < 0.001, reducedP && reducedP.style.lineHeight);

// A freeform shape draws its own silhouette: the path is scaled from the
// path space to the shape box, the arcTo becomes an SVG arc with the computed
// endpoint, and the outline is stroked on the path rather than on the box.
const freeform = shapeEls.find((el) => el.querySelector('svg path[fill="#00B050"], svg path[fill="#00b050"]'));
check("a custom shape draws its path", Boolean(freeform));
if (freeform) {
  const path = freeform.querySelector("svg path");
  check("the path is scaled and the arc resolved",
    path.getAttribute("d") === "M0 0 L200 0 A100 50 0 0 1 100 50 Z", path.getAttribute("d"));
  check("the outline follows the path", path.getAttribute("stroke") === "#000000", path.getAttribute("stroke"));
  check("no box background is left behind", !freeform.style.background || freeform.style.background === "none", freeform.style.background);
}

// A scaled group multiplies geometry through its child space, without a CSS
// transform: the child at (500,500) 200x200 in a 1000 unit space mapped to
// 2000000 EMU lands at 105px with a 42px box.
const scaledChild = shapeEls.find((el) => (el.style.background || "").toLowerCase().indexOf("rgb(18, 52, 86)") !== -1 || (el.style.background || "").indexOf("#123456") !== -1);
check("a scaled group draws its child", Boolean(scaledChild), scaledChild && scaledChild.style.background);
if (scaledChild) {
  const left = parseFloat(scaledChild.style.left);
  const width = parseFloat(scaledChild.style.width);
  check("the group scale maps the position", Math.abs(left - 105) < 0.5, String(left));
  check("the group scale maps the size", Math.abs(width - 42) < 0.5, String(width));
  check("the wrapper carries no css transform", !scaledChild.parentElement || !scaledChild.parentElement.style.transform, scaledChild.parentElement && scaledChild.parentElement.style.transform);
}

// A layout or master date placeholder is a prototype, never content; the
// slide's own slide number placeholder is content and takes the live number.
check("layout date placeholder stays off the slide", (host.textContent || "").indexOf("1/2/03") === -1);
const slideNumber = byText("1");
check("slide number placeholder resolves", Boolean(slideNumber) && slideNumber.textContent.trim() === "1", slideNumber && JSON.stringify(slideNumber.textContent));

console.log("");
console.log("pptx-layout:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
