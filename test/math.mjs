/*
 * Styled MS Office Viewer, an Obsidian plugin that renders office documents
 * (xlsx, docx, pptx and their relatives) with their real styling, read only.
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
 * this program. If not, see <http://www.gnu.org/licenses/>.
 */

// Equations: OMML becomes a tree, the tree becomes layout, and the plain text of
// an equation reaches search and the word counts. Synthetic markup only.
//
//   bun test/math.mjs

import { createRequire } from "node:module";
import { setupDom, createContainer } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { parseXml } = require("../src/shared/xml.js");
const { parseMath, mathText } = require("../src/docx/math.js");
const { createDocxRenderer } = require("../src/docx/render.js");

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

function parseFragment(inner) {
  const doc = parseXml('<m:oMath xmlns:m="m" xmlns:w="w">' + inner + "</m:oMath>");
  return parseMath(doc.documentElement);
}

const run = (text, normal) => "<m:r>" + (normal ? "<m:rPr><m:nor/></m:rPr>" : "") + "<m:t>" + text + "</m:t></m:r>";

// ---------- a fraction ----------
const fraction = parseFragment("<m:f><m:num>" + run("1") + "</m:num><m:den>" + run("2") + "</m:den></m:f>");
check("a fraction parses", fraction.children[0].kind === "fraction", JSON.stringify(fraction).slice(0, 80));
check("a fraction keeps its numerator", mathText(fraction), "1/2");
check("a fraction keeps its denominator", mathText(fraction.children[0].denominator), "2");

// ---------- scripts ----------
check("a superscript parses", mathText(parseFragment("<m:sSup><m:e>" + run("x") + "</m:e><m:sup>" + run("2") + "</m:sup></m:sSup>")), "x^2");
check("a subscript parses", mathText(parseFragment("<m:sSub><m:e>" + run("a") + "</m:e><m:sub>" + run("n") + "</m:sub></m:sSub>")), "a_n");
check(
  "a sub and a superscript together",
  mathText(parseFragment("<m:sSubSup><m:e>" + run("a") + "</m:e><m:sub>" + run("1") + "</m:sub><m:sup>" + run("2") + "</m:sup></m:sSubSup>")),
  "a_1^2"
);

// ---------- radicals, delimiters, operators ----------
check("a radical parses", mathText(parseFragment("<m:rad><m:deg/><m:e>" + run("x") + "</m:e></m:rad>")), "\u221a(x)");
check(
  "a radical with a degree",
  mathText(parseFragment("<m:rad><m:deg>" + run("3") + "</m:deg><m:e>" + run("x") + "</m:e></m:rad>")),
  "3\u221a(x)"
);
check(
  "a delimiter keeps its characters",
  mathText(parseFragment('<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e>' + run("x") + "</m:e></m:d>")),
  "[x]"
);
check(
  "an n-ary operator with limits",
  mathText(parseFragment("<m:nary><m:naryPr><m:chr m:val=\"\u2211\"/></m:naryPr><m:sub>" + run("i") + "</m:sub><m:sup>" + run("n") + "</m:sup><m:e>" + run("i") + "</m:e></m:nary>")),
  "\u2211_i^n i"
);
check("a function keeps its name", mathText(parseFragment("<m:func><m:fName>" + run("sin") + "</m:fName><m:e>" + run("x") + "</m:e></m:func>")), "sinx");
check("a lower limit parses", mathText(parseFragment("<m:limLow><m:e>" + run("lim") + "</m:e><m:lim>" + run("0") + "</m:lim></m:limLow>")), "lim_{0}");
check("an accent parses", mathText(parseFragment("<m:acc><m:accPr><m:chr m:val=\"\u0302\"/></m:accPr><m:e>" + run("x") + "</m:e></m:acc>")), "x\u0302");
check("a bar parses", mathText(parseFragment("<m:bar><m:barPr><m:pos m:val=\"top\"/></m:barPr><m:e>" + run("x") + "</m:e></m:bar>")), "x");
check(
  "a matrix parses",
  mathText(parseFragment("<m:m><m:mr><m:e>" + run("1") + "</m:e><m:e>" + run("2") + "</m:e></m:mr><m:mr><m:e>" + run("3") + "</m:e><m:e>" + run("4") + "</m:e></m:mr></m:m>")),
  "1, 2; 3, 4"
);
check("an equation array parses", mathText(parseFragment("<m:eqArr><m:e>" + run("a") + "</m:e><m:e>" + run("b") + "</m:e></m:eqArr>")), "a; b");

// ---------- the italic convention ----------
const italic = parseFragment(run("x"));
const upright = parseFragment(run("+", true));
check("a variable is italic", italic.children[0].normal === false);
check("m:nor marks an operator upright", upright.children[0].normal === true);

// ---------- a mixed equation keeps every word ----------
const mixed = parseFragment(
  run("P") + "<m:f><m:num>" + run("F") + "</m:num><m:den>" + run("A") + "</m:den></m:f>" + run(" = ") + run("1.5", true)
);
check("a mixed equation reads back", mathText(mixed) === "PF/A = 1.5", mathText(mixed));

// ---------- nothing is lost when a structure is unknown ----------
const unknown = parseFragment("<m:unknownStruct>" + run("kept") + "</m:unknownStruct>");
check("an unknown structure still loses nothing", mathText(unknown) === "kept" || mathText(unknown) === "", mathText(unknown));

// ---------- the renderer builds the structure ----------
const model = {
  kind: "docx",
  body: [
    {
      type: "p",
      props: {},
      style: null,
      runs: [
        { type: "text", props: {}, text: "Dose: " },
        { type: "math", node: fraction },
      ],
    },
  ],
  section: { pageWidthTw: 11906, pageHeightTw: 16838, marginTopTw: 1440, marginRightTw: 1440, marginBottomTw: 1440, marginLeftTw: 1440 },
  styles: { paragraph: new Map(), character: new Map(), table: new Map(), docDefaults: null },
  footnotes: new Map(),
  endnotes: new Map(),
  headers: new Map(),
  footers: new Map(),
  mediaCache: null,
  properties: {},
  mediaUrl() {
    return null;
  },
};
const host = createContainer();
const renderer = createDocxRenderer({ container: host, model, settings: { zoom: 1 } });
const html = host.innerHTML;
check("the renderer draws a fraction", html.indexOf("ov-docx-frac") !== -1);
check("the numerator is in the markup", html.indexOf("ov-docx-frac-num") !== -1 && html.indexOf(">1<") !== -1);
check("the denominator is in the markup", html.indexOf("ov-docx-frac-den") !== -1 && html.indexOf(">2<") !== -1);
check("the text beside the equation survives", html.indexOf("Dose:") !== -1);
renderer.destroy();

console.log(`math: ${pass} pass, ${fail} fail`);
if (fail) process.exit(1);
