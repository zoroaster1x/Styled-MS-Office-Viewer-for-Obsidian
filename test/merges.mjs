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
 */

// A merge that spans a whole column or row is formatting, not content.
//
//   bun test/merges.mjs
//
// Excel writes a spacer column as a merge like D1:D1048576. Letting that grow
// the used range made border resolution walk a million rows and a small
// timetable took nine seconds to draw. The dimension is the authority for the
// last row and column; the merge keeps its stated range in the model, and the
// renderer clamps it to the used range when it builds the cover map.

import { createRequire } from "node:module";
import { setupDom } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { parseSheet } = require("../src/spreadsheet/sheet.js");

let pass = 0;
let fail = 0;
function check(name, condition, detail) {
	if (condition) {
		pass++;
		console.log("ok   " + name);
	} else {
		fail++;
		console.log("FAIL " + name + (detail !== undefined ? "  -> " + detail : ""));
	}
}

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

function sheetWith(mergeRef) {
	const xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet ${NS}>
  <dimension ref="A1:B2"/>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Week</t></is></c></row></sheetData>
  <mergeCells count="1"><mergeCell ref="${mergeRef}"/></mergeCells>
</worksheet>`;
	return parseSheet(xml, { name: "Merge", sharedStrings: [], rels: new Map() });
}

const columnMerge = sheetWith("B1:B1048576");
check("whole-column merge keeps the stated last row", columnMerge.dims.r2 === 2, columnMerge.dims.r2);
check("whole-column merge keeps its own range", columnMerge.merges[0].r2 === 1048576, columnMerge.merges[0].r2);
check("content is still parsed", columnMerge.cells.size === 1, columnMerge.cells.size);

const rowMerge = sheetWith("A1:XFD1");
check("whole-row merge keeps the stated last column", rowMerge.dims.c2 === 2, rowMerge.dims.c2);
check("whole-row merge keeps its own range", rowMerge.merges[0].c2 === 16384, rowMerge.merges[0].c2);

const insideMerge = sheetWith("A1:B2");
check("a merge inside the dimension changes nothing", insideMerge.dims.r2 === 2 && insideMerge.dims.c2 === 2);

console.log("");
console.log("merges:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
