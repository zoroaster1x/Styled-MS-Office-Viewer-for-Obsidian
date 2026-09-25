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

// Number format checks. The date cases pin Excel's 1900 leap year bug, which is
// what decides whether serial 1 is 1900-01-01 or 1899-12-31 in a spreadsheet the
// user opens.

import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const { formatValue } = require("../src/spreadsheet/numfmt.js");

let pass = 0;
let fail = 0;
function check(name, got, want) {
  if (got === want) {
    pass++;
    return;
  }
  fail++;
  console.log("FAIL", name, "->", JSON.stringify(got), "wanted", JSON.stringify(want));
}

function text(value, fmt, opts) {
  return formatValue(value, fmt, opts || {}).text;
}

// 1900 system, the leap bug.
check("serial 1", text(1, "yyyy-mm-dd"), "1900-01-01");
check("serial 2", text(2, "yyyy-mm-dd"), "1900-01-02");
check("serial 59", text(59, "yyyy-mm-dd"), "1900-02-28");
check("serial 60 (the invented day)", text(60, "yyyy-mm-dd"), "1900-02-29");
check("serial 61", text(61, "yyyy-mm-dd"), "1900-03-01");
check("serial 61 with day names", text(61, "ddd"), "Thu");
check("serial 45000", text(45000, "yyyy-mm-dd"), "2023-03-15");

// 1904 system, no leap bug.
check("1904 serial 0", text(0, "yyyy-mm-dd", { date1904: true }), "1904-01-01");
check("1904 serial 1", text(1, "yyyy-mm-dd", { date1904: true }), "1904-01-02");

// Times and the m month/minute split.
check("half day", text(0.5, "hh:mm:ss"), "12:00:00");
check("quarter hour", text(0.25, "hh:mm"), "06:00");
check("minute token", text(0.5, "hh:mm"), "12:00");
check("elapsed hours", text(1.5, "[h]:mm"), "36:00");

// Numbers, sections and literals.
check("general integer", text(42, "General"), "42");
check("thousands separator", text(1234567, "#,##0"), "1,234,567");
check("decimals", text(3.14159, "0.00"), "3.14");
check("percent", text(0.25, "0%"), "25%");
check("currency literal", text(12.5, '"$"#,##0.00'), "$12.50");
check("negative section", text(-5, "0;[Red]-0"), "-5");
check("negative zero section", text(-5, "0;(0)"), "(5)");
check("text section", text("note", "0;0;0;@"), "note");
check("scientific", text(12345, "0.00E+00"), "1.23E+04");
check("fraction halves", text(1.5, "# ?/?"), "1 1/2");
check("accounting spaces", text(-3.5, '$#,##0.00_);[Red]($#,##0.00)'), "($3.50)");

console.log(`numfmt: ${pass} pass, ${fail} fail`);
if (fail) process.exit(1);
