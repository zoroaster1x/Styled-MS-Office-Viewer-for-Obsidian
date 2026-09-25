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

// Privacy guard. Fails when a tracked file mentions a personal path, a home
// directory, or anything you add to the local list. It keeps the repository
// free of machine details, which matters when the tests are asked to run
// against real documents.
//
//   bun test/privacy.mjs
//
// The general patterns below cover paths that belong to a specific machine.
// Anything more specific (your own folder names, a course code, a document
// prefix) goes in the gitignored .testenv as OV_PRIVACY_PATTERNS, a comma
// separated list of regular expressions, so the repository never holds the
// words it is meant to keep out.

import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { testEnv } from "./env.mjs";

const TRACKED = execFileSync("git", ["ls-files"], { encoding: "utf8" })
  .split("\n")
  .map((line) => line.trim())
  .filter(Boolean);

const GENERAL = [
  { name: "absolute home path", re: /\/home\/(?!you\/|user\/)[a-z0-9._-]+\//i },
  { name: "macOS user path", re: /\/Users\/[a-z0-9._-]+\//i },
  { name: "Windows user path", re: /[A-Z]:\\Users\\/i },
  { name: "private email address", re: /[a-z0-9._%+-]+@(?!example\.|test\.|localhost)[a-z0-9.-]+\.[a-z]{2,}/i },
  { name: "machine hostname", re: /\b[\w-]+\.(local|lan|home|internal)\b/i },
];

// Extra patterns from the local, gitignored configuration.
const LOCAL = (testEnv().OV_PRIVACY_PATTERNS || "")
  .split(",")
  .map((piece) => piece.trim())
  .filter(Boolean)
  .map((piece) => {
    try {
      return { name: "local pattern", re: new RegExp(piece, "i") };
    } catch (err) {
      console.error(`FAIL .testenv OV_PRIVACY_PATTERNS has an invalid expression: ${piece}`);
      process.exit(1);
      return null;
    }
  })
  .filter(Boolean);

// Files whose job is to write the patterns down.
const ALLOWED = new Set([
  "test/privacy.mjs",
  ".testenv.example",
  ".gitignore",
  "LICENSE",
]);

let failed = 0;
let checked = 0;
let skipped = 0;
for (const file of TRACKED) {
  if (ALLOWED.has(file)) continue;
  if (file.split("/").some((part) => part === "node_modules" || part === "__pycache__")) {
    skipped++;
    continue;
  }
  let text;
  try {
    text = readFileSync(file, "utf8");
  } catch (err) {
    continue;
  }
  if (text.indexOf("\u0000") !== -1) continue;
  checked++;
  for (const rule of GENERAL.concat(LOCAL)) {
    if (rule.re.test(text)) {
      failed++;
      console.error(`FAIL ${file}: looks like a ${rule.name}`);
    }
  }
}

console.log(`privacy: checked ${checked} tracked files (${skipped} skipped), ${GENERAL.length + LOCAL.length} patterns, ${failed} violations`);
if (!LOCAL.length) {
  console.log("        add your own terms to OV_PRIVACY_PATTERNS in .testenv to extend this list");
}
process.exit(failed ? 1 : 0);
