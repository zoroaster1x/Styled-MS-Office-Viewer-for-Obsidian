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

// Local test configuration.
//
// The tests that read real documents take their paths from a gitignored
// `.testenv` file at the repository root, so no personal path or document name
// ever lands in the repository. Copy `.testenv.example` to `.testenv` and fill
// in your own paths. Anything you leave out is skipped, with a notice, rather
// than failed.
//
// The synthetic tests (smoke, filters, copy) need no configuration and always
// run.

import { existsSync, readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");

const KEYS = {
  OV_TEST_XLSX: "a workbook (.xlsx) to try the grid against",
  OV_TEST_DOCX: "a document (.docx) to try the page renderer against",
  OV_TEST_PPTX: "a presentation (.pptx) to try the slide renderer against",
  OV_TEST_FOLDER: "a folder of course documents to sweep with the stress and fidelity runs",
};

let loaded = null;

export function testEnv() {
  if (loaded) return loaded;
  const env = { ...process.env };
  const path = join(ROOT, ".testenv");
  if (existsSync(path)) {
    for (const line of readFileSync(path, "utf8").split(/\r?\n/)) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#")) continue;
      const at = trimmed.indexOf("=");
      if (at === -1) continue;
      const key = trimmed.slice(0, at).trim();
      let value = trimmed.slice(at + 1).trim();
      if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
        value = value.slice(1, -1);
      }
      if (value) env[key] = value;
    }
  }
  loaded = env;
  return env;
}

// Returns the path for a key, or null when it is not configured. A configured
// path that does not exist is a hard error, so a typo in .testenv is noticed.
export function testPath(key) {
  const value = testEnv()[key];
  if (!value) return null;
  if (!existsSync(value)) {
    console.error(`FAIL ${key} points at ${value}, which does not exist. Fix .testenv.`);
    process.exit(1);
  }
  return value;
}

// Prints why a group of checks is being skipped, once per key.
const announced = new Set();
export function skipNotice(key, label) {
  if (announced.has(key)) return;
  announced.add(key);
  console.log(`skip ${label}: set ${key} in .testenv to run this (see .testenv.example)`);
  console.log(`     ${KEYS[key] || ""}`);
}

// Prints the configuration state at the top of a run, so it is obvious which
// checks are running and which are waiting on .testenv.
export function reportConfiguration(keys) {
  const configured = keys.filter((key) => Boolean(testEnv()[key]));
  if (configured.length === keys.length) return;
  console.log("Local documents are not configured yet; the real-file checks are skipped.");
  console.log("Copy .testenv.example to .testenv and fill in the paths to run them.\n");
}
