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
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
 * FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
 * details.
 *
 * You should have received a copy of the GNU General Public License along with
 * this program. If not, see <https://www.gnu.org/licenses/>.
 */

// Development oracle: draws the same deck twice, once with DeckCraft (a clean
// room PowerPoint reimplementation in Rust) and once with this viewer, then
// reports a per-slide difference and writes a side by side contact sheet.
//
//   bun test/oracle.mjs <deck.pptx> [--slides 1,2,3] [--out /tmp/oracle]
//
// DeckCraft supplies the reference render through `deckcraft-cli render`, this
// viewer renders to a bare HTML page through test/render-html.mjs, and a
// headless Firefox screenshots it. ImageMagick compares the two.
//
// Nothing here runs in CI: it needs deckcraft-cli, a Firefox binary and
// ImageMagick, none of which the runner has. Missing tools skip with a notice.
//
//   DECKCRAFT_BIN   deckcraft-cli (default: on PATH)
//   FIREFOX_BIN     a Firefox that supports --headless --screenshot
//   MAGICK_BIN      magick or convert (default: on PATH)

import { execFileSync, spawnSync } from "node:child_process";
import { mkdirSync, writeFileSync, existsSync, readdirSync } from "node:fs";
import { homedir } from "node:os";
import { basename, join } from "node:path";

const args = process.argv.slice(2);
const file = args.find((a) => !a.startsWith("--"));
function valueOf(flag, fallback) {
  const i = args.indexOf(flag);
  return i === -1 ? fallback : args[i + 1];
}
if (!file || !existsSync(file)) {
  console.log("usage: bun test/oracle.mjs <deck.pptx> [--slides 1,2,3] [--out DIR]");
  process.exit(2);
}
const outDir = valueOf("--out", "/tmp/oracle");
const only = valueOf("--slides", "");
mkdirSync(outDir, { recursive: true });
mkdirSync(join(outDir, "ref"), { recursive: true });
mkdirSync(join(outDir, "ours"), { recursive: true });
mkdirSync(join(outDir, "diff"), { recursive: true });

function have(binary, args2) {
  if (!binary) return false;
  const probe = spawnSync(binary, args2 || ["--version"], { encoding: "utf8" });
  return probe.status === 0;
}

const deckcraft = process.env.DECKCRAFT_BIN || "deckcraft-cli";
const magick = process.env.MAGICK_BIN || "magick";

// Firefox is whatever the environment names, then a firefox on PATH, then the
// newest playwright build under this user's cache. No path of anybody's
// machine belongs in a public repository.
function findFirefox() {
  if (process.env.FIREFOX_BIN) return process.env.FIREFOX_BIN;
  const onPath = spawnSync("sh", ["-c", "command -v firefox"], { encoding: "utf8" });
  if (onPath.status === 0 && (onPath.stdout || "").trim()) return onPath.stdout.trim();
  const cache = join(homedir(), ".cache", "invisible-playwright");
  if (!existsSync(cache)) return "";
  const builds = readdirSync(cache)
    .filter((name) => name.startsWith("firefox-"))
    .sort()
    .reverse();
  for (const build of builds) {
    const candidate = join(cache, build, "firefox");
    if (existsSync(candidate)) return candidate;
  }
  return "";
}
const firefox = findFirefox();

if (!have(deckcraft, ["--version"])) {
  console.log("skip: deckcraft-cli is not installed (set DECKCRAFT_BIN)");
  process.exit(0);
}
if (!have(magick, ["-version"])) {
  console.log("skip: ImageMagick is not installed (set MAGICK_BIN)");
  process.exit(0);
}
if (!have(firefox, ["--version"])) {
  console.log("skip: no Firefox with headless screenshots (set FIREFOX_BIN)");
  process.exit(0);
}

// This viewer draws a slide at 96 dpi, DeckCraft at 72 points per inch, so the
// reference has to come out at 4/3 scale for the two images to line up.
const SCALE = 4 / 3;
const WIDTH = Math.round(960 * SCALE);
const HEIGHT = Math.round(540 * SCALE);

const name = basename(file);
const root = join(import.meta.dir, "..");
console.log("deck", name);

// ---------- the reference ----------
let refFiles = [];
try {
  execFileSync(deckcraft, ["render", file, "--all", "--scale", String(SCALE), join(outDir, "ref")], { stdio: "pipe" });
  refFiles = readdirSync(join(outDir, "ref")).filter((f) => f.endsWith(".png")).sort();
} catch (err) {
  console.log("deckcraft could not render this deck:", String(err.stderr || err.message).trim().slice(0, 200));
  process.exit(1);
}
if (!refFiles.length) {
  console.log("deckcraft produced no slides");
  process.exit(1);
}

const wanted = only
  ? only.split(",").map((n) => Number(n)).filter((n) => n >= 1 && n <= refFiles.length)
  : refFiles.map((_, i) => i + 1);

// ---------- our render, one slide at a time ----------
const rows = [];
for (const slide of wanted) {
  const html = join(outDir, "ours", "slide-" + String(slide).padStart(2, "0") + ".html");
  const png = join(outDir, "ours", "slide-" + String(slide).padStart(2, "0") + ".png");
  const render = spawnSync("bun", [join(root, "test", "render-html.mjs"), file, "--slide", String(slide), "--bare", "--out", html], { encoding: "utf8", cwd: root, timeout: 300000 });
  if (!existsSync(html)) {
    rows.push({ slide, note: "render failed: " + String(render.stderr || "").trim().split("\n").pop() });
    continue;
  }
  const shot = spawnSync(firefox, ["--headless", "--screenshot", png, "--window-size=" + WIDTH + "," + HEIGHT, "file://" + html], { encoding: "utf8", timeout: 180000 });
  if (!existsSync(png)) {
    rows.push({ slide, note: "screenshot failed: " + String(shot.stderr || "").trim().split("\n").pop() });
    continue;
  }
  const diff = join(outDir, "diff", "slide-" + String(slide).padStart(2, "0") + ".png");
  const compare = spawnSync(magick, ["compare", "-metric", "RMSE", join(outDir, "ref", refFiles[slide - 1]), png, diff], { encoding: "utf8" });
  const metric = (compare.stderr || "").trim().split(" ")[0];
  const parts = /\(([\d.]+)\)/.exec(compare.stderr || "");
  const rmse = parts ? Number(parts[1]) : null;
  rows.push({ slide, rmse, metric, diff: existsSync(diff) ? diff : null });
}

// ---------- the contact sheet, reference above ours ----------
const sheet = [];
for (const row of rows) {
  if (!row.diff) continue;
  const ref = join(outDir, "ref", refFiles[row.slide - 1]);
  const mine = join(outDir, "ours", "slide-" + String(row.slide).padStart(2, "0") + ".png");
  sheet.push(ref, mine);
}
if (sheet.length) {
  try {
    execFileSync(magick, ["montage", ...sheet, "-tile", "2x", "-geometry", "+2+2", "-background", "#202020", join(outDir, "contact-sheet.png")]);
  } catch (err) {
    console.log("contact sheet failed:", String(err.message).slice(0, 120));
  }
}

console.log("");
console.log("slide   RMSE    reference / ours");
for (const row of rows) {
  if (row.note) {
    console.log(String(row.slide).padStart(5) + "   " + row.note);
    continue;
  }
  const pct = row.rmse == null ? "?" : (row.rmse * 100).toFixed(1) + "%";
  console.log(String(row.slide).padStart(5) + "   " + pct.padStart(6) + "  ref " + refFiles[row.slide - 1] + "  diff " + (row.diff ? basename(row.diff) : "none"));
}

const scored = rows.filter((r) => r.rmse != null);
const mean = scored.length ? scored.reduce((a, b) => a + b.rmse, 0) / scored.length : null;
console.log("");
console.log("slides compared", scored.length, mean == null ? "" : "mean RMSE " + (mean * 100).toFixed(1) + "%");
console.log("images under", outDir, "(ref/, ours/, diff/, contact-sheet.png)");

writeFileSync(join(outDir, "report.json"), JSON.stringify({ deck: name, width: WIDTH, height: HEIGHT, rows }, null, 1));
