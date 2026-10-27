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

// Structure oracle: compares this viewer's parsed slide geometry against
// DeckCraft's, shape by shape. Pixels say something looks wrong; this says
// which shape is out of place and by how many points.
//
// DeckCraft has to be running with its control channel:
//
//   xvfb-run -a deckcraft --control 7979 deck.pptx &
//   bun test/oracle-structure.mjs deck.pptx --slides 1,5,12
//
// Then quit it with `deckcraft-cli app --method app.quit`.
//
// Boxes are slide points. Our parser keeps EMU, so 12700 EMU is one point.

import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { setupDom } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { parsePptx } = require("../src/pptx/parse.js");

const args = process.argv.slice(2);
const file = args.find((a) => !a.startsWith("--"));
function valueOf(flag, fallback) {
  const i = args.indexOf(flag);
  return i === -1 ? fallback : args[i + 1];
}
if (!file) {
  console.log("usage: bun test/oracle-structure.mjs <deck.pptx> [--slides 1,2] [--port 7979] [--tolerance 2]");
  process.exit(2);
}
const port = valueOf("--port", "7979");
const tolerance = Number(valueOf("--tolerance", "2"));
const only = valueOf("--slides", "");

// Commands run through engine.execute; control methods are called directly.
function deckcraft(method, params, command) {
  const argv = command
    ? ["app", "--port", port, "--method", "engine.execute", JSON.stringify({ command: method, params: params || {} })]
    : ["app", "--port", port, "--method", method, JSON.stringify(params || {})];
  const call = spawnSync("deckcraft-cli", argv, { encoding: "utf8" });
  if (call.status !== 0) return null;
  try {
    return JSON.parse(call.stdout);
  } catch (err) {
    return null;
  }
}

const probe = deckcraft("document.inspect", {});
if (!probe) {
  console.log("skip: no DeckCraft app on port " + port + " (start one with --control)");
  process.exit(0);
}

const EMU_PER_POINT = 12700;
const model = parsePptx(new Uint8Array(readFileSync(file)));
const indexed = probe.slides.map((s) => s.index);
const wanted = only
  ? only.split(",").map(Number)
  : model.slides.map((_, i) => i + 1);

function boxOf(shape) {
  return [shape.xEmu / EMU_PER_POINT, shape.yEmu / EMU_PER_POINT, shape.cxEmu / EMU_PER_POINT, shape.cyEmu / EMU_PER_POINT];
}

// Text as it should compare: whitespace collapsed, no bullet glyphs.
function clean(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function centre(box) {
  return [box[0] + box[2] / 2, box[1] + box[3] / 2];
}

console.log("deck", file.split("/").pop());
console.log("");
console.log("slide  ref/ours  matched  max Δpt  mean Δpt  text differ");

let totalShapes = 0;
let totalMatched = 0;
let totalTextDiff = 0;
let worst = { delta: 0 };

for (const slideNo of wanted) {
  if (!indexed.includes(slideNo - 1)) continue;
  const ref = deckcraft("slide.inspect", { index: slideNo - 1 }, true);
  if (!ref || !ref.shapes) continue;
  const ours = model.slides[slideNo - 1];
  const ourShapes = (ours ? ours.shapes : []).filter((s) => s && typeof s.xEmu === "number");

  const taken = new Set();
  let matched = 0;
  let maxDelta = 0;
  let sumDelta = 0;
  let textDiff = 0;

  for (const shape of ref.shapes) {
    if (!shape.box) continue;
    const [cx, cy] = centre(shape.box);
    let best = -1;
    let bestDist = Infinity;
    for (let i = 0; i < ourShapes.length; i++) {
      if (taken.has(i)) continue;
      const [ox, oy] = centre(boxOf(ourShapes[i]));
      const dist = Math.hypot(ox - cx, oy - cy);
      if (dist < bestDist) {
        bestDist = dist;
        best = i;
      }
    }
    // A shape further than a third of the slide away is not the same shape.
    if (best === -1 || bestDist > 320) continue;
    taken.add(best);
    matched++;
    const mine = boxOf(ourShapes[best]);
    const delta = Math.max(
      Math.abs(mine[0] - shape.box[0]),
      Math.abs(mine[1] - shape.box[1]),
      Math.abs(mine[2] - shape.box[2]),
      Math.abs(mine[3] - shape.box[3]),
    );
    sumDelta += delta;
    if (delta > maxDelta) maxDelta = delta;
    if (delta > worst.delta) worst = { delta, slide: slideNo, name: shape.name || shape.kind, mine, ref: shape.box };
    const a = clean(shape.text);
    const b = clean(ourShapes[best].text || "");
    if (a && a !== b) textDiff++;
  }

  const counted = ref.shapes.filter((s) => s.box).length;
  totalShapes += counted;
  totalMatched += matched;
  totalTextDiff += textDiff;
  const mean = matched ? sumDelta / matched : 0;
  const flag = maxDelta > tolerance || textDiff ? "  <-- review" : "";
  console.log(
    String(slideNo).padStart(5) + "  " + (counted + "/" + ourShapes.length).padStart(8) +
    "  " + String(matched).padStart(7) +
    "  " + maxDelta.toFixed(2).padStart(7) +
    "  " + mean.toFixed(2).padStart(8) +
    "  " + String(textDiff).padStart(10) + flag,
  );
}

console.log("");
console.log("shapes", totalMatched + "/" + totalShapes, "matched;", totalTextDiff, "text differences");
if (worst.delta) {
  console.log("worst placement: slide " + worst.slide + " " + worst.name + " off by " + worst.delta.toFixed(1) + " pt");
  console.log("  ours", worst.mine.map((n) => n.toFixed(1)).join(", "), " deckcraft", worst.ref.map((n) => n.toFixed(1)).join(", "));
}
