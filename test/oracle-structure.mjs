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
// No window is needed: `deckcraft-cli run` inspects a deck headlessly.
//
//   bun test/oracle-structure.mjs deck.pptx --slides 1,5,12
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
const tolerance = Number(valueOf("--tolerance", "2"));
const only = valueOf("--slides", "");
const quiet = args.includes("--quiet");
const verbose = args.includes("--verbose");

// One headless run inspects every slide we care about.
function inspectSlides(filePath, indices) {
  const argv = ["run", "--in", filePath, "--print"];
  for (const index of indices) argv.push("--cmd", 'slide.inspect={"index":' + index + "}");
  const call = spawnSync("deckcraft-cli", argv, { encoding: "utf8", timeout: 600000 });
  if (call.status !== 0) return null;
  const out = new Map();
  for (const line of String(call.stdout || "").split("\n")) {
    const at = line.indexOf("\u2192 ");
    if (at === -1) continue;
    try {
      const slide = JSON.parse(line.slice(at + 2));
      if (slide && typeof slide.index === "number") out.set(slide.index, slide);
    } catch (err) {
      // A line that is not a slide reply is not ours.
    }
  }
  return out;
}

const EMU_PER_POINT = 12700;
const model = parsePptx(new Uint8Array(readFileSync(file)));
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

const inspected = inspectSlides(file, wanted.map((n) => n - 1)) || new Map();
if (!inspected.size) {
  console.log("skip: DeckCraft could not inspect this deck");
  process.exit(0);
}

if (!quiet) {
  console.log("deck", file.split("/").pop());
  console.log("");
  console.log("slide  ref/ours  matched  max Δpt  mean Δpt  text differ");
}

let totalShapes = 0;
let totalMatched = 0;
let totalTextDiff = 0;
let worst = { delta: 0 };

for (const slideNo of wanted) {
  const ref = inspected.get(slideNo - 1);
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
    // The name is the reliable key: two shapes at the same spot are otherwise
    // easy to pair the wrong way round, which reads as a placement bug.
    if (shape.name) {
      const sameName = [];
      for (let i = 0; i < ourShapes.length; i++) {
        if (!taken.has(i) && ourShapes[i].name === shape.name) sameName.push(i);
      }
      if (sameName.length === 1) {
        best = sameName[0];
        const [ox, oy] = centre(boxOf(ourShapes[best]));
        bestDist = Math.hypot(ox - cx, oy - cy);
      }
    }
    if (best === -1) {
      for (let i = 0; i < ourShapes.length; i++) {
        if (taken.has(i)) continue;
        const [ox, oy] = centre(boxOf(ourShapes[i]));
        const dist = Math.hypot(ox - cx, oy - cy);
        if (dist < bestDist) {
          bestDist = dist;
          best = i;
        }
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
    if (a && a !== b) {
      textDiff++;
      if (verbose && textDiff <= 12) {
        console.log("  text differs: slide " + slideNo + " " + (shape.name || shape.kind));
        console.log("    deckcraft: " + a.slice(0, 100));
        console.log("    ours     : " + b.slice(0, 100));
      }
    }
  }

  const counted = ref.shapes.filter((s) => s.box).length;
  totalShapes += counted;
  totalMatched += matched;
  totalTextDiff += textDiff;
  const mean = matched ? sumDelta / matched : 0;
  const flag = maxDelta > tolerance || textDiff ? "  <-- review" : "";
  if (quiet) continue;
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
