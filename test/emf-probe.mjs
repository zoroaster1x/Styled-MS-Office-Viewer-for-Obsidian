// Records what the EMF player asks the canvas to do, so a vector logo can be
// checked without a rasteriser: fills, their colours, their subpath counts and
// their bounds, plus clips and transforms.
import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { setupDom } from "./harness.mjs";
setupDom();
const require = createRequire(import.meta.url);
const { Package } = require("../src/shared/package.js");

const calls = [];
const state = { fillStyle: "#000000", strokeStyle: "#000000", lineWidth: 1, transform: [1, 0, 0, 1, 0, 0] };
const stack = [];
function record(kind, arg) { calls.push({ kind, arg: arg === undefined ? null : arg, fill: state.fillStyle, stroke: state.strokeStyle }); }

const ctx = {
  get fillStyle() { return state.fillStyle; },
  set fillStyle(v) { state.fillStyle = v; },
  get strokeStyle() { return state.strokeStyle; },
  set strokeStyle(v) { state.strokeStyle = v; },
  get lineWidth() { return state.lineWidth; },
  set lineWidth(v) { state.lineWidth = v; },
  lineJoin: "round", lineCap: "butt", textBaseline: "alphabetic", fillRule: "nonzero",
  save() { stack.push(Object.assign({}, state)); },
  restore() { const s = stack.pop(); if (s) Object.assign(state, s); },
  translate() {}, scale() {}, rotate() {}, transform() {}, setTransform() {},
  beginPath() { state.path = []; record("beginPath"); },
  moveTo(x, y) { state.path.push({ t: "M", x, y }); },
  lineTo(x, y) { state.path.push({ t: "L", x, y }); },
  closePath() { state.path.push({ t: "Z" }); },
  quadraticCurveTo(x1, y1, x2, y2) { state.path.push({ t: "Q", x1, y1, x2, y2 }); },
  bezierCurveTo(...a) { state.path.push({ t: "C", a }); },
  rect(x, y, w, h) { state.path.push({ t: "rect", x, y, w, h }); record("rect", { x, y, w, h }); },
  ellipse(x, y, rx, ry) { state.path.push({ t: "ellipse", x, y, rx, ry }); },
  arc() {}, clip() { record("clip", (state.path || []).length); },
  fill(rule) {
    const points = (state.path || []).filter((p) => p.x !== undefined);
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    for (const pt of points) {
      if (pt.x < minX) minX = pt.x;
      if (pt.y < minY) minY = pt.y;
      if (pt.x > maxX) maxX = pt.x;
      if (pt.y > maxY) maxY = pt.y;
    }
    record("fill", {
      rule: rule || state.fillRule,
      subpaths: (state.path || []).filter((p) => p.t === "M").length,
      box: points.length ? [Math.round(minX), Math.round(minY), Math.round(maxX - minX), Math.round(maxY - minY)] : null,
    });
  },
  stroke() { record("stroke", { subpaths: (state.path || []).filter((p) => p.t === "M").length, width: state.lineWidth }); },
  fillRect() {}, strokeRect() {}, drawImage() {},
  setLineDash() {}, measureText() { return { width: 10 }; }, fillText() { record("fillText"); }, putImageData() {},
  createImageData(w, h) { return { width: w, height: h, data: new Uint8ClampedArray(w * h * 4) }; },
};
const canvas = {
  width: 0, height: 0,
  getContext() { return ctx; },
  toDataURL() { return "data:image/png;base64,probe"; },
};

const originalCreate = globalThis.document.createElement.bind(globalThis.document);
globalThis.document.createElement = function (tag) {
  if (String(tag).toLowerCase() === "canvas") return canvas;
  return originalCreate(tag);
};

const { parseDocx } = require("../src/docx/parse.js");
const model = parseDocx(new Uint8Array(readFileSync(process.argv[2])));
const rel = [...model.rels.entries()].find(([, r]) => model.pkg.resolve(model.partPath, r) === "word/media/image1.emf");
console.log("rel for the emf:", rel ? rel[0] : "none");
const url = rel ? model.mediaUrl(rel[0]) : null;
console.log("transcoded:", url ? url.slice(0, 24) : "FAILED");
const fills = calls.filter((c) => c.kind === "fill");
console.log("fills:", fills.length);
const byColour = new Map();
for (const f of fills) {
  const key = f.fill + "|" + (f.arg ? f.arg.rule : "");
  const entry = byColour.get(key) || { n: 0, subpaths: [] };
  entry.n++;
  entry.subpaths.push(f.arg ? f.arg.subpaths : 0);
  byColour.set(key, entry);
}
for (const [key, entry] of byColour) {
  console.log("   colour/rule", key, "x" + entry.n, "subpaths:", entry.subpaths.join(","));
}
console.log("fill boxes (canvas units):");
for (const f of fills) console.log("   ", f.fill, "box", f.arg ? JSON.stringify(f.arg.box) : "-", "subpaths", f.arg ? f.arg.subpaths : "-");
const strokes = calls.filter((c) => c.kind === "stroke");
const strokeColours = new Map();
for (const st of strokes) strokeColours.set(st.stroke + " w" + (st.arg ? st.arg.width : "?"), (strokeColours.get(st.stroke + " w" + (st.arg ? st.arg.width : "?")) || 0) + 1);
console.log("clips:", calls.filter((c) => c.kind === "clip").length, "strokes:", strokes.length);
for (const [key, n] of strokeColours) console.log("   stroke", key, "x" + n);
