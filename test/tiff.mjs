/*
 * The TIFF predictor restarts at the first pixel of every row. Running the
 * horizontal differencing across the whole strip instead adds the previous
 * row's last pixel to the next row's first one, and the error walks through the
 * image: colours drift and thin text goes rough. This builds tiny files byte by
 * byte and reads the decoded PNG back.
 *
 *   bun test/tiff.mjs
 */

import { createRequire } from "node:module";
import { setupDom } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { unzlibSync } = require("fflate");
const { tiffToDataUrl } = require("../src/media/media.js");

let pass = 0;
let fail = 0;
function check(name, condition, detail) {
  if (condition) { pass++; console.log("ok   " + name); }
  else { fail++; console.log("FAIL " + name + (detail !== undefined ? "  -> " + detail : "")); }
}

// Little-endian baseline TIFF, one uncompressed strip, predictor 2.
function buildTiff(width, height, samples, rawRows) {
  const rowBytes = width * samples;
  const data = new Uint8Array(rowBytes * height);
  for (let y = 0; y < height; y++) {
    const row = rawRows[y];
    for (let i = 0; i < samples; i++) data[y * rowBytes + i] = row[i];
    for (let i = samples; i < rowBytes; i++) data[y * rowBytes + i] = (row[i] - row[i - samples]) & 255;
  }
  const entries = [
    [256, 3, 1, width],
    [257, 3, 1, height],
    [258, 3, samples, samples <= 2 ? [8, 8].slice(0, samples) : "bits"],
    [259, 3, 1, 1],
    [262, 3, 1, 2],
    [273, 4, 1, "strip"],
    [277, 3, 1, samples],
    [278, 4, 1, height],
    [279, 4, 1, data.length],
    [284, 3, 1, 1],
    [317, 3, 1, 2],
  ];
  const ifd = 8;
  const dataStart = ifd + 2 + entries.length * 12 + 4;
  const bitsOffset = dataStart;
  const stripOffset = dataStart + (samples > 2 ? samples * 2 : 0);
  const total = stripOffset + data.length;
  const buffer = new Uint8Array(total);
  const view = new DataView(buffer.buffer);
  view.setUint16(0, 0x4949, true);
  view.setUint16(2, 42, true);
  view.setUint32(4, ifd, true);
  view.setUint16(ifd, entries.length, true);
  entries.forEach(([tag, type, count, value], index) => {
    const at = ifd + 2 + index * 12;
    view.setUint16(at, tag, true);
    view.setUint16(at + 2, type, true);
    view.setUint32(at + 4, count, true);
    if (value === "bits") {
      view.setUint32(at + 8, bitsOffset, true);
      for (let i = 0; i < samples; i++) view.setUint16(bitsOffset + i * 2, 8, true);
    } else if (value === "strip") {
      view.setUint32(at + 8, stripOffset, true);
    } else if (Array.isArray(value)) {
      value.forEach((v, i) => view.setUint16(at + 8 + i * 2, v, true));
    } else if (type === 3) {
      view.setUint16(at + 8, value, true);
    } else {
      view.setUint32(at + 8, value, true);
    }
  });
  buffer.set(data, stripOffset);
  return buffer;
}

// Our PNGs are RGBA, filter 0 on every row, so the read back is short.
function readPng(dataUrl) {
  const buffer = Buffer.from(dataUrl.slice(dataUrl.indexOf(",") + 1), "base64");
  let offset = 8;
  let width = 0;
  let height = 0;
  const idat = [];
  while (offset + 12 <= buffer.length) {
    const length = buffer.readUInt32BE(offset);
    const type = buffer.toString("ascii", offset + 4, offset + 8);
    const body = buffer.subarray(offset + 8, offset + 8 + length);
    if (type === "IHDR") {
      width = body.readUInt32BE(0);
      height = body.readUInt32BE(4);
    } else if (type === "IDAT") {
      idat.push(body);
    }
    offset += 12 + length;
  }
  const raw = unzlibSync(Buffer.concat(idat));
  const pixels = [];
  const stride = width * 4;
  for (let y = 0; y < height; y++) {
    const start = y * (stride + 1) + 1;
    for (let x = 0; x < stride; x++) pixels.push(raw[start + x]);
  }
  return { width, height, pixels };
}

const pixel = (png, x, y) => png.pixels.slice((y * png.width + x) * 4, (y * png.width + x) * 4 + 4).join(",");

// Three samples, two rows: the second row's first pixel must stay absolute.
const rgb = buildTiff(2, 2, 3, [
  [10, 20, 30, 40, 50, 60],
  [70, 80, 90, 100, 110, 120],
]);
const rgbPng = readPng(tiffToDataUrl(new Uint8Array(rgb)));
check("the size survives", rgbPng.width === 2 && rgbPng.height === 2, rgbPng.width + "x" + rgbPng.height);
check("the first pixel reads as stored", pixel(rgbPng, 0, 0) === "10,20,30,255", pixel(rgbPng, 0, 0));
check("the second pixel of the row differs", pixel(rgbPng, 1, 0) === "40,50,60,255", pixel(rgbPng, 1, 0));
check("the new row's first pixel is absolute", pixel(rgbPng, 0, 1) === "70,80,90,255", pixel(rgbPng, 0, 1));
check("the new row's second pixel is right", pixel(rgbPng, 1, 1) === "100,110,120,255", pixel(rgbPng, 1, 1));

// Four samples: alpha is the extra sample and goes through the predictor too.
const rgba = buildTiff(2, 2, 4, [
  [10, 20, 30, 200, 40, 50, 60, 210],
  [70, 80, 90, 220, 100, 110, 120, 230],
]);
const rgbaPng = readPng(tiffToDataUrl(new Uint8Array(rgba)));
check("alpha rides the predictor", pixel(rgbaPng, 0, 0) === "10,20,30,200", pixel(rgbaPng, 0, 0));
check("alpha on the second row is absolute", pixel(rgbaPng, 0, 1) === "70,80,90,220", pixel(rgbaPng, 0, 1));
check("alpha on the last pixel is right", pixel(rgbaPng, 1, 1) === "100,110,120,230", pixel(rgbaPng, 1, 1));

console.log("");
console.log("tiff:", pass, "pass,", fail, "fail");
process.exit(fail ? 1 : 0);
