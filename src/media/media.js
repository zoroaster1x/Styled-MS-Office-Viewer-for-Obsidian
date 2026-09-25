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

"use strict";

// Turns embedded media into something an <img> or <video> can display, without
// writing anything to disk. Chromium decodes PNG, JPEG, GIF, WebP, BMP and SVG
// directly. TIFF and the EMF/WMF metafiles it cannot, so they are transcoded
// here: TIFF through a baseline decoder, metafiles through a record renderer.
// JPEG XR goes through the WebAssembly codec in ./jxr.

const { inflateSync, zlibSync } = require("fflate");
const { decodeJxrRgba } = require("./jxr");

// Results are cached per part path, so a picture reused on ten slides is
// resolved once. JPEG and PNG bytes become blob URLs, which keeps the DOM
// small: a slide shown ten times costs one object URL, not ten base64 copies.
// Formats Chromium cannot decode are transcoded to a data URL instead.
function createMediaCache() {
  const cache = new Map();
  const objectUrls = new Set();
  return {
    get(key) {
      return cache.has(key) ? cache.get(key) : undefined;
    },
    set(key, value) {
      if (value && value.url && value.objectUrl) objectUrls.add(value.url);
      cache.set(key, value);
      return value;
    },
    // Resolve through the cache: returns a string URL or null.
    resolve(key, factory) {
      if (cache.has(key)) return cache.get(key);
      const value = factory();
      this.set(key, value);
      return value;
    },
    release() {
      for (const url of objectUrls) {
        try {
          URL.revokeObjectURL(url);
        } catch (err) {
          // Already revoked.
        }
      }
      objectUrls.clear();
      cache.clear();
    },
  };
}

function bytesToBase64(bytes) {
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

function bytesToDataUrl(bytes, mime) {
  return "data:" + mime + ";base64," + bytesToBase64(bytes);
}

// Resolves one package part to a displayable URL. Returns null when the part
// cannot be drawn; callers show their own placeholder.
function mediaUrl(pkg, path, cache) {
  if (!pkg || !path) return null;
  if (cache && typeof cache.resolve === "function") {
    return cache.resolve(path, () => buildMediaUrl(pkg, path));
  }
  if (cache) {
    const hit = cache.get(path);
    if (hit !== undefined) return hit;
  }
  const out = buildMediaUrl(pkg, path);
  if (cache && typeof cache.set === "function") cache.set(path, out);
  return out;
}

function buildMediaUrl(pkg, path) {
  const bytes = pkg.bytes(path);
  if (!bytes) return null;
  const mime = pkg.mimeOf(path);
  try {
    if (mime === "image/tiff" || mime === "image/tif") return tiffToDataUrl(bytes);
    // Chromium has no JPEG XR decoder, so the WebAssembly codec draws it. The
    // module is warmed before the parse; when it is not ready (or the host
    // refuses WebAssembly) null shows the caller's named placeholder instead of
    // a broken image. JPEG XL is still undecodable here.
    if (/^image\/(vnd\.ms-photo|jxr)$/i.test(mime) || /\.(wdp|hdp|jxr)$/i.test(path)) {
      const image = decodeJxrRgba(bytes);
      return image ? pngDataUrl(image.data, image.width, image.height) : null;
    }
    if (/^image\/jxl$/i.test(mime) || /\.jxl$/i.test(path)) return null;
    if (mime === "image/emf" || /\.emf$/i.test(path)) return metafileToDataUrl(bytes, "emf");
    if (mime === "image/wmf" || /\.wmf$/i.test(path)) return metafileToDataUrl(bytes, "wmf");
    if (mime === "image/svg+xml") {
      return "data:image/svg+xml;charset=utf-8," + encodeURIComponent(new TextDecoder("utf-8").decode(bytes));
    }
    if (mime.indexOf("image/") === 0 || mime.indexOf("audio/") === 0 || mime.indexOf("video/") === 0) {
      return blobUrl(bytes, mime);
    }
  } catch (err) {
    return null;
  }
  return null;
}

// Blob URLs let Chromium decode the pixels lazily and keep the markup small.
function blobUrl(bytes, mime) {
  if (typeof URL === "undefined" || typeof URL.createObjectURL !== "function") {
    return bytesToDataUrl(bytes, mime);
  }
  try {
    return URL.createObjectURL(new Blob([bytes], { type: mime }));
  } catch (err) {
    return bytesToDataUrl(bytes, mime);
  }
}

// ---------- TIFF ----------

// Baseline TIFF plus LZW, PackBits and Deflate compression. The decks in this
// vault store photographs and exported figures as TIFF, which Chromium will not
// decode, so the common baseline variants are covered here.
const TIFF_TAGS = {
  width: 256,
  height: 257,
  bitsPerSample: 258,
  compression: 259,
  photometric: 262,
  fillOrder: 266,
  stripOffsets: 273,
  orientation: 274,
  samplesPerPixel: 277,
  rowsPerStrip: 278,
  stripByteCounts: 279,
  planarConfig: 284,
  predictor: 317,
  colorMap: 320,
  extraSamples: 338,
};

function tiffToDataUrl(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.byteLength < 8) return null;
  const magic = view.getUint16(0, false);
  let le;
  if (magic === 0x4949) le = true;
  else if (magic === 0x4d4d) le = false;
  else return null;
  if (view.getUint16(2, le) !== 42) return null;
  const ifdOffset = view.getUint32(4, le);
  const ifd = readTiffIfd(view, ifdOffset, le, bytes);
  if (!ifd) return null;
  const width = ifd.number(TIFF_TAGS.width);
  const height = ifd.number(TIFF_TAGS.height);
  // A TIFF header is a handful of numbers, and a wrong one turns into a
  // multi gigabyte allocation for four bytes of file. Bound the canvas.
  if (!width || !height || width > 100000 || height > 100000 || width * height > 50000000) return null;
  const bits = ifd.numbers(TIFF_TAGS.bitsPerSample) || [8];
  const compression = ifd.number(TIFF_TAGS.compression) || 1;
  const photometric = ifd.number(TIFF_TAGS.photometric);
  const samples = ifd.number(TIFF_TAGS.samplesPerPixel) || 1;
  const predictor = ifd.number(TIFF_TAGS.predictor) || 1;
  const planar = ifd.number(TIFF_TAGS.planarConfig) || 1;
  const rowsPerStrip = ifd.number(TIFF_TAGS.rowsPerStrip) || height;
  const stripOffsets = ifd.numbers(TIFF_TAGS.stripOffsets);
  const stripCounts = ifd.numbers(TIFF_TAGS.stripByteCounts);
  const colorMap = ifd.numbers(TIFF_TAGS.colorMap);
  const extraSamples = ifd.numbers(TIFF_TAGS.extraSamples) || [];
  if (!stripOffsets || !stripOffsets.length || stripOffsets.length > 20000) return null;
  if (planar !== 1) return null;
  // 1, 4 and 8 bit samples are the ones Office writes; 16 bit is rare enough
  // that a placeholder beats half-decoding it.
  if (bits.some((b) => b !== 8)) return null;

  const rgba = new Uint8ClampedArray(width * height * 4);
  const stripRows = Math.min(rowsPerStrip, height);
  const rowBytes = width * samples;
  let row = 0;
  for (let s = 0; s < stripOffsets.length && row < height; s++) {
    const start = stripOffsets[s];
    const count = stripCounts && stripCounts[s] !== undefined ? stripCounts[s] : undefined;
    const raw = count === undefined ? bytes.subarray(start) : bytes.subarray(start, start + count);
    const rowsHere = Math.min(stripRows, height - row);
    const decoded = decodeTiffStrip(raw, compression, rowBytes * rowsHere, predictor, samples, rowBytes);
    if (!decoded || decoded.length < rowBytes * rowsHere) return null;
    for (let y = 0; y < rowsHere; y++) {
      for (let x = 0; x < width; x++) {
        const src = (y * width + x) * samples;
        const dst = ((row + y) * width + x) * 4;
        if (photometric === 2 || photometric === 3) {
          rgba[dst] = decoded[src];
          rgba[dst + 1] = decoded[src + 1];
          rgba[dst + 2] = decoded[src + 2];
          rgba[dst + 3] = samples === 4 || extraSamples.length ? decoded[src + 3] : 255;
        } else if (photometric === 1 || photometric === 0) {
          const g = decoded[src];
          rgba[dst] = g;
          rgba[dst + 1] = g;
          rgba[dst + 2] = g;
          rgba[dst + 3] = 255;
        } else if (photometric === 5 && colorMap) {
          const idx = decoded[src];
          const n = colorMap.length / 3;
          rgba[dst] = colorMap[idx] >> 8;
          rgba[dst + 1] = colorMap[idx + n] >> 8;
          rgba[dst + 2] = colorMap[idx + n * 2] >> 8;
          rgba[dst + 3] = 255;
        } else {
          return null;
        }
      }
    }
    row += rowsHere;
  }
  const orientation = ifd.number(TIFF_TAGS.orientation) || 1;
  const oriented = orientRgba(rgba, width, height, orientation);
  return pngDataUrl(oriented.data, oriented.width, oriented.height);
}

function readTiffIfd(view, offset, le, bytes) {
  if (offset + 2 > view.byteLength) return null;
  const count = view.getUint16(offset, le);
  const entries = new Map();
  for (let i = 0; i < count; i++) {
    const entry = offset + 2 + i * 12;
    if (entry + 12 > view.byteLength) return null;
    const tag = view.getUint16(entry, le);
    const type = view.getUint16(entry + 2, le);
    const num = view.getUint32(entry + 4, le);
    const size = tiffTypeSize(type);
    if (num * size > 4) {
      const ptr = view.getUint32(entry + 8, le);
      entries.set(tag, { type, count: num, offset: ptr });
    } else {
      entries.set(tag, { type, count: num, inline: entry + 8 });
    }
  }
  const api = {
    numbers(tag) {
      const t = entries.get(tag);
      if (!t) return null;
      const out = [];
      for (let i = 0; i < t.count; i++) {
        const at = t.inline != null ? t.inline : t.offset;
        out.push(tiffReadValue(view, t.type, at + i * tiffTypeSize(t.type), le));
      }
      return out;
    },
    number(tag) {
      const list = api.numbers(tag);
      return list && list.length ? list[0] : null;
    },
  };
  void bytes;
  return api;
}

function tiffTypeSize(type) {
  switch (type) {
    case 1: case 2: case 6: case 7: return 1;
    case 3: case 8: return 2;
    case 4: case 9: case 11: return 4;
    case 5: case 10: case 12: return 8;
    default: return 1;
  }
}

function tiffReadValue(view, type, offset, le) {
  switch (type) {
    case 1: case 2: case 6: case 7: return view.getUint8(offset);
    case 3: case 8: return view.getUint16(offset, le);
    case 4: case 9: return view.getUint32(offset, le);
    case 5: case 10: return view.getUint32(offset, le);
    default: return view.getUint8(offset);
  }
}

function decodeTiffStrip(raw, compression, expected, predictor, samples, rowBytes) {
  let out;
  if (compression === 1) out = raw;
  else if (compression === 5) out = lzwDecode(raw, expected);
  else if (compression === 32773) out = packBitsDecode(raw, expected);
  else if (compression === 8 || compression === 32946) {
    try {
      out = inflateSync(raw);
    } catch (err) {
      return null;
    }
  } else {
    return null;
  }
  if (!out) return null;
  if (predictor === 2 && rowBytes > 0) {
    // Horizontal differencing restarts at the first pixel of every row: that
    // pixel is stored absolute. Running the loop across the whole strip adds
    // the previous row's last pixel into the next row's first one, and the
    // error then walks through the image, drifting the colours and roughing
    // up the text.
    for (let row = 0; row < out.length; row += rowBytes) {
      const end = Math.min(row + rowBytes, out.length);
      for (let i = row + samples; i < end; i++) {
        out[i] = (out[i] + out[i - samples]) & 255;
      }
    }
  }
  return out;
}

// TIFF LZW: MSB first bit packing, codes from 9 to 12 bits, table resets on a
// clear code. The expected size only hints at the output buffer; the result is
// trimmed to what was actually produced.
function lzwDecode(input, expected) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  let out = new Uint8Array(Math.max(1024, expected > 0 ? expected : bytes.length * 3));
  let outPos = 0;
  let bitPos = 0;
  const CLEAR = 256;
  const EOI = 257;
  let prefixes = new Int32Array(4096);
  let suffix = new Uint8Array(4096);
  let codeSize = 9;
  let nextCode = 258;
  let prevCode = -1;

  function readCode() {
    if (bitPos + codeSize > bytes.length * 8) return EOI;
    let code = 0;
    for (let i = 0; i < codeSize; i++) {
      const byte = bytes[bitPos >> 3];
      code = (code << 1) | ((byte >> (7 - (bitPos & 7))) & 1);
      bitPos++;
    }
    return code;
  }

  function resetTable() {
    for (let i = 0; i < 256; i++) {
      prefixes[i] = -1;
      suffix[i] = i;
    }
    nextCode = 258;
    codeSize = 9;
    prevCode = -1;
  }

  function grow() {
    const bigger = new Uint8Array(out.length * 2);
    bigger.set(out);
    out = bigger;
  }

  // TIFF LZW uses "early change": the code width grows one code before the
  // table is full, at 511, 1023 and 2047 entries.
  function noteNewEntry() {
    if (nextCode === 511 && codeSize === 9) codeSize = 10;
    else if (nextCode === 1023 && codeSize === 10) codeSize = 11;
    else if (nextCode === 2047 && codeSize === 11) codeSize = 12;
  }

  function emitCode(code) {
    // Walk the prefix chain into a scratch buffer, then reverse it.
    let stack = [];
    let c = code;
    let guard = 0;
    while (c >= 0 && guard++ < 4096) {
      stack.push(suffix[c]);
      c = prefixes[c];
    }
    for (let i = stack.length - 1; i >= 0; i--) {
      if (outPos >= out.length) grow();
      out[outPos++] = stack[i];
    }
    return stack.length ? stack[stack.length - 1] : 0;
  }

  resetTable();
  while (true) {
    const code = readCode();
    if (code === EOI) break;
    if (code === CLEAR) {
      resetTable();
      continue;
    }
    if (code < 256) {
      if (outPos >= out.length) grow();
      out[outPos++] = code;
      if (prevCode >= 0 && nextCode < 4096) {
        prefixes[nextCode] = prevCode;
        suffix[nextCode] = code;
        nextCode++;
        noteNewEntry();
      }
      prevCode = code;
      continue;
    }
    if (code < nextCode) {
      const first = emitCode(code);
      if (prevCode >= 0 && nextCode < 4096) {
        prefixes[nextCode] = prevCode;
        suffix[nextCode] = first;
        nextCode++;
        noteNewEntry();
      }
      prevCode = code;
    } else if (prevCode >= 0 && nextCode < 4096) {
      // KwKwK: the code names the entry that is being defined by this very
      // step, so it is prevCode plus the first byte of prevCode.
      const first = firstByteOf(prevCode);
      prefixes[nextCode] = prevCode;
      suffix[nextCode] = first;
      emitCode(nextCode);
      nextCode++;
      noteNewEntry();
      prevCode = nextCode - 1;
    } else {
      break;
    }
  }

  function firstByteOf(code) {
    let c = code;
    let guard = 0;
    while (prefixes[c] >= 0 && guard++ < 4096) c = prefixes[c];
    return suffix[c];
  }

  return out.subarray(0, outPos);
}

function packBitsDecode(input, expected) {
  const out = new Uint8Array(expected > 0 ? expected : input.length * 2);
  let outPos = 0;
  let i = 0;
  while (i < input.length && outPos < out.length) {
    const n = input[i++] << 24 >> 24;
    if (n >= 0) {
      for (let k = 0; k <= n && i < input.length; k++) out[outPos++] = input[i++];
    } else if (n !== -128) {
      const byte = input[i++];
      for (let k = 0; k < 1 - n; k++) out[outPos++] = byte;
    }
  }
  return out.subarray(0, outPos);
}

// ---------- PNG encoding ----------
//
// Pixels become a PNG through a small encoder rather than a canvas round trip.
// That keeps the TIFF decoder usable outside a browser, so the tests see real
// pixels, and it lets EXIF orientation be applied to the buffer before encoding
// instead of through an asynchronous image load.

const PNG_SIGNATURE = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data) {
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length);
  for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
  out.set(data, 8);
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
  return out;
}

function concatBytes(...arrays) {
  let total = 0;
  for (const array of arrays) total += array.length;
  const out = new Uint8Array(total);
  let at = 0;
  for (const array of arrays) {
    out.set(array, at);
    at += array.length;
  }
  return out;
}

function base64FromBytes(bytes) {
  const parts = [];
  for (let i = 0; i < bytes.length; i += 0x8000) {
    parts.push(String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)));
  }
  return btoa(parts.join(""));
}

function pngDataUrl(rgba, width, height) {
  if (!width || !height || rgba.length < width * height * 4) return null;
  const stride = width * 4;
  const filtered = new Uint8Array((stride + 1) * height);
  for (let y = 0; y < height; y++) {
    filtered[y * (stride + 1)] = 0;
    filtered.set(rgba.subarray(y * stride, y * stride + stride), y * (stride + 1) + 1);
  }
  const ihdr = new Uint8Array(13);
  const head = new DataView(ihdr.buffer);
  head.setUint32(0, width);
  head.setUint32(4, height);
  ihdr[8] = 8;
  ihdr[9] = 6;
  const idat = zlibSync(filtered, { level: 6 });
  const png = concatBytes(PNG_SIGNATURE, pngChunk("IHDR", ihdr), pngChunk("IDAT", idat), pngChunk("IEND", new Uint8Array(0)));
  return "data:image/png;base64," + base64FromBytes(png);
}

// The eight EXIF orientations, applied to the pixel buffer.
function orientRgba(rgba, width, height, orientation) {
  if (!orientation || orientation === 1) return { data: rgba, width, height };
  const swap = orientation >= 5;
  const outWidth = swap ? height : width;
  const outHeight = swap ? width : height;
  const out = new Uint8ClampedArray(rgba.length);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      let nx;
      let ny;
      switch (orientation) {
        case 2: nx = width - 1 - x; ny = y; break;
        case 3: nx = width - 1 - x; ny = height - 1 - y; break;
        case 4: nx = x; ny = height - 1 - y; break;
        case 5: nx = y; ny = x; break;
        case 6: nx = height - 1 - y; ny = x; break;
        case 7: nx = height - 1 - y; ny = width - 1 - x; break;
        case 8: nx = y; ny = width - 1 - x; break;
        default: nx = x; ny = y; break;
      }
      const src = (y * width + x) * 4;
      const dst = (ny * outWidth + nx) * 4;
      out[dst] = rgba[src];
      out[dst + 1] = rgba[src + 1];
      out[dst + 2] = rgba[src + 2];
      out[dst + 3] = rgba[src + 3];
    }
  }
  return { data: out, width: outWidth, height: outHeight };
}

// ---------- canvas helpers ----------

// ---------- EMF ----------
//
// An EMF is a stream of GDI records. This player covers the records that real
// Word documents in this vault use: pen and brush creation and selection,
// path building with fill and stroke, the 16 bit polyline and polygon family,
// rectangles and ellipses, clipping rectangles, and text. Record numbers come
// from MS-EMF, not the older WMF numbering, which differs.

const EMR = {
  HEADER: 1,
  POLYBEZIER: 2,
  POLYGON: 3,
  POLYLINE: 4,
  POLYBEZIERTO: 5,
  POLYLINETO: 6,
  SETWINDOWEXTEX: 9,
  SETWINDOWORGEX: 10,
  SETVIEWPORTEXTEX: 11,
  SETVIEWPORTORGEX: 12,
  EOF: 14,
  SETMAPMODE: 17,
  SETBKMODE: 18,
  SETPOLYFILLMODE: 19,
  SETTEXTALIGN: 22,
  SETTEXTCOLOR: 24,
  SETBKCOLOR: 25,
  MOVETOEX: 27,
  INTERSECTCLIPRECT: 30,
  SAVEDC: 33,
  RESTOREDC: 34,
  SELECTOBJECT: 37,
  CREATEPEN: 38,
  CREATEBRUSHINDIRECT: 39,
  DELETEOBJECT: 40,
  ELLIPSE: 42,
  RECTANGLE: 43,
  ROUNDRECT: 44,
  LINETO: 54,
  BEGINPATH: 59,
  ENDPATH: 60,
  CLOSEFIGURE: 61,
  FILLPATH: 62,
  STROKEANDFILLPATH: 63,
  STROKEPATH: 64,
  SELECTCLIPPATH: 67,
  EXTSELECTCLIPRGN: 75,
  EXTTEXTOUTW: 84,
  POLYBEZIER16: 85,
  POLYGON16: 86,
  POLYLINE16: 87,
  POLYBEZIERTO16: 88,
  POLYLINE16X: 89,
  POLYLINETO16: 90,
  POLYPOLYGON16: 91,
  EXTCREATEPEN: 82,
};

function metafileToDataUrl(bytes, kind) {
  if (bytes.length < 24) return null;
  if (isGzip(bytes)) {
    try {
      bytes = inflateSync(bytes);
    } catch (err) {
      return null;
    }
  }
  try {
    if (kind === "wmf" && looksLikeWmf(bytes)) return wmfToDataUrl(bytes);
    if (looksLikeEmf(bytes)) return emfToDataUrl(bytes);
    if (looksLikeWmf(bytes)) return wmfToDataUrl(bytes);
  } catch (err) {
    return null;
  }
  return null;
}

function isGzip(bytes) {
  return bytes.length > 2 && bytes[0] === 0x1f && bytes[1] === 0x8b;
}

function looksLikeEmf(bytes) {
  if (bytes.length < 44) return false;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return view.getUint32(0, true) === 1 && view.getUint32(40, true) === 0x464d4520;
}

function looksLikeWmf(bytes) {
  if (bytes.length < 18) return false;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.getUint32(0, true) === 0x9ac6cdd7) return true;
  const type = view.getUint16(0, true);
  return (type === 1 || type === 2) && view.getUint16(2, true) === 9;
}

// The stock objects GDI provides without a create call. An EMF that only
// creates a brush still draws with a pen, and that pen is often what makes the
// white parts of a logo.
const STOCK_OBJECTS = new Map([
  [0x80000000, { kind: "brush", color: "#ffffff", style: "solid" }],
  [0x80000001, { kind: "brush", color: "#c0c0c0", style: "solid" }],
  [0x80000002, { kind: "brush", color: "#808080", style: "solid" }],
  [0x80000003, { kind: "brush", color: "#404040", style: "solid" }],
  [0x80000004, { kind: "brush", color: "#000000", style: "solid" }],
  [0x80000005, { kind: "brush", color: null, style: "hollow" }],
  [0x80000006, { kind: "pen", color: "#ffffff", width: 1 }],
  [0x80000007, { kind: "pen", color: "#000000", width: 1 }],
  [0x80000008, { kind: "pen", color: null, width: 0 }],
]);

function createGdiState() {
  return {
    objects: new Map(),
    pen: { color: "#000000", width: 1, dash: [], style: "solid" },
    brush: { color: null, style: "hollow" },
    font: null,
    textColor: "#000000",
    bkColor: "#ffffff",
    bkMode: "opaque",
    textAlign: 0,
    polyFill: "alternate",
    windowOrg: null,
    windowExt: null,
    viewportOrg: null,
    viewportExt: null,
    useMapping: false,
  };
}

function cloneGdiState(state) {
  return {
    objects: new Map(),
    pen: Object.assign({}, state.pen),
    brush: Object.assign({}, state.brush),
    font: state.font,
    textColor: state.textColor,
    bkColor: state.bkColor,
    bkMode: state.bkMode,
    textAlign: state.textAlign,
    polyFill: state.polyFill,
    windowOrg: state.windowOrg && Object.assign({}, state.windowOrg),
    windowExt: state.windowExt && Object.assign({}, state.windowExt),
    viewportOrg: state.viewportOrg && Object.assign({}, state.viewportOrg),
    viewportExt: state.viewportExt && Object.assign({}, state.viewportExt),
    useMapping: state.useMapping,
  };
}

function colorRefFromRgb(r, g, b) {
  const to = (v) => Math.max(0, Math.min(255, v | 0)).toString(16).padStart(2, "0");
  return "#" + to(r) + to(g) + to(b);
}

// COLORREF is 0x00BBGGRR in a little endian long.
function colorRef(value) {
  return colorRefFromRgb(value & 0xff, (value >> 8) & 0xff, (value >> 16) & 0xff);
}

function emfToDataUrl(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const header = readEmfHeader(view);
  if (!header) return null;
  const bounds = header.bounds;
  const logicalW = Math.max(1, bounds.right - bounds.left);
  const logicalH = Math.max(1, bounds.bottom - bounds.top);
  const canvas = document.createElement("canvas");
  const cap = 2400;
  const fit = Math.min(1, cap / Math.max(logicalW, logicalH));
  canvas.width = Math.max(1, Math.round(logicalW * fit));
  canvas.height = Math.max(1, Math.round(logicalH * fit));
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.save();
  // Device space starts at the header bounds, scaled to the canvas.
  ctx.scale(fit, fit);
  ctx.translate(-bounds.left, -bounds.top);
  const state = createGdiState();
  state.windowOrg = { x: bounds.left, y: bounds.top };
  state.windowExt = { cx: logicalW, cy: logicalH };
  state.viewportOrg = { x: bounds.left, y: bounds.top };
  state.viewportExt = { cx: logicalW, cy: logicalH };
  playEmf(ctx, view, state, { bounds, logicalW, logicalH });
  ctx.restore();
  return canvas.toDataURL("image/png");
}

function readEmfHeader(view) {
  if (view.getUint32(0, true) !== 1) return null;
  const bounds = {
    left: view.getInt32(8, true),
    top: view.getInt32(12, true),
    right: view.getInt32(16, true),
    bottom: view.getInt32(20, true),
  };
  if (bounds.right === bounds.left) bounds.right = bounds.left + 1000;
  if (bounds.bottom === bounds.top) bounds.bottom = bounds.top + 1000;
  return { bounds };
}

function playEmf(ctx, view, state, layout) {
  const { bounds } = layout;
  let offset = 0;
  let path = [];
  let current = null;
  let textPos = { x: 0, y: 0 };
  let clip = null;
  const stack = [];
  let guard = 0;

  const mapPoint = (x, y) => {
    if (!state.useMapping || !state.windowExt || !state.viewportExt) return { x, y };
    const wx = state.windowOrg ? state.windowOrg.x : 0;
    const wy = state.windowOrg ? state.windowOrg.y : 0;
    const vx = state.viewportOrg ? state.viewportOrg.x : 0;
    const vy = state.viewportOrg ? state.viewportOrg.y : 0;
    const sx = state.viewportExt.cx / (state.windowExt.cx || 1);
    const sy = state.viewportExt.cy / (state.windowExt.cy || 1);
    return { x: vx + (x - wx) * sx, y: vy + (y - wy) * sy };
  };

  const applyClip = () => {
    if (!clip) return;
    ctx.save();
    ctx.beginPath();
    ctx.rect(clip.left, clip.top, clip.right - clip.left, clip.bottom - clip.top);
    ctx.clip();
    clip = null;
  };

  const beginPath = () => {
    path = [];
    current = null;
  };

  const moveTo = (x, y) => {
    const p = mapPoint(x, y);
    path.push({ type: "M", x: p.x, y: p.y });
    current = p;
  };

  const lineTo = (x, y) => {
    const p = mapPoint(x, y);
    if (!path.length) path.push({ type: "M", x: p.x, y: p.y });
    else path.push({ type: "L", x: p.x, y: p.y });
    current = p;
  };

  const closeFigure = () => {
    if (path.length) path.push({ type: "Z" });
  };

  const buildPath = () => {
    ctx.beginPath();
    for (const segment of path) {
      if (segment.type === "M") ctx.moveTo(segment.x, segment.y);
      else if (segment.type === "L") ctx.lineTo(segment.x, segment.y);
      else if (segment.type === "C") ctx.bezierCurveTo(segment.c1x, segment.c1y, segment.c2x, segment.c2y, segment.x, segment.y);
      else if (segment.type === "Z") ctx.closePath();
    }
  };

  const strokeStyleOf = () => {
    const pen = state.pen;
    ctx.strokeStyle = pen.color || "#000000";
    ctx.lineWidth = Math.max(0.5, pen.width || 1);
    ctx.lineJoin = "round";
    ctx.lineCap = "butt";
    ctx.setLineDash(pen.dash && pen.dash.length ? pen.dash : []);
  };

  const strokePath = (close) => {
    if (!path.length) return;
    strokeStyleOf();
    if (close) closeFigure();
    buildPath();
    ctx.stroke();
  };

  const fillPath = () => {
    const brush = state.brush;
    const paint = brush && brush.style !== "hollow" && brush.color;
    if (paint) {
      ctx.fillStyle = brush.color;
      ctx.fillRule = state.polyFill === "winding" ? "nonzero" : "evenodd";
      buildPath();
      ctx.fill(ctx.fillRule);
    }
    // GDI's FillPath also draws the path's outline with the current pen.
    const pen = state.pen;
    if (pen && pen.color && pen.width > 0) {
      strokeStyleOf();
      buildPath();
      ctx.stroke();
    }
  };

  // Paths used as a clip region: approximate with the path's bounding box,
  // which is what SelectClipPath means for the shapes Word exports.
  const pathBounds = () => {
    if (!path.length) return null;
    let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity;
    for (const segment of path) {
      const x = segment.x;
      const y = segment.y;
      if (x < left) left = x;
      if (y < top) top = y;
      if (x > right) right = x;
      if (y > bottom) bottom = y;
    }
    return { left, top, right, bottom };
  };

  const rectOf = (dataOff) => {
    const left = view.getInt32(dataOff, true);
    const top = view.getInt32(dataOff + 4, true);
    const right = view.getInt32(dataOff + 8, true);
    const bottom = view.getInt32(dataOff + 12, true);
    return { left, top, right, bottom };
  };

  // EMR_POLYLINE16 and its relatives lay out as rclBounds (16 bytes), cpts (4),
  // then the points. Reading the count out of the bounds turned every contour
  // into a single point.
  const points16 = (dataOff) => {
    const count = view.getInt32(dataOff + 16, true);
    const base = dataOff + 20;
    if (count < 0 || count > 4096) return [];
    const out = [];
    for (let i = 0; i < count; i++) {
      if (base + i * 4 + 4 > view.byteLength) break;
      const x = view.getInt16(base + i * 4, true);
      const y = view.getInt16(base + i * 4 + 2, true);
      out.push({ x, y });
    }
    return out;
  };

  // A malformed record ends the metafile instead of the render: the records
  // already played stay on the canvas.
  try {
  while (offset + 8 <= view.byteLength && guard++ < 200000) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (type === 0 || size < 8 || offset + size > view.byteLength) break;
    const data = offset + 8;

    switch (type) {
      case EMR.HEADER:
      case EMR.EOF:
      case EMR.SETWINDOWEXTEX:
        if (type === EMR.SETWINDOWEXTEX) {
          state.windowExt = { cx: view.getInt32(data, true), cy: view.getInt32(data + 4, true) };
          state.useMapping = true;
        }
        break;
      case EMR.SETWINDOWORGEX:
        state.windowOrg = { x: view.getInt32(data, true), y: view.getInt32(data + 4, true) };
        state.useMapping = true;
        break;
      case EMR.SETVIEWPORTEXTEX:
        state.viewportExt = { cx: view.getInt32(data, true), cy: view.getInt32(data + 4, true) };
        state.useMapping = true;
        break;
      case EMR.SETVIEWPORTORGEX:
        state.viewportOrg = { x: view.getInt32(data, true), y: view.getInt32(data + 4, true) };
        state.useMapping = true;
        break;
      case EMR.SETMAPMODE:
        break;
      case EMR.SETBKMODE:
        state.bkMode = view.getUint32(data, true) === 1 ? "transparent" : "opaque";
        break;
      case EMR.SETPOLYFILLMODE:
        state.polyFill = view.getUint32(data, true) === 2 ? "winding" : "alternate";
        break;
      case EMR.SETTEXTALIGN:
        state.textAlign = view.getUint32(data, true);
        break;
      case EMR.SETTEXTCOLOR:
        state.textColor = colorRef(view.getUint32(data, true));
        break;
      case EMR.SETBKCOLOR:
        state.bkColor = colorRef(view.getUint32(data, true));
        break;
      case EMR.CREATEPEN: {
        const handle = view.getUint32(data, true);
        const style = view.getUint32(data + 4, true);
        const width = view.getInt32(data + 8, true);
        putObject(state, handle, {
          kind: "pen",
          color: "#000000",
          width: Math.max(1, width),
          dash: dashForStyle(style),
        });
        break;
      }
      case EMR.EXTCREATEPEN: {
        const handle = view.getUint32(data, true);
        const color = view.getUint32(data + 24, true);
        const style = view.getUint32(data + 32, true);
        const width = size >= 44 ? view.getUint32(data + 36, true) : 1;
        putObject(state, handle, {
          kind: "pen",
          color: colorRef(color),
          width: Math.max(1, width || 1),
          dash: dashForStyle(style),
        });
        break;
      }
      case EMR.CREATEBRUSHINDIRECT: {
        const handle = view.getUint32(data, true);
        const style = view.getUint32(data + 4, true);
        const color = view.getUint32(data + 8, true);
        putObject(state, handle, {
          kind: "brush",
          style: style === 1 ? "hollow" : "solid",
          color: colorRef(color),
        });
        break;
      }
      case EMR.SELECTOBJECT: {
        const handle = view.getUint32(data, true);
        const stock = STOCK_OBJECTS.get(handle);
        const object = stock || state.objects.get(handle);
        if (object) {
          if (object.kind === "pen") state.pen = object;
          else if (object.kind === "brush") state.brush = object;
          else if (object.kind === "font") state.font = object;
        }
        break;
      }
      case EMR.DELETEOBJECT:
        state.objects.delete(view.getUint32(data, true));
        break;
      case EMR.SAVEDC:
        stack.push(cloneGdiState(state));
        break;
      case EMR.RESTOREDC: {
        const saved = stack.pop();
        if (saved) {
          state.pen = saved.pen;
          state.brush = saved.brush;
          state.textColor = saved.textColor;
          state.bkColor = saved.bkColor;
          state.bkMode = saved.bkMode;
          state.textAlign = saved.textAlign;
          state.polyFill = saved.polyFill;
        }
        break;
      }
      case EMR.INTERSECTCLIPRECT: {
        const rect = rectOf(data);
        ctx.save();
        ctx.beginPath();
        ctx.rect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);
        ctx.clip();
        break;
      }
      case EMR.SELECTCLIPPATH: {
        const box = pathBounds();
        if (box) {
          ctx.save();
          ctx.beginPath();
          ctx.rect(box.left, box.top, box.right - box.left, box.bottom - box.top);
          ctx.clip();
        }
        break;
      }
      case EMR.EXTSELECTCLIPRGN:
        break;
      case EMR.MOVETOEX: {
        const x = view.getInt32(data, true);
        const y = view.getInt32(data + 4, true);
        if (path.length && current) lineTo(x, y);
        else moveTo(x, y);
        break;
      }
      case EMR.LINETO: {
        const x = view.getInt32(data, true);
        const y = view.getInt32(data + 4, true);
        lineTo(x, y);
        break;
      }
      case EMR.BEGINPATH:
        beginPath();
        break;
      case EMR.ENDPATH:
        break;
      case EMR.CLOSEFIGURE:
        closeFigure();
        break;
      case EMR.FILLPATH:
        fillPath();
        break;
      case EMR.STROKEPATH:
        strokePath(false);
        break;
      case EMR.STROKEANDFILLPATH:
        fillPath();
        strokePath(false);
        break;
      case EMR.POLYGON:
      case EMR.POLYLINE:
      case EMR.POLYBEZIER: {
        // The 32 bit forms use POINTL points: bounds (16), cpts (4), points.
        const count = view.getInt32(data + 16, true);
        if (count < 0 || count > 4096) break;
        const base = data + 20;
        const points = [];
        for (let i = 0; i < count; i++) {
          if (base + i * 8 + 8 > view.byteLength) break;
          points.push({ x: view.getInt32(base + i * 8, true), y: view.getInt32(base + i * 8 + 4, true) });
        }
        if (!points.length) break;
        const mapped = points.map((pt) => mapPoint(pt.x, pt.y));
        ctx.beginPath();
        ctx.moveTo(mapped[0].x, mapped[0].y);
        for (let i = 1; i < mapped.length; i++) ctx.lineTo(mapped[i].x, mapped[i].y);
        if (type === EMR.POLYGON) {
          ctx.closePath();
          fillPathDirect(ctx, state);
        }
        strokeStyleOf();
        ctx.stroke();
        break;
      }
      case EMR.POLYLINE16:
      case EMR.POLYBEZIER16: {
        // These start a new contour at their first point. Chaining them onto
        // the previous contour would weld a letter's counter to its outline and
        // fill the hole in, which is exactly what a solid logo block is.
        const points = points16(data);
        if (!points.length) break;
        moveTo(points[0].x, points[0].y);
        for (let i = 1; i < points.length; i++) lineTo(points[i].x, points[i].y);
        break;
      }
      case EMR.POLYLINETO16:
      case EMR.POLYBEZIERTO16:
      case EMR.POLYLINE16X: {
        // The continuation forms carry on from the current point.
        const points = points16(data);
        for (let i = 0; i < points.length; i++) lineTo(points[i].x, points[i].y);
        break;
      }
      case EMR.POLYGON16: {
        const points = points16(data);
        if (points.length < 2) break;
        const mapped = points.map((p) => mapPoint(p.x, p.y));
        if (path.length) {
          for (const p of mapped) lineTo(p.x, p.y);
          closeFigure();
          break;
        }
        ctx.beginPath();
        ctx.moveTo(mapped[0].x, mapped[0].y);
        for (let i = 1; i < mapped.length; i++) ctx.lineTo(mapped[i].x, mapped[i].y);
        ctx.closePath();
        fillPathDirect(ctx, state);
        strokeStyleOf();
        ctx.stroke();
        break;
      }
      case EMR.POLYPOLYGON16: {
        const polygons = readPolyPolygon16(view, data);
        for (const points of polygons) {
          if (points.length < 2) continue;
          const mapped = points.map((p) => mapPoint(p.x, p.y));
          ctx.beginPath();
          ctx.moveTo(mapped[0].x, mapped[0].y);
          for (let i = 1; i < mapped.length; i++) ctx.lineTo(mapped[i].x, mapped[i].y);
          ctx.closePath();
          fillPathDirect(ctx, state);
          strokeStyleOf();
          ctx.stroke();
        }
        break;
      }
      case EMR.RECTANGLE: {
        const rect = rectOf(data);
        ctx.beginPath();
        ctx.rect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);
        fillPathDirect(ctx, state);
        strokeStyleOf();
        ctx.stroke();
        break;
      }
      case EMR.ROUNDRECT: {
        const rect = rectOf(data);
        const radius = Math.max(1, Math.abs(view.getInt32(data + 16, true)) / 2);
        ctx.beginPath();
        roundRectPath(ctx, rect.left, rect.top, rect.right, rect.bottom, radius);
        fillPathDirect(ctx, state);
        strokeStyleOf();
        ctx.stroke();
        break;
      }
      case EMR.ELLIPSE: {
        const rect = rectOf(data);
        ctx.beginPath();
        ctx.ellipse((rect.left + rect.right) / 2, (rect.top + rect.bottom) / 2,
          Math.abs(rect.right - rect.left) / 2, Math.abs(rect.bottom - rect.top) / 2, 0, 0, Math.PI * 2);
        fillPathDirect(ctx, state);
        strokeStyleOf();
        ctx.stroke();
        break;
      }
      case EMR.EXTTEXTOUTW: {
        // rclBounds (16), iGraphicsMode (4), exScale (4), eyScale (4), then
        // the EMRTEXT: reference point (8), nChars (4), offString (4).
        const referenceX = view.getInt32(data + 28, true);
        const referenceY = view.getInt32(data + 32, true);
        const chars = view.getInt32(data + 36, true);
        const offString = view.getUint32(data + 40, true);
        if (offString && chars > 0 && offString + chars * 2 <= view.byteLength) {
          let text = "";
          for (let i = 0; i < chars; i++) text += String.fromCharCode(view.getUint16(offString + i * 2, true));
          const font = state.font || { size: 12, family: "sans-serif" };
          const point = mapPoint(referenceX, referenceY);
          ctx.font = font.weight + " " + Math.abs(font.size || 12) + "px " + font.family;
          ctx.textBaseline = "alphabetic";
          ctx.fillStyle = state.textColor;
          ctx.fillText(text, point.x, point.y);
          textPos = { x: referenceX, y: referenceY };
        }
        break;
      }
      default:
        break;
    }
    offset += size;
  }
  } catch (err) {
    // Stop playing; keep what was drawn.
  }
  void textPos;
  void applyClip;
  void current;
}

function fillPathDirect(ctx, state) {
  const brush = state.brush;
  if (!brush || brush.style === "hollow" || !brush.color) return;
  ctx.fillStyle = brush.color;
  ctx.fill();
}

function roundRectPath(ctx, left, top, right, bottom, radius) {
  const r = Math.min(radius, Math.abs(right - left) / 2, Math.abs(bottom - top) / 2);
  ctx.moveTo(left + r, top);
  ctx.lineTo(right - r, top);
  ctx.quadraticCurveTo(right, top, right, top + r);
  ctx.lineTo(right, bottom - r);
  ctx.quadraticCurveTo(right, bottom, right - r, bottom);
  ctx.lineTo(left + r, bottom);
  ctx.quadraticCurveTo(left, bottom, left, bottom - r);
  ctx.lineTo(left, top + r);
  ctx.quadraticCurveTo(left, top, left + r, top);
  ctx.closePath();
}

function dashForStyle(style) {
  switch (style & 0x0f) {
    case 1:
    case 5: return [8, 6];
    case 2: return [2, 4];
    case 3: return [8, 4, 2, 4];
    case 4: return [8, 4, 2, 4, 2, 4];
    default: return [];
  }
}

// PolyPolygon16 carries a count per polygon followed by all their points.
function readPolyPolygon16(view, dataOff) {
  const polygons = view.getInt32(dataOff + 16, true);
  if (polygons < 0 || polygons > 1024) return [];
  const countsOff = dataOff + 20;
  let total = 0;
  const counts = [];
  for (let i = 0; i < polygons; i++) {
    if (countsOff + i * 4 + 4 > view.byteLength) return [];
    const count = view.getInt32(countsOff + i * 4, true);
    counts.push(count);
    total += count;
  }
  const pointsOff = countsOff + polygons * 4;
  const out = [];
  let at = 0;
  for (const count of counts) {
    const points = [];
    const usable = Math.min(count, 65536);
    for (let i = 0; i < usable; i++) {
      const spot = pointsOff + (at + i) * 4;
      if (spot + 4 > view.byteLength) break;
      points.push({
        x: view.getInt16(spot, true),
        y: view.getInt16(spot + 2, true),
      });
    }
    at += count;
    out.push(points);
  }
  void total;
  return out;
}

// A metafile can create objects forever. Past a few thousand the map is a
// memory leak with a drawing attached, so new handles stop being recorded.
const MAX_METAFILE_OBJECTS = 4096;

function putObject(state, handle, object) {
  if (state.objects.size >= MAX_METAFILE_OBJECTS && !state.objects.has(handle)) return;
  state.objects.set(handle, object);
}

// ---------- WMF ----------

function wmfToDataUrl(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 0;
  let placeable = false;
  let width = 800;
  let height = 600;
  if (view.getUint32(0, true) === 0x9ac6cdd7) {
    placeable = true;
    const left = view.getInt16(6, true);
    const top = view.getInt16(8, true);
    const right = view.getInt16(10, true);
    const bottom = view.getInt16(12, true);
    const units = view.getUint16(14, true) || 1;
    width = Math.max(1, Math.abs(right - left) * units);
    height = Math.max(1, Math.abs(bottom - top) * units);
    offset = 22;
  } else {
    offset = 18;
  }
  if (offset >= bytes.length) return null;
  const canvas = document.createElement("canvas");
  const fit = Math.min(1, 2400 / Math.max(width, height));
  canvas.width = Math.max(1, Math.round(width * fit));
  canvas.height = Math.max(1, Math.round(height * fit));
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.save();
  ctx.scale(fit, fit);
  const state = createGdiState();
  const stack = [];
  let x = 0;
  let y = 0;
  // A malformed record ends the metafile instead of the render.
  try {
  while (offset + 6 <= bytes.length) {
    const size = view.getUint32(offset, true);
    if (size < 3 || offset + size * 2 > bytes.length) break;
    const fn = view.getUint16(offset + 4, true);
    const p = (i) => offset + 6 + i * 2;
    switch (fn) {
      case 0x0a32:
        ctx.fillStyle = state.bkColor;
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        break;
      case 0x0213:
      case 0x0214:
        x = view.getInt16(p(0), true);
        y = view.getInt16(p(1), true);
        break;
      case 0x0216: {
        const nx = view.getInt16(p(0), true);
        const ny = view.getInt16(p(1), true);
        ctx.strokeStyle = state.pen.color;
        ctx.lineWidth = Math.max(1, state.pen.width);
        ctx.beginPath();
        ctx.moveTo(x, y);
        ctx.lineTo(nx, ny);
        ctx.stroke();
        x = nx;
        y = ny;
        break;
      }
      case 0x021E: {
        const count = Math.min(Math.max(0, view.getInt16(p(0), true)), 4096);
        const points = [];
        for (let i = 0; i < count; i++) {
          if (offset + 10 + i * 4 + 2 > bytes.length) break;
          points.push({ x: view.getInt16(offset + 8 + i * 4, true), y: view.getInt16(offset + 10 + i * 4, true) });
        }
        drawWmfPolygon(ctx, points, state, true);
        break;
      }
      case 0x0324: {
        const count = Math.min(Math.max(0, view.getInt16(p(0), true)), 4096);
        const points = [];
        for (let i = 0; i < count; i++) {
          if (offset + 10 + i * 4 + 2 > bytes.length) break;
          points.push({ x: view.getInt16(offset + 8 + i * 4, true), y: view.getInt16(offset + 10 + i * 4, true) });
        }
        drawWmfPolygon(ctx, points, state, false);
        break;
      }
      case 0x041B: {
        const w = view.getInt16(p(0), true);
        const h = view.getInt16(p(1), true);
        ctx.strokeStyle = state.pen.color;
        ctx.lineWidth = Math.max(1, state.pen.width);
        ctx.strokeRect(x, y, w, h);
        break;
      }
      case 0x0418: {
        const h = view.getInt16(p(0), true);
        const w = view.getInt16(p(1), true);
        ctx.beginPath();
        ctx.ellipse(x + w / 2, y + h / 2, Math.abs(w / 2), Math.abs(h / 2), 0, 0, Math.PI * 2);
        if (state.brush.color) {
          ctx.fillStyle = state.brush.color;
          ctx.fill();
        }
        ctx.strokeStyle = state.pen.color;
        ctx.lineWidth = Math.max(1, state.pen.width);
        ctx.stroke();
        break;
      }
      case 0x02FA:
        state.textColor = colorRef(view.getUint32(p(0), true));
        break;
      case 0x02FB:
        state.bkColor = colorRef(view.getUint32(p(0), true));
        break;
      case 0x02FC: {
        const count = Math.min(Math.max(0, view.getInt16(p(0), true)), 1024);
        for (let i = 0; i < count; i++) {
          const entry = offset + 8 + i * 8;
          if (entry + 8 > bytes.length) break;
          const objectType = view.getUint16(entry, true) & 0xff00;
          if (objectType === 0x0600) {
            const color = colorRef(view.getUint32(entry + 4, true));
            state.pen = { color, width: Math.max(1, view.getInt16(entry + 6, true)), dash: [], style: "solid" };
            state.brush = { color, style: "solid" };
          }
        }
        break;
      }
      case 0x001E:
        stack.push({ pen: state.pen, brush: state.brush });
        break;
      case 0x012E: {
        const saved = stack.pop();
        if (saved) {
          state.pen = saved.pen;
          state.brush = saved.brush;
        }
        break;
      }
      case 0x0621:
        if (state.brush.color) {
          ctx.fillStyle = state.brush.color;
          ctx.fillRect(0, 0, canvas.width, canvas.height);
        }
        break;
      default:
        break;
    }
    offset += size * 2;
  }
  } catch (err) {
    // A malformed record ends the metafile, not the render: whatever was drawn
    // up to that point is still shown.
  }
  ctx.restore();
  return canvas.toDataURL("image/png");
}

function drawWmfPolygon(ctx, points, state, close) {
  if (points.length < 2) return;
  ctx.beginPath();
  ctx.moveTo(points[0].x, points[0].y);
  for (let i = 1; i < points.length; i++) ctx.lineTo(points[i].x, points[i].y);
  if (close) ctx.closePath();
  if (close && state.brush.color) {
    ctx.fillStyle = state.brush.color;
    ctx.fill();
  }
  ctx.strokeStyle = state.pen.color;
  ctx.lineWidth = Math.max(1, state.pen.width);
  ctx.stroke();
}

module.exports = {
  createMediaCache,
  mediaUrl,
  buildMediaUrl,
  blobUrl,
  bytesToDataUrl,
  tiffToDataUrl,
  metafileToDataUrl,
  lzwDecode,
  packBitsDecode,
};
