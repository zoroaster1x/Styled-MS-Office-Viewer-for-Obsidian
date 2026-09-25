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
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
 * details.
 *
 * You should have received a copy of the GNU General Public License along with
 * this program. If not, see <https://www.gnu.org/licenses/>.
 */

// The compound document reader and the Word 97 text reader, against files this
// test builds byte by byte: a plain document, a piece table with one 8 bit and
// one 16 bit piece, field markup, and damaged containers.
//
//   bun test/doc-legacy.mjs

import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const { openCompound, isCompound } = require("../src/legacy/ole.js");
const { readDoc, cleanText } = require("../src/legacy/doc.js");

let pass = 0;
let fail = 0;
function check(name, condition, detail) {
  if (condition) {
    pass++;
    return;
  }
  fail++;
  console.log("FAIL " + name + (detail ? "  -> " + detail : ""));
}

// ---------- a compound file builder ----------
//
// 512 byte sectors, one FAT sector, one directory sector, and a mini stream for
// anything under 4096 bytes, which is the shape real files have.

const END = 0xfffffffe;
const FREE = 0xffffffff;
const FATSECT = 0xfffffffd;

function padToSector(bytes, size) {
  const padded = new Uint8Array(Math.ceil(Math.max(bytes.length, 1) / size) * size);
  padded.set(bytes);
  return padded;
}

function buildCompound(streams) {
  const sectorSize = 512;
  const miniSize = 64;
  const dataSectors = []; // one entry per sector: Uint8Array of 512 bytes
  const fat = [];
  const miniFat = [];
  const miniData = [];

  const nextSector = () => dataSectors.length;

  // Big streams first, each as its own chain.
  const dirEntries = [];
  const bigStreams = streams.filter((s) => s.bytes.length >= 4096);
  const smallStreams = streams.filter((s) => s.bytes.length < 4096);
  for (const stream of bigStreams) {
    const padded = padToSector(stream.bytes, sectorSize);
    const start = nextSector();
    const count = padded.length / sectorSize;
    for (let i = 0; i < count; i++) dataSectors.push(padded.subarray(i * sectorSize, (i + 1) * sectorSize));
    for (let i = 0; i < count; i++) fat[start + i] = i === count - 1 ? END : start + i + 1;
    dirEntries.push({ name: stream.name, start, size: stream.size != null ? stream.size : stream.bytes.length });
  }

  // Small streams share the root entry's mini stream, 64 bytes at a time.
  let miniRootStart = END;
  if (smallStreams.length) {
    const miniBytes = [];
    for (const stream of smallStreams) {
      const startMini = miniData.length;
      const padded = padToSector(stream.bytes, miniSize);
      const count = padded.length / miniSize;
      for (let i = 0; i < count; i++) miniData.push(padded.subarray(i * miniSize, (i + 1) * miniSize));
      for (let i = 0; i < count; i++) miniFat[startMini + i] = i === count - 1 ? END : startMini + i + 1;
      dirEntries.push({ name: stream.name, start: startMini, size: stream.size != null ? stream.size : stream.bytes.length });
    }
    const flat = new Uint8Array(miniData.length * miniSize);
    miniData.forEach((chunk, i) => flat.set(chunk, i * miniSize));
    const padded = padToSector(flat, sectorSize);
    miniRootStart = nextSector();
    const count = padded.length / sectorSize;
    for (let i = 0; i < count; i++) dataSectors.push(padded.subarray(i * sectorSize, (i + 1) * sectorSize));
    for (let i = 0; i < count; i++) fat[miniRootStart + i] = i === count - 1 ? END : miniRootStart + i + 1;
    // A mini FAT sector.
    const miniFatSector = nextSector();
    const miniFatBytes = new Uint8Array(sectorSize);
    const miniView = new DataView(miniFatBytes.buffer);
    for (let i = 0; i < sectorSize / 4; i++) miniView.setUint32(i * 4, i < miniFat.length ? miniFat[i] : FREE, true);
    dataSectors.push(miniFatBytes);
    fat[miniFatSector] = FATSECT;
    buildCompound.miniFatSector = miniFatSector;
    buildCompound.miniFatCount = 1;
  } else {
    buildCompound.miniFatSector = END;
    buildCompound.miniFatCount = 0;
  }

  // The directory.
  const dirSector = nextSector();
  const directory = new Uint8Array(sectorSize * Math.max(1, Math.ceil((dirEntries.length + 1) / 4)));
  const dirView = new DataView(directory.buffer);
  const writeEntry = (index, entry) => {
    const at = index * 128;
    const name = entry.name + "\u0000";
    for (let i = 0; i < name.length && i < 32; i++) {
      dirView.setUint16(at + i * 2, name.charCodeAt(i), true);
    }
    dirView.setUint16(at + 64, Math.min(64, (entry.name.length + 1) * 2), true);
    dirView.setUint8(at + 66, entry.type);
    dirView.setUint8(at + 67, 1);
    dirView.setUint32(at + 68, FREE, true);
    dirView.setUint32(at + 72, FREE, true);
    dirView.setUint32(at + 76, FREE, true);
    dirView.setUint32(at + 116, entry.start, true);
    dirView.setBigUint64(at + 120, BigInt(entry.size), true);
  };
  writeEntry(0, { name: "Root Entry", type: 5, start: miniRootStart, size: miniData.length * miniSize });
  dirEntries.forEach((entry, i) => writeEntry(i + 1, { name: entry.name, type: 2, start: entry.start, size: entry.size }));
  const dirCount = Math.ceil(directory.length / sectorSize);
  for (let i = 0; i < dirCount; i++) {
    const start = nextSector();
    dataSectors.push(directory.subarray(i * sectorSize, (i + 1) * sectorSize));
    fat[start] = i === dirCount - 1 ? END : start + 1;
  }

  // The FAT itself.
  const fatSector = nextSector();
  const fatBytes = new Uint8Array(sectorSize);
  const fatView = new DataView(fatBytes.buffer);
  for (let i = 0; i < sectorSize / 4; i++) fatView.setUint32(i * 4, i < fat.length ? fat[i] : FREE, true);
  dataSectors.push(fatBytes);

  const header = new Uint8Array(512);
  const view = new DataView(header.buffer);
  const signature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  signature.forEach((byte, i) => view.setUint8(i, byte));
  view.setUint16(24, 0x003e, true);
  view.setUint16(26, 0x0003, true);
  view.setUint16(28, 0xfffe, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(40, 0, true);
  view.setUint32(44, 1, true);
  view.setUint32(48, dirSector, true);
  view.setUint32(52, 0, true);
  view.setUint32(56, 4096, true);
  view.setUint32(60, buildCompound.miniFatSector, true);
  view.setUint32(64, buildCompound.miniFatCount, true);
  view.setUint32(68, END, true);
  view.setUint32(72, 0, true);
  for (let i = 0; i < 109; i++) view.setUint32(76 + i * 4, i === 0 ? fatSector : FREE, true);

  const out = new Uint8Array(512 + dataSectors.length * sectorSize);
  out.set(header, 0);
  dataSectors.forEach((sector, i) => out.set(sector, 512 + i * sectorSize));
  return out;
}

// ---------- a Word 97 document builder ----------

const FIB_SIZE = 0x200;

function buildDocStream(pieces, options) {
  const opts = options || {};
  // pieces: [{ text, wide }]
  const fib = new Uint8Array(FIB_SIZE);
  const view = new DataView(fib.buffer);
  view.setUint16(0, 0xa5ec, true);
  view.setUint16(2, 193, true);
  view.setUint16(10, opts.complex ? 0x0204 : 0x0000, true);
  const body = [];
  let offset = FIB_SIZE;
  const pcds = [];
  for (const piece of pieces) {
    const wide = Boolean(piece.wide);
    pcds.push({ offset, wide });
    if (wide) {
      const bytes = new Uint8Array(piece.text.length * 2);
      const dataView = new DataView(bytes.buffer);
      for (let i = 0; i < piece.text.length; i++) dataView.setUint16(i * 2, piece.text.charCodeAt(i), true);
      body.push(bytes);
      offset += bytes.length;
    } else {
      const bytes = new Uint8Array(piece.text.length);
      for (let i = 0; i < piece.text.length; i++) bytes[i] = piece.text.charCodeAt(i) & 0xff;
      body.push(bytes);
      offset += bytes.length;
    }
  }
  const textBytes = new Uint8Array(body.reduce((sum, part) => sum + part.length, 0));
  let at = 0;
  for (const part of body) {
    textBytes.set(part, at);
    at += part.length;
  }
  const word = new Uint8Array(FIB_SIZE + textBytes.length);
  word.set(fib, 0);
  word.set(textBytes, FIB_SIZE);
  // fcMin and ccpText describe the simple case.
  const totalCharacters = pieces.reduce((sum, piece) => sum + piece.text.length, 0);
  const wordView = new DataView(word.buffer);
  wordView.setUint32(0x18, FIB_SIZE, true);
  wordView.setUint32(0x4c, totalCharacters, true);

  let table = new Uint8Array(0);
  if (opts.complex) {
    const clx = buildClx(pieces, pcds);
    table = new Uint8Array(clx.length);
    table.set(clx, 0);
    wordView.setUint32(0x01a2, 0, true);
    wordView.setUint32(0x01a6, clx.length, true);
  } else {
    wordView.setUint32(0x01a6, 0, true);
  }
  return { word, table };
}

function buildClx(pieces, pcds) {
  const count = pieces.length;
  const plcLength = 4 * (count + 1) + 8 * count;
  const clxLength = 1 + 4 + plcLength;
  const out = new Uint8Array(clxLength);
  const view = new DataView(out.buffer);
  out[0] = 0x02;
  view.setUint32(1, plcLength, true);
  let cp = 0;
  for (let i = 0; i <= count; i++) {
    view.setUint32(5 + i * 4, cp, true);
    if (i < count) cp += pieces[i].text.length;
  }
  const pcdAt = 5 + (count + 1) * 4;
  for (let i = 0; i < count; i++) {
    const pcd = pcds[i];
    const fc = pcd.wide ? pcd.offset : ((pcd.offset << 1) | 0) + 0x40000000;
    view.setUint32(pcdAt + i * 8 + 2, pcd.wide ? pcd.offset : (pcd.offset << 1) | 0x40000000, true);
    void fc;
  }
  return out;
}

function makeDoc(pieces, options) {
  const { word, table } = buildDocStream(pieces, options);
  const streams = [{ name: "WordDocument", bytes: word }];
  if (table.length) streams.push({ name: "1Table", bytes: table });
  // Keep the streams in the big sector path by padding, as real files are.
  const padded = streams.map((stream) => ({
    name: stream.name,
    bytes: stream.bytes.length >= 4096 ? stream.bytes : padToSector(stream.bytes, 512),
    size: stream.bytes.length,
  }));
  const doc = buildCompound(padded);
  // The declared sizes must stay the real sizes, not the padding.
  return { bytes: doc, sizes: streams.map((s) => [s.name, s.bytes.length]) };
}

// ---------- the container ----------

const plain = makeDoc([{ text: "Hello compound world\rSecond line\r" }], {});
check("a built file looks like a compound document", isCompound(plain.bytes));
check("junk is not compound", isCompound(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8])) === false);
const compound = openCompound(plain.bytes);
check("the stream list has the document", compound.names().indexOf("WordDocument") !== -1, compound.names().join(","));
check("a missing stream is null", compound.stream("Nope") === null);

// ---------- plain text ----------

const plainModel = readDoc(plain.bytes);
const plainText = plainModel.body.map((block) => (block.runs || []).map((run) => run.text).join("")).join("\n");
check("plain text survives a round trip", plainText === "Hello compound world\nSecond line", JSON.stringify(plainText));
check("paragraphs counted", plainModel.paragraphCount, 2);

// ---------- a piece table with one 8 bit and one 16 bit piece ----------

// The 8 bit piece uses characters cp1252 really has; an en dash would be a
// bare control byte, which is the builder's limitation, not the reader's.
const pieces = [
  { text: "Premi\u00e8re piece, 8 bit. ", wide: false },
  { text: "Second piece \u2013 16 bit \u2603.", wide: true },
];
const complexDoc = makeDoc(pieces, { complex: true });
const complexModel = readDoc(complexDoc.bytes);
const complexText = complexModel.body.map((block) => (block.runs || []).map((run) => run.text).join("")).join("\n");
check("both pieces are read in order", complexText === "Premi\u00e8re piece, 8 bit. Second piece \u2013 16 bit \u2603.", JSON.stringify(complexText));
check("a cp1252 accent survives the 8 bit piece", complexText.indexOf("\u00e8") !== -1);
check("a 16 bit character survives", complexText.indexOf("\u2603") !== -1);

// ---------- control characters ----------

check("soft breaks become newlines", cleanText("one\u000btwo") === "one\ntwo");
check("cell marks become paragraph breaks", cleanText("cell\u0007next") === "cell\nnext");
check("field instructions are dropped and results kept", cleanText("total \u0013 PAGE \u0014 7 \u0015 end") === "total  7  end");
check("non breaking hyphens are kept as hyphens", cleanText("re\u001eentry") === "re-entry");
check("picture anchors are dropped", cleanText("before\u0001after") === "beforeafter");

// ---------- damaged input ----------

function reasonOf(fn) {
  try {
    fn();
    return null;
  } catch (err) {
    return err && err.message ? err.message : String(err);
  }
}

const truncated = plain.bytes.subarray(0, 600);
check("a truncated container is refused", Boolean(reasonOf(() => openCompound(truncated))), String(reasonOf(() => openCompound(truncated))));
check("truncated bytes are not compound", isCompound(new Uint8Array(8)) === false);

const badHeader = plain.bytes.slice();
badHeader[0] = 0x00;
check("a broken signature is refused", Boolean(reasonOf(() => openCompound(badHeader))));

const zeroFib = makeDoc([{ text: "text" }], {});
zeroFib.bytes[512] = 0;
zeroFib.bytes[513] = 0;
check("a wrong Word signature is refused", Boolean(reasonOf(() => readDoc(zeroFib.bytes))));

// A stream whose chain points at itself must not loop forever.
const looping = plain.bytes.slice();
const head = new DataView(looping.buffer, looping.byteOffset, looping.byteLength);
const dirSector = head.getUint32(48, true);
const dirOffset = 512 + dirSector * 512;
head.setUint32(dirOffset + 128 + 116, dirSector, true); // WordDocument starts at the directory
const started = Date.now();
const looped = openCompound(looping);
const read = looped.stream("WordDocument");
check("a looping chain stops", read === null || read.length <= 512 * 4096, "length " + (read ? read.length : "null"));
check("a looping chain is fast", Date.now() - started < 2000, (Date.now() - started) + " ms");

// A declared size larger than the file must not allocate it.
const greedy = plain.bytes.slice();
const greedyView = new DataView(greedy.buffer, greedy.byteOffset, greedy.byteLength);
greedyView.setBigUint64(dirOffset + 128 + 120, BigInt(0x7fffffff), true);
const greedyStart = Date.now();
const greedyRead = openCompound(greedy).stream("WordDocument");
check("a huge declared size is capped", !greedyRead || greedyRead.length <= 256 * 1024 * 1024);
check("a huge declared size is fast", Date.now() - greedyStart < 2000, (Date.now() - greedyStart) + " ms");

console.log(`doc-legacy: ${pass} pass, ${fail} fail`);
if (fail) process.exit(1);
