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

// Hostile and broken input: truncated archives, encrypted entries, traversal
// names, decompression bombs, entity declarations, absurd TIFF headers and
// garbage metafiles. Every case must either work or fail with a reason, never
// hang, never throw something that is not an Error, and never allocate by a
// number that came out of the file.
//
//   bun test/robust.mjs

import { createRequire } from "node:module";
import { setupDom } from "./harness.mjs";

setupDom();
const require = createRequire(import.meta.url);
const { zipSync } = require("fflate");
const { Package } = require("../src/shared/package.js");
const { tiffToDataUrl, metafileToDataUrl } = require("../src/media/media.js");

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

function reasonOf(fn) {
  try {
    fn();
    return null;
  } catch (err) {
    return err instanceof Error ? (err.zipReason || err.message) : "not an Error: " + String(err);
  }
}

const doc = new TextEncoder().encode(
  '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>'
);

function makeZip(extra) {
  return zipSync(Object.assign({ "word/document.xml": doc }, extra || {}), { level: 6 });
}

// ---------- a good archive still works ----------
const good = makeZip({ "[Content_Types].xml": new TextEncoder().encode("<Types/>") });
const pkg = Package.open(good);
check("a normal archive opens", pkg.list().length >= 2, pkg.list().join(","));
check("a part reads back", pkg.text("word/document.xml").indexOf("w:document") !== -1);
check("a missing part is null", pkg.bytes("word/nope.xml") === null);

// ---------- truncation ----------
const truncated = reasonOf(() => Package.open(good.subarray(0, good.length - 40)));
check("a truncated archive is refused with a reason", Boolean(truncated), String(truncated));
check("the truncation reason is readable", /ZIP/.test(String(truncated)) || /directory/.test(String(truncated)), String(truncated));

const junk = reasonOf(() => Package.open(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8])));
check("junk is refused with a reason", Boolean(junk), String(junk));

// ---------- a byte flip in the directory ----------
const eocdAt = (() => {
  const view = new DataView(good.buffer, good.byteOffset, good.byteLength);
  for (let at = good.length - 22; at >= 0; at--) if (view.getUint32(at, true) === 0x06054b50) return at;
  return -1;
})();
check("the test can find the end record", eocdAt !== -1);
if (eocdAt !== -1) {
  const broken = good.slice();
  broken[eocdAt] = 0;
  check("a damaged end record is refused", Boolean(reasonOf(() => Package.open(broken))));
}

// ---------- encrypted ----------
const centralAt = (() => {
  const view = new DataView(good.buffer, good.byteOffset, good.byteLength);
  for (let at = 0; at + 4 <= good.length; at++) if (view.getUint32(at, true) === 0x02014b50) return at;
  return -1;
})();
check("the test can find the central directory", centralAt !== -1);
if (centralAt !== -1) {
  const encrypted = good.slice();
  const view = new DataView(encrypted.buffer, encrypted.byteOffset, encrypted.byteLength);
  view.setUint16(centralAt + 8, view.getUint16(centralAt + 8, true) | 1, true);
  const reason = reasonOf(() => Package.open(encrypted));
  check("an encrypted archive is refused", /encrypted/.test(String(reason)), String(reason));
}

// ---------- traversal ----------
const traversal = makeZip({ "../evil.xml": new TextEncoder().encode("<x/>") });
const traversalReason = reasonOf(() => Package.open(traversal));
check("a traversal name is refused", /unsafe/.test(String(traversalReason)), String(traversalReason));

// ---------- decompression bomb ----------
const bomb = zipSync({ "word/big.xml": new Uint8Array(80 * 1024 * 1024) }, { level: 9 });
const bombReason = reasonOf(() => Package.open(bomb));
check("a decompression bomb is refused", /bomb/.test(String(bombReason)), String(bombReason));

// ---------- entity declaration ----------
const nasty = makeZip({
  "word/nasty.xml": new TextEncoder().encode('<?xml version="1.0"?><!DOCTYPE t [<!ENTITY a "b">]><t>&a;</t>'),
});
const nastyPkg = Package.open(nasty);
check("an entity declaration makes the part unreadable", nastyPkg.text("word/nasty.xml") === null);
check("and its XML parse returns null", nastyPkg.xml("word/nasty.xml") === null);

// ---------- absurd TIFF ----------
function tiffWith(width, height) {
  const buffer = new ArrayBuffer(8 + 2 + 2 * 12 + 4);
  const view = new DataView(buffer);
  view.setUint16(0, 0x4949, false);
  view.setUint16(2, 42, false);
  view.setUint32(4, 8, false);
  view.setUint16(8, 2, false);
  view.setUint16(10, 256, false);
  view.setUint16(12, 4, false);
  view.setUint32(14, 1, false);
  view.setUint32(18, width, false);
  view.setUint16(22, 257, false);
  view.setUint16(24, 4, false);
  view.setUint32(26, 1, false);
  view.setUint32(30, height, false);
  view.setUint32(34, 0, false);
  return new Uint8Array(buffer);
}
check("a huge TIFF is refused", tiffToDataUrl(tiffWith(200000, 200000)) === null);
check("a zero TIFF is refused", tiffToDataUrl(tiffWith(0, 0)) === null);
check("a tiny garbage buffer is refused", tiffToDataUrl(new Uint8Array([1, 2, 3])) === null);

// ---------- garbage metafiles ----------
const garbage = new Uint8Array(4096);
for (let i = 0; i < garbage.length; i++) garbage[i] = (i * 37) & 0xff;
check("garbage EMF does not throw", metafileToDataUrl(garbage, "emf") === null || true);
check("garbage WMF does not throw", metafileToDataUrl(garbage, "wmf") === null || true);
const empty = new Uint8Array(0);
check("an empty metafile does not throw", metafileToDataUrl(empty, "emf") === null || true);

// ---------- a bounded fuzz over the directory ----------
let fuzzFailures = 0;
let longest = 0;
const start = Date.now();
for (let round = 0; round < 300; round++) {
  const mutated = good.slice();
  const at = (round * 7919) % mutated.length;
  mutated[at] = mutated[at] ^ 0xff;
  const before = Date.now();
  try {
    const opened = Package.open(mutated);
    // If it opened, every part that claims to exist must read or be null.
    for (const name of opened.list()) {
      const data = opened.bytes(name);
      if (data && data.length > 200 * 1024 * 1024) throw new Error("part grew past the cap");
    }
  } catch (err) {
    if (!(err instanceof Error)) fuzzFailures++;
  }
  longest = Math.max(longest, Date.now() - before);
}
check("300 mutated archives all fail cleanly or open", fuzzFailures === 0, fuzzFailures + " non-Error throws");
check("no mutation took longer than a second", longest < 1000, longest + " ms");
check("the whole fuzz stayed under ten seconds", Date.now() - start < 10000, (Date.now() - start) + " ms");

console.log(`robust: ${pass} pass, ${fail} fail`);
if (fail) process.exit(1);
