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

"use strict";

// Compound File Binary reader, the container behind .doc, .xls and .ppt.
//
// Only what a text extractor needs: the directory and named streams. The format
// is a small file system whose every pointer is chosen by the file, so each
// sector reference is checked against the file length, each chain walk carries
// a visited set, and the stream sizes are capped. A damaged file produces a
// reason, never a loop or an allocation chosen by the file.

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const MAX_SECTORS = 1 << 22;
const MAX_STREAM = 256 * 1024 * 1024;
const END_OF_CHAIN = 0xfffffffe;
const FREE_SECTOR = 0xffffffff;

function isCompound(bytes) {
  if (!bytes || bytes.length < 8) return false;
  for (let i = 0; i < 8; i++) if (bytes[i] !== SIGNATURE[i]) return false;
  return true;
}

class CompoundError extends Error {}

function fail(reason) {
  throw new CompoundError(reason);
}

function openCompound(bytes) {
  if (!isCompound(bytes)) fail("not a compound document");
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.byteLength < 512) fail("the compound header is truncated");
  const sectorShift = view.getUint16(30, true);
  const miniShift = view.getUint16(32, true);
  if (sectorShift < 7 || sectorShift > 20) fail("implausible sector size");
  if (miniShift < 4 || miniShift > sectorShift) fail("implausible mini sector size");
  const sectorSize = 1 << sectorShift;
  const miniSize = 1 << miniShift;
  const totalSectors = Math.floor((view.byteLength - 512) / sectorSize);
  if (totalSectors > MAX_SECTORS) fail("too many sectors");

  const sectorOffset = (sector) => {
    const offset = 512 + sector * sectorSize;
    if (sector < 0 || offset + sectorSize > view.byteLength) fail("a sector points outside the file");
    return offset;
  };

  // The DIFAT lists the sectors that hold the FAT, first the 109 in the header
  // and then as many chained sectors as the file needs.
  const fatSectors = [];
  for (let i = 0; i < 109; i++) {
    const sector = view.getUint32(76 + i * 4, true);
    if (sector === FREE_SECTOR || sector === END_OF_CHAIN) continue;
    fatSectors.push(sector);
  }
  let difat = view.getUint32(68, true);
  let difatCount = view.getUint32(72, true);
  const seenDifat = new Set();
  while (difat !== END_OF_CHAIN && difat !== FREE_SECTOR && difatCount > 0) {
    if (seenDifat.has(difat)) fail("the DIFAT loops");
    seenDifat.add(difat);
    const offset = sectorOffset(difat);
    for (let i = 0; i < sectorSize / 4 - 1; i++) {
      const sector = view.getUint32(offset + i * 4, true);
      if (sector === FREE_SECTOR || sector === END_OF_CHAIN) continue;
      fatSectors.push(sector);
    }
    difat = view.getUint32(offset + sectorSize - 4, true);
    difatCount--;
  }
  if (!fatSectors.length) fail("the file has no allocation table");
  if (fatSectors.length > 4096) fail("the allocation table is implausibly large");

  // The FAT itself: an array of next-sector pointers, indexed by sector.
  const fat = [];
  for (const sector of fatSectors) {
    const offset = sectorOffset(sector);
    for (let i = 0; i < sectorSize / 4; i++) fat.push(view.getUint32(offset + i * 4, true));
  }

  // Reserved values: 0xfffffffa and up end a chain rather than naming a sector
  // (0xfffffffc is a DIFAT sector, 0xfffffffd a FAT sector, 0xfffffffe the end
  // and 0xffffffff free). Treating them as sectors is how a mini FAT chain used
  // to walk off the table.
  const chain = (start, limit) => {
    const out = [];
    const seen = new Set();
    let current = start;
    const cap = Math.min(limit || MAX_SECTORS, MAX_SECTORS);
    while (current < 0xfffffffa && out.length < cap) {
      if (current >= fat.length) fail("a chain leaves the allocation table");
      if (seen.has(current)) fail("a chain loops");
      seen.add(current);
      out.push(current);
      current = fat[current];
    }
    return out;
  };

  // Directory entries: name, type, start sector and size.
  const directorySector = view.getUint32(48, true);
  const directorySize = chain(directorySector, MAX_SECTORS).length * sectorSize;
  const entries = new Map();
  let root = null;
  for (const sector of chain(directorySector, 4096)) {
    const offset = sectorOffset(sector);
    for (let at = 0; at + 128 <= sectorSize; at += 128) {
      const nameLength = view.getUint16(offset + at + 64, true);
      if (!nameLength || nameLength > 64) continue;
      const type = view.getUint8(offset + at + 66);
      let name = "";
      for (let i = 0; i < nameLength - 2; i += 2) {
        name += String.fromCharCode(view.getUint16(offset + at + i, true));
      }
      const start = view.getUint32(offset + at + 116, true);
      const size = Number(view.getBigUint64(offset + at + 120, true));
      const entry = { name, type, start, size: Math.min(size, MAX_STREAM) };
      if (type === 5) root = entry;
      else if (type === 2 && name) entries.set(name, entry);
    }
  }
  void directorySize;
  if (!root) fail("the directory has no root entry");

  // Small streams live in the mini stream, which is the root entry's own
  // chain cut into mini sectors.
  const miniFat = [];
  const miniFatSectors = chain(view.getUint32(60, true), 4096);
  for (const sector of miniFatSectors) {
    const offset = sectorOffset(sector);
    for (let i = 0; i < sectorSize / 4; i++) miniFat.push(view.getUint32(offset + i * 4, true));
  }
  const miniCutoff = view.getUint32(56, true) || 4096;
  const miniStream = [];
  for (const sector of chain(root.start, 1 << 20)) miniStream.push(sectorOffset(sector));

  // Sector numbers, not byte offsets: every one goes through sectorOffset so
  // an out of range sector is caught rather than read as file bytes.
  const readSectors = (list, size) => {
    const out = new Uint8Array(Math.min(size, list.length * sectorSize));
    let at = 0;
    for (const sector of list) {
      if (at >= out.length) break;
      const offset = sectorOffset(sector);
      const take = Math.min(sectorSize, out.length - at);
      out.set(bytes.subarray(offset, offset + take), at);
      at += take;
    }
    return out;
  };

  const readMini = (entry) => {
    const out = new Uint8Array(entry.size);
    const seen = new Set();
    let current = entry.start;
    let at = 0;
    while (current !== END_OF_CHAIN && current !== FREE_SECTOR && at < out.length) {
      if (current >= miniFat.length || seen.has(current)) break;
      seen.add(current);
      const miniIndex = current;
      const byteOffset = miniIndex * miniSize;
      const containerIndex = Math.floor(byteOffset / sectorSize);
      if (containerIndex >= miniStream.length) break;
      const withinSector = byteOffset - containerIndex * sectorSize;
      const take = Math.min(miniSize, out.length - at, sectorSize - withinSector);
      out.set(bytes.subarray(miniStream[containerIndex] + withinSector, miniStream[containerIndex] + withinSector + take), at);
      at += take;
      current = miniFat[miniIndex];
    }
    return at === out.length ? out : out.subarray(0, at);
  };

  return {
    names() {
      return Array.from(entries.keys());
    },
    has(name) {
      return entries.has(name);
    },
    // Stream bytes, or null when it is missing or damaged beyond a clean read.
    stream(name) {
      const entry = entries.get(name);
      if (!entry) return null;
      try {
        if (entry.size < miniCutoff) return readMini(entry);
        return readSectors(chain(entry.start, MAX_SECTORS), entry.size);
      } catch (err) {
        return null;
      }
    },
  };
}

module.exports = { openCompound, isCompound, CompoundError };
