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
 * this program. If not, see <http://www.gnu.org/licenses/>.
 */

"use strict";

// Reads a zipped OOXML or ODF package.
//
// Office decks routinely ship 50 to 200 MB of media, and decompressing all of
// it up front costs seconds. Only the XML parts are decompressed at open time;
// media and any other part that a parser asks for later is inflated on demand
// and kept. Nothing is ever written back.

const { inflateSync } = require("fflate");
const { parseXml, attr, tagName } = require("./xml");

const INDEXED_PART = /\.(xml|rels|vml)$/i;

function dirOf(path) {
  const slash = path.lastIndexOf("/");
  return slash === -1 ? "" : path.slice(0, slash + 1);
}

function baseOf(path) {
  const slash = path.lastIndexOf("/");
  return slash === -1 ? path : path.slice(slash + 1);
}

// Resolves a relationship target against the folder its part lives in.
function resolvePath(baseDir, target) {
  if (!target) return null;
  if (target.startsWith("/")) return target.slice(1);
  const parts = (baseDir ? baseDir + "/" + target : target).split("/");
  const out = [];
  for (const part of parts) {
    if (part === "" || part === ".") continue;
    if (part === "..") out.pop();
    else out.push(part);
  }
  return out.join("/");
}

function relsPathFor(partPath) {
  return dirOf(partPath) + "_rels/" + baseOf(partPath) + ".rels";
}

function parseRels(xmlText) {
  const map = new Map();
  if (!xmlText) return map;
  let doc;
  try {
    doc = parseXml(xmlText);
  } catch (err) {
    return map;
  }
  const list = doc.documentElement.children || [];
  for (let i = 0; i < list.length; i++) {
    const el = list[i];
    if (tagName(el) !== "Relationship") continue;
    map.set(attr(el, "Id"), {
      type: attr(el, "Type") || "",
      target: attr(el, "Target"),
      mode: attr(el, "TargetMode") || "",
    });
  }
  return map;
}

// True when a relationship is of the given kind, judged by the final segment
// of its type URI.
function relTypeIs(rel, kind) {
  if (!rel || !rel.type) return false;
  return rel.type.slice(rel.type.lastIndexOf("/") + 1) === kind;
}

class Package {
  constructor(bytes) {
    this.archive = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    // The central directory is read and checked first, then only the XML parts
    // are inflated. A 200 MB deck therefore costs its directory plus its XML,
    // not a walk over every media entry, and a malformed or hostile archive is
    // rejected with a reason instead of failing somewhere inside a parser.
    const directory = readCentralDirectory(this.archive);
    this._entries = directory.entries;
    this._names = Array.from(directory.entries.keys()).sort();
    this._files = {};
    for (const [name, entry] of directory.entries) {
      if (!INDEXED_PART.test(name)) continue;
      const data = inflateEntry(this.archive, entry);
      if (data) this._files[name] = data;
    }
    this._decoder = new TextDecoder("utf-8");
    this._text = new Map();
    this._rels = new Map();
    this._contentTypes = null;
  }

  static open(input) {
    try {
      return new Package(input);
    } catch (err) {
      if (err && err.zipReason) throw new Error("This file is not a readable ZIP package (" + err.zipReason + ").");
      throw new Error("This file is not a readable ZIP package (" + err.message + ").");
    }
  }

  has(path) {
    return Boolean(path && this._names.indexOf(path) !== -1);
  }

  // The uncompressed size of a part, from the central directory. No inflation.
  partSize(path) {
    const entry = this._entries.get(path);
    return entry ? entry.uncompressed || 0 : 0;
  }

  // Every part with its uncompressed size, largest first. Used by the file
  // details panel to name the heavy parts without inflating anything.
  largestParts(limit) {
    const out = [];
    for (const [name, entry] of this._entries) out.push([name, entry.uncompressed || 0]);
    out.sort((a, b) => b[1] - a[1]);
    return out.slice(0, limit || 5);
  }

  list() {
    return this._names.slice();
  }

  // Part bytes, decompressing on first use.
  bytes(path) {
    if (!path) return null;
    const held = this._files[path];
    if (held && held.length) return held;
    const entry = this._entries.get(path);
    if (!entry) return null;
    try {
      const data = inflateEntry(this.archive, entry);
      if (data) this._files[path] = data;
      return data;
    } catch (err) {
      return null;
    }
  }

  text(path) {
    if (!path) return null;
    if (this._text.has(path)) return this._text.get(path);
    const data = this.bytes(path);
    let out = data ? this._decoder.decode(data) : null;
    // No Office part declares a document type or an entity, and a parser fed a
    // billion laughs payload stops responding. The part is refused instead.
    if (out && (out.indexOf("<!ENTITY") !== -1 || out.indexOf("<!DOCTYPE") !== -1)) out = null;
    this._text.set(path, out);
    return out;
  }

  // Parsed XML for a part, or null when the part is missing or malformed.
  xml(path) {
    const text = this.text(path);
    if (text === null) return null;
    if (text.length > 64 * 1024 * 1024) return null;
    try {
      return parseXml(text);
    } catch (err) {
      return null;
    }
  }

  rels(partPath) {
    const relsPath = relsPathFor(partPath);
    if (this._rels.has(relsPath)) return this._rels.get(relsPath);
    const map = parseRels(this.text(relsPath));
    this._rels.set(relsPath, map);
    return map;
  }

  // Resolves a part-relative relationship to a package path.
  resolve(partPath, rel) {
    if (!rel || !rel.target) return null;
    if (rel.mode === "External") return rel.target;
    return resolvePath(dirOf(partPath), rel.target);
  }

  // Convenience: the target path of a relationship id on a part.
  relTarget(partPath, relId) {
    const rel = this.rels(partPath).get(relId);
    return rel ? this.resolve(partPath, rel) : null;
  }

  // The default content type for an extension, from [Content_Types].xml.
  contentTypes() {
    if (this._contentTypes) return this._contentTypes;
    const out = { defaults: new Map(), overrides: new Map() };
    const doc = this.xml("[Content_Types].xml");
    if (doc) {
      const list = doc.documentElement.children || [];
      for (let i = 0; i < list.length; i++) {
        const el = list[i];
        const tag = tagName(el);
        if (tag === "Default") {
          out.defaults.set(String(attr(el, "Extension") || "").toLowerCase(), attr(el, "ContentType") || "");
        } else if (tag === "Override") {
          out.overrides.set(String(attr(el, "PartName") || "").replace(/^\//, ""), attr(el, "ContentType") || "");
        }
      }
    }
    this._contentTypes = out;
    return out;
  }

  // Best known MIME type for a part, from the content types table then extension.
  mimeOf(path) {
    if (!path) return "application/octet-stream";
    const types = this.contentTypes();
    if (types.overrides.has(path)) return types.overrides.get(path);
    const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
    if (types.defaults.has(ext)) return types.defaults.get(ext);
    return MIME_BY_EXT[ext] || "application/octet-stream";
  }

  // The main document part, found through _rels/.rels.
  mainPart(kind) {
    const rels = parseRels(this.text("_rels/.rels"));
    for (const rel of rels.values()) {
      if (kind && relTypeIs(rel, kind)) return resolvePath("", rel.target);
      if (!kind && relTypeIs(rel, "officeDocument")) return resolvePath("", rel.target);
    }
    if (kind === "content") return "content.xml";
    return null;
  }
}

// ---------- ZIP directory ----------
//
// A zipped Office file is untrusted input: it can be truncated, encrypted,
// built from several disks, use 64 bit sizes, or claim a part is a gigabyte of
// zeros. The directory is walked once, with caps, so every failure below has a
// readable reason attached. Nothing here writes anything.

const ZIP_EOCD = 0x06054b50;
const ZIP_EOCD64 = 0x06064b50;
const ZIP_LOCATOR64 = 0x07064b50;
const ZIP_CENTRAL = 0x02014b50;
const ZIP_LOCAL = 0x04034b50;
const ZIP_MAX_ENTRIES = 20000;
const ZIP_MAX_PART = 512 * 1024 * 1024;
const ZIP_MAX_TOTAL = 2 * 1024 * 1024 * 1024;
const ZIP_MAX_RATIO = 1000;

function zipError(reason) {
  const err = new Error(reason);
  err.zipReason = reason;
  return err;
}

function zip64Numbers(bytes, view, start, len, count) {
  const out = [];
  let at = start;
  const end = start + len;
  while (at + 4 <= end) {
    const id = view.getUint16(at, true);
    const size = view.getUint16(at + 2, true);
    if (id === 0x0001) {
      let cursor = at + 4;
      const limit = Math.min(end, at + 4 + size);
      for (let i = 0; i < count; i++) {
        if (cursor + 8 > limit) break;
        out.push(Number(view.getBigUint64(cursor, true)));
        cursor += 8;
      }
      return out;
    }
    at += 4 + size;
  }
  return out;
}

// Reads the central directory, validating as it goes. Returns the entries in
// archive order plus the uncompressed total.
function readCentralDirectory(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.byteLength < 22) throw zipError("the file is too small to hold a ZIP directory");
  let eocd = -1;
  const scanFrom = Math.max(0, view.byteLength - 65557);
  for (let at = view.byteLength - 22; at >= scanFrom; at--) {
    if (view.getUint32(at, true) === ZIP_EOCD) {
      if (at + 22 + view.getUint16(at + 20, true) <= view.byteLength) {
        eocd = at;
        break;
      }
    }
  }
  if (eocd === -1) throw zipError("no ZIP end record was found");
  const disk = view.getUint16(eocd + 4, true);
  const directoryDisk = view.getUint16(eocd + 6, true);
  const diskEntries = view.getUint16(eocd + 8, true);
  let count = view.getUint16(eocd + 10, true);
  let directorySize = view.getUint32(eocd + 12, true);
  let directoryOffset = view.getUint32(eocd + 16, true);
  if (disk !== 0 || directoryDisk !== 0 || diskEntries !== count) {
    throw zipError("multi part archives are not supported");
  }
  if (count === 0xffff || directoryOffset === 0xffffffff || directorySize === 0xffffffff) {
    const locator = eocd - 20;
    if (locator < 0 || view.getUint32(locator, true) !== ZIP_LOCATOR64) {
      throw zipError("the 64 bit ZIP directory is missing");
    }
    const eocd64 = Number(view.getBigUint64(locator + 8, true));
    if (eocd64 + 56 > view.byteLength || view.getUint32(eocd64, true) !== ZIP_EOCD64) {
      throw zipError("the 64 bit ZIP directory is damaged");
    }
    count = Number(view.getBigUint64(eocd64 + 32, true));
    directorySize = Number(view.getBigUint64(eocd64 + 40, true));
    directoryOffset = Number(view.getBigUint64(eocd64 + 48, true));
  }
  if (!count) throw zipError("the archive is empty");
  if (count > ZIP_MAX_ENTRIES) throw zipError("the archive holds too many entries");
  if (directoryOffset + directorySize > view.byteLength) throw zipError("the ZIP directory runs past the end of the file");
  const entries = new Map();
  const decoder = new TextDecoder("utf-8");
  let at = directoryOffset;
  let total = 0;
  for (let i = 0; i < count; i++) {
    if (at + 46 > view.byteLength || view.getUint32(at, true) !== ZIP_CENTRAL) {
      throw zipError("the ZIP directory is damaged at entry " + (i + 1));
    }
    const flags = view.getUint16(at + 8, true);
    const method = view.getUint16(at + 10, true);
    let compressed = view.getUint32(at + 20, true);
    let uncompressed = view.getUint32(at + 24, true);
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);
    let localOffset = view.getUint32(at + 42, true);
    const nameEnd = at + 46 + nameLength;
    if (nameEnd > view.byteLength) throw zipError("an entry name runs past the end of the file");
    const name = decoder.decode(bytes.subarray(at + 46, nameEnd));
    if (nameLength && (uncompressed === 0xffffffff || compressed === 0xffffffff || localOffset === 0xffffffff)) {
      // The 64 bit extra field carries only the values whose 32 bit slot is the
      // sentinel, in this order: uncompressed, compressed, local offset.
      const want = [];
      if (uncompressed === 0xffffffff) want.push("uncompressed");
      if (compressed === 0xffffffff) want.push("compressed");
      if (localOffset === 0xffffffff) want.push("offset");
      const values = zip64Numbers(bytes, view, nameEnd, extraLength, want.length);
      const valueOf = (key) => values[want.indexOf(key)];
      if (want.indexOf("uncompressed") !== -1 && valueOf("uncompressed") != null) uncompressed = valueOf("uncompressed");
      if (want.indexOf("compressed") !== -1 && valueOf("compressed") != null) compressed = valueOf("compressed");
      if (want.indexOf("offset") !== -1 && valueOf("offset") != null) localOffset = valueOf("offset");
    }
    at = nameEnd + extraLength + commentLength;
    if (!name || name.endsWith("/")) continue;
    if (flags & 0x1) throw zipError("the archive is encrypted");
    if (name.indexOf("\\") !== -1 || name.startsWith("/") || name.split("/").indexOf("..") !== -1) {
      throw zipError("an entry name is unsafe: " + name);
    }
    if (uncompressed > ZIP_MAX_PART) throw zipError("a part claims to be implausibly large");
    if (compressed > 0 && uncompressed > 10 * 1024 * 1024 && uncompressed / compressed > ZIP_MAX_RATIO) {
      throw zipError("an entry looks like a decompression bomb: " + name);
    }
    total += uncompressed;
    if (total > ZIP_MAX_TOTAL) throw zipError("the archive expands to more than the viewer will hold");
    if (localOffset + 30 > view.byteLength || view.getUint32(localOffset, true) !== ZIP_LOCAL) {
      throw zipError("the data for an entry is missing: " + name);
    }
    const localNameLength = view.getUint16(localOffset + 26, true);
    const localExtraLength = view.getUint16(localOffset + 28, true);
    const dataOffset = localOffset + 30 + localNameLength + localExtraLength;
    if (dataOffset + compressed > view.byteLength) throw zipError("an entry runs past the end of the file: " + name);
    if (method !== 0 && method !== 8) continue;
    entries.set(name, { method, compressed, uncompressed, dataOffset });
  }
  return { entries, total };
}

// Inflates one entry from its directory record. Method 0 is stored, method 8 is
// raw deflate; anything else was skipped when the directory was read.
function inflateEntry(bytes, entry) {
  const data = bytes.subarray(entry.dataOffset, entry.dataOffset + entry.compressed);
  if (entry.method === 0) return entry.uncompressed ? data.subarray(0, entry.uncompressed) : data;
  const out = entry.uncompressed ? new Uint8Array(entry.uncompressed) : undefined;
  return inflateSync(data, out ? { out } : undefined);
}

const MIME_BY_EXT = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  svg: "image/svg+xml",
  webp: "image/webp",
  tif: "image/tiff",
  tiff: "image/tiff",
  wdp: "image/vnd.ms-photo",
  hdp: "image/vnd.ms-photo",
  jxr: "image/jxr",
  emf: "image/emf",
  wmf: "image/wmf",
  mp3: "audio/mpeg",
  m4a: "audio/mp4",
  wav: "audio/wav",
  mp4: "video/mp4",
  avi: "video/x-msvideo",
  mov: "video/quicktime",
  wmv: "video/x-ms-wmv",
};

module.exports = {
  Package,
  parseRels,
  relsPathFor,
  resolvePath,
  dirOf,
  baseOf,
  relTypeIs,
  MIME_BY_EXT,
};
