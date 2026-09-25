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

// Which controller opens which extension. The four byte signature is checked
// too, so a mislabelled file still lands on the right reader.

const FORMATS = {
  xlsx: { kind: "spreadsheet", label: "Workbook" },
  xlsm: { kind: "spreadsheet", label: "Workbook" },
  xltx: { kind: "spreadsheet", label: "Template" },
  xltm: { kind: "spreadsheet", label: "Template" },
  csv: { kind: "spreadsheet", label: "Comma separated" },
  tsv: { kind: "spreadsheet", label: "Tab separated" },
  ods: { kind: "spreadsheet", label: "OpenDocument spreadsheet" },
  docx: { kind: "document", label: "Document" },
  docm: { kind: "document", label: "Document" },
  dotx: { kind: "document", label: "Template" },
  dotm: { kind: "document", label: "Template" },
  odt: { kind: "document", label: "OpenDocument text" },
  rtf: { kind: "document", label: "Rich text" },
  pptx: { kind: "presentation", label: "Presentation" },
  pptm: { kind: "presentation", label: "Presentation" },
  ppsx: { kind: "presentation", label: "Slide show" },
  ppsm: { kind: "presentation", label: "Slide show" },
  potx: { kind: "presentation", label: "Template" },
  potm: { kind: "presentation", label: "Template" },
  odp: { kind: "presentation", label: "OpenDocument presentation" },
  // Word 97 documents are read as text: paragraphs and words, no formatting.
  doc: { kind: "document", label: "Word 97 document (text only)" },
};

// Extensions this plugin claims. .xls and .ppt remain out of scope: see the
// README for what is read and what is not.
const OFFICE_EXTENSIONS = Object.keys(FORMATS);

// A file that lies about its extension still opens correctly: the zip header
// and the presence of a known part decide.
function detectFormat(fileName, bytes) {
  const ext = extensionOf(fileName);
  const declared = FORMATS[ext] || null;
  const signature = sniff(bytes);
  if (signature === "zip") {
    const kind = detectZipKind(bytes);
    if (kind) return { kind, ext, label: declared ? declared.label : kind, trusted: true };
    if (declared && declared.kind !== "unsupported") return { kind: declared.kind, ext, label: declared.label, trusted: true };
    return { kind: "unsupported", ext, label: "Unknown zip package", trusted: false };
  }
  if (signature === "ole") {
    // .doc is read as text. A workbook or a deck stored the same way is not, so
    // those keep the explanation.
    if (ext === "doc" || ext === "dot") {
      return { kind: "document", ext: "doc", label: "Word 97 document (text only)", trusted: true };
    }
    return {
      kind: "unsupported",
      ext,
      label: "Legacy binary Office file",
      trusted: false,
      reason: "Old binary .xls and .ppt files store their content in a compound document this viewer reads for .doc only. Convert to .xlsx or .pptx, or open the file in the app that made it.",
    };
  }
  if (signature === "rtf") {
    return { kind: "document", ext: "rtf", label: "Rich text", trusted: true };
  }
  if (declared && (declared.kind === "spreadsheet" || declared.kind === "document")) {
    // csv, tsv and other text formats have no signature.
    if (!signature || signature === "text") return { kind: declared.kind, ext, label: declared.label, trusted: true };
  }
  if (declared) return { kind: declared.kind, ext, label: declared.label, trusted: true };
  return { kind: "unsupported", ext, label: "Unsupported", trusted: false };
}

function detectZipKind(bytes) {
  const names = readZipNames(bytes);
  if (!names) return null;
  if (names.indexOf("word/document.xml") !== -1) return "document";
  if (names.indexOf("ppt/presentation.xml") !== -1) return "presentation";
  if (names.indexOf("xl/workbook.xml") !== -1) return "spreadsheet";
  if (names.indexOf("content.xml") !== -1) {
    if (names.indexOf("META-INF/manifest.xml") !== -1) {
      // ODF: the manifest declares the document class.
      return "spreadsheet";
    }
  }
  return null;
}

function readZipNames(bytes) {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (data.length < 4 || data[0] !== 0x50 || data[1] !== 0x4b) return null;
  const decoder = new TextDecoder("utf-8");
  const names = [];
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const scanFrom = Math.max(0, data.length - 65557);
  let eocd = -1;
  for (let i = data.length - 22; i >= scanFrom; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd === -1) return null;
  const count = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  for (let i = 0; i < count; i++) {
    if (offset + 46 > data.length || view.getUint32(offset, true) !== 0x02014b50) break;
    const nameLen = view.getUint16(offset + 28, true);
    names.push(decoder.decode(data.subarray(offset + 46, offset + 46 + nameLen)));
    offset += 46 + nameLen + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
  }
  return names.length ? names : null;
}

function sniff(bytes) {
  if (!bytes || bytes.length < 4) return "text";
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (data.length >= 8 && data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4e && data[3] === 0x47) return "image";
  if (data[0] === 0x50 && data[1] === 0x4b && (data[2] === 0x03 || data[2] === 0x05 || data[2] === 0x07)) return "zip";
  if (data[0] === 0xd0 && data[1] === 0xcf && data[2] === 0x11 && data[3] === 0xe0) return "ole";
  if (data[0] === 0x7b && data[1] === 0x5c) return "rtf";
  if (data[0] === 0x25 && data[1] === 0x50 && data[2] === 0x44 && data[3] === 0x46) return "pdf";
  return "text";
}

function extensionOf(fileName) {
  const name = String(fileName || "");
  const dot = name.lastIndexOf(".");
  return dot === -1 ? "" : name.slice(dot + 1).toLowerCase();
}

module.exports = {
  FORMATS,
  OFFICE_EXTENSIONS,
  detectFormat,
  extensionOf,
};
