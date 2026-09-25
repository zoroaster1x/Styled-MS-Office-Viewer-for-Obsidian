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

// A presentation can carry its own fonts in ppt/fonts/*.fntdata. They are EOT
// containers whose payload is MicroType Express compressed, so the text renders
// as the author saw it even when the host has neither Calibri nor the face.
// mtx-decompressor (MPL-2.0) turns the container into a plain TTF, which goes
// into the document font set through the FontFace API.

const { eotToTtf } = require("mtx-decompressor");

// The same bounds the reference implementation uses: a crafted package must not
// turn a font part into an unbounded decode.
const MAX_FACES = 16;
const MAX_INPUT_BYTES = 8 * 1024 * 1024;
const MAX_OUTPUT_BYTES = 16 * 1024 * 1024;

function canUseFonts() {
  return typeof FontFace === "function" && typeof document !== "undefined" && document.fonts
    && typeof document.fonts.add === "function";
}

// Registers every face the model asks for and resolves when they are ready to
// measure with. A face that will not decode is skipped, so the host fallback
// still draws the text.
function ensureEmbeddedFonts(model) {
  if (!model || model.__ovFontsReady || !model.embeddedFonts || !model.embeddedFonts.length) return Promise.resolve();
  if (!canUseFonts()) return Promise.resolve();
  model.__ovFontsReady = true;
  const faces = [];
  const jobs = [];
  for (const face of model.embeddedFonts.slice(0, MAX_FACES)) {
    try {
      const bytes = model.pkg && model.pkg.bytes(face.path);
      if (!bytes || !bytes.length || bytes.length > MAX_INPUT_BYTES) continue;
      const ttf = eotToTtf(bytes);
      if (!ttf || !ttf.length || ttf.length > MAX_OUTPUT_BYTES) continue;
      const fontFace = new FontFace(face.family, ttf, { weight: face.weight, style: face.style });
      document.fonts.add(fontFace);
      faces.push(fontFace);
      jobs.push(Promise.resolve(fontFace.load ? fontFace.load() : null).catch(() => {}));
    } catch (err) {
      // The host font stack covers this face.
    }
  }
  model.__ovFontFaces = faces;
  return Promise.all(jobs);
}

// A cached model owns its faces; dropping the model hands them back.
function releaseEmbeddedFonts(model) {
  if (!model || !model.__ovFontFaces) return;
  for (const face of model.__ovFontFaces) {
    try {
      document.fonts.delete(face);
    } catch (err) {
      // A host without a removable font set keeps the face; nothing else to do.
    }
  }
  model.__ovFontFaces = null;
  model.__ovFontsReady = false;
}

module.exports = { ensureEmbeddedFonts, releaseEmbeddedFonts, MAX_FACES };
