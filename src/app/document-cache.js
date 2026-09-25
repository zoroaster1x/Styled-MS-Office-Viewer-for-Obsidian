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

const { releaseEmbeddedFonts } = require("../pptx/fonts");

// A small LRU of parsed documents. Reopening the file you were just looking at,
// or flipping back and forth between two, should not parse either of them
// again. Entries are keyed by path, size and modification time, so a file that
// changed on disk never serves stale content, and the vault's own modify event
// drops the entry as well.
//
// A cached model owns its media object URLs, so evicting an entry must revoke
// them: that is what releaseModel is for, and why the view asks before
// releasing a model it is done with.

class DocumentCache {
  constructor(opts) {
    const options = opts || {};
    this.limit = Math.max(1, options.limit || 3);
    // Media is the heavy part of a model, so a huge deck stays out of the
    // cache: it is cheaper to parse again than to hold a hundred megabytes of
    // decoded pictures for a file the reader may never return to.
    this.maxBytes = options.maxBytes || 32 * 1024 * 1024;
    this.entries = new Map();
  }

  static keyFor(path, stat) {
    if (!stat) return String(path);
    return path + "\u0000" + (stat.size || 0) + "\u0000" + (stat.mtime || 0);
  }

  // The entry for a key, or null. Reading it counts as a use, so it moves to
  // the back of the queue.
  lookup(key) {
    const hit = this.entries.get(key);
    if (!hit) return null;
    this.entries.delete(key);
    this.entries.set(key, hit);
    return hit;
  }

  // Stores a parsed model. Returns false when the document is too large to
  // keep, in which case the caller still owns it.
  store(key, value, size) {
    if (size && size > this.maxBytes) return false;
    const existing = this.entries.get(key);
    if (existing && existing.model !== value.model) releaseModel(existing.model);
    this.entries.delete(key);
    this.entries.set(key, value);
    this.trim();
    return true;
  }

  // Drops every entry for a path, whatever its size and time were. Used by the
  // vault's modify, delete and rename events.
  dropPath(path) {
    const prefix = path + "\u0000";
    for (const key of Array.from(this.entries.keys())) {
      if (key === path || key.indexOf(prefix) === 0) {
        const entry = this.entries.get(key);
        this.entries.delete(key);
        releaseModel(entry && entry.model);
      }
    }
  }

  clear() {
    for (const entry of this.entries.values()) releaseModel(entry.model);
    this.entries.clear();
  }

  trim() {
    while (this.entries.size > this.limit) {
      const oldest = this.entries.keys().next().value;
      const entry = this.entries.get(oldest);
      this.entries.delete(oldest);
      releaseModel(entry && entry.model);
    }
  }
}

// Revokes the object URLs a model owns, if it still holds a media cache.
function releaseModel(model) {
  const cache = model && model.mediaCache;
  if (cache && typeof cache.release === "function") cache.release();
  if (model) {
    releaseEmbeddedFonts(model);
    model.__ovCached = false;
  }
}

module.exports = { DocumentCache, releaseModel };
