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

// JPEG XR (wdp, hdp, jxr) decoding. Chromium cannot decode the format, so the
// libjxr codec compiled to WebAssembly draws it. The module is instantiated
// once, lazily, and a document without a JPEG XR picture never pays for it.
//
// The codec is @jsquash/jxr by Jamie Sinclair (Apache-2.0), which in turn wraps
// jxrlib (BSD-2-Clause). Both notices live in src/media/jxr/NOTICE.txt.

import moduleFactory from "./jxr/jxr_dec.js";
import wasmBinary from "./jxr/jxr_dec.wasm";

let instance = null;
let loading = null;

// Resolves true once the codec can decode; false on a host that refuses
// WebAssembly, where the caller keeps its named placeholder.
export function ensureJxrDecoder() {
  if (instance) return Promise.resolve(true);
  if (!loading) {
    loading = new Promise((resolve) => {
      let mod;
      try {
        // The emscripten shell asserts when it sees a Node global, and
        // Obsidian's desktop renderer exposes one. Hiding it for the factory
        // call selects the browser path, which is the environment this is.
        const hadProcess = typeof globalThis !== "undefined" && "process" in globalThis;
        const savedProcess = hadProcess ? globalThis.process : undefined;
        if (hadProcess) globalThis.process = undefined;
        try {
          // locateFile keeps the glue from building a URL out of import.meta,
          // which the bundled CommonJS output does not carry; the binary below
          // is what actually gets instantiated.
          mod = moduleFactory({ noInitialRun: true, wasmBinary, locateFile: (file) => file });
        } finally {
          if (hadProcess) globalThis.process = savedProcess;
        }
      } catch (err) {
        resolve(false);
        return;
      }
      Promise.resolve(mod).then(
        (ready) => {
          instance = ready;
          resolve(Boolean(ready && typeof ready.decode === "function"));
        },
        () => resolve(false)
      );
    });
  }
  return loading;
}

// Synchronous once ensureJxrDecoder has resolved. A decode failure returns
// null, which the callers turn into their own placeholder.
export function decodeJxrRgba(bytes) {
  if (!instance || !bytes || !bytes.length) return null;
  try {
    const image = instance.decode(bytes);
    if (!image || !image.width || !image.height || !image.data) return null;
    // The same bound as the TIFF path: a hostile header must not turn into a
    // multi gigabyte PNG encode.
    if (image.width * image.height > 50000000) return null;
    return { data: image.data, width: image.width, height: image.height };
  } catch (err) {
    return null;
  }
}
