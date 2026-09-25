# Ideas, dependencies and sources

The renderers here are written from the file formats up. Where a technique or a
piece of code was learned from another source, it is listed below with its
licence. No code was copied from a project whose licence is incompatible with
the GPL.

## Dependencies

* **fflate** (MIT): ZIP reading and DEFLATE, the runtime ZIP and XML layer.
* **linkedom** (ISC, development only): the DOM used by the test harness, so
  the parsers and renderers run outside Obsidian.
* **mtx-decompressor** (MPL-2.0) by Christopher Van Ravenswaay: decompresses
  the MicroType Express payload of an EOT container into a TrueType face, which
  `src/pptx/fonts.js` registers through the FontFace API.
* **@jsquash/jxr** (Apache-2.0) by Jamie Sinclair, a WebAssembly build of
  **jxrlib** (BSD-2-Clause, Microsoft): the JPEG XR decoder in `src/media/jxr/`.
  The files are vendored unmodified with their notices beside them.

## Incorporated code

* **Preset shape geometry** (`src/pptx/presets.js`): the a:prstGeom silhouette
  formulas for roughly a hundred presets, adapted from the preset geometry
  module of Office Attachment Viewer (maic-for-office) by Aqu-Lab, Apache
  License 2.0. Changes: TypeScript types removed, module converted to
  CommonJS, and the connector helpers kept while the canvas renderer that used
  them is not part of this plugin. The formulas themselves follow the avLst
  defaults in the ECMA-376 presetShapeDefinitions annex.

## Format references

* **ECMA-376** (Office Open XML) and **ISO/IEC 26300** (OpenDocument) are the
  normative references for the parts, properties and geometry this plugin
  reads. The preset geometry table follows the ECMA-376 annex.
* The TIFF baseline, LZW and predictor rules, the EMF and WMF record layouts,
  the compound document sectors and the RLE and packed formats are implemented
  from their published specifications and pinned by tests in `test/`.

## Why a reference render can disagree

* Some converters scale pictures by their DPI metadata and clip them to the
  frame, so a picture the OOXML stretches to its frame can come out cropped.
* Some converters drop the alpha channel of a PNG and paint the transparent
  area black.
* Some shrink autofit text further than the file declares.
* One deck in this machine's library places its chart bars about 170 px from
  where the OOXML puts them in one converter.

When a reference disagrees with the file, the file wins, and the difference is
written down rather than copied.
