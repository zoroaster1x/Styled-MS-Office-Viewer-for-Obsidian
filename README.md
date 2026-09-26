# Styled MS Office Viewer

Read Microsoft Office and OpenDocument files inside Obsidian with the styling they were saved with. Workbooks keep their fills, fonts, borders, merged cells and frozen panes. Word documents are laid out as pages. Presentations open as slides with a thumbnail rail, a slideshow mode and speaker notes.

The plugin is **strictly read only**. It parses a file in memory and draws it; it never writes to your documents and has no save path at all.

**[Supported formats](#supported-formats)** · **[Install](#installing)** · **[How it works](#how-it-works)** · **[Settings](#settings)** · **[Benchmarks](#measured-against-real-documents)** · **[Known limits](#known-limits)** · **[Funding](#funding)** · **[License](#license)**

---

## Supported formats

| Family | Extensions | What you get |
|---|---|---|
| Excel | `.xlsx` `.xlsm` `.xltx` `.xltm` | Styled grid: solid and gradient fills, theme colours with tints, fonts, every border style, merged cells, frozen rows and columns, hidden rows and columns, row and column sizes, wrapped text, number formats including dates, times and elapsed time, autofilters with filter chips, row and column outline groups, conditional formatting (colour scales, data bars, cell and text rules, icon sets), sheet tabs, hyperlinks, search |
| Delimited | `.csv` `.tsv` | The same grid. The delimiter is sniffed (comma, tab, semicolon, pipe), quoting follows RFC 4180, and numbers become real numbers |
| OpenDocument | `.ods` | The same grid, read from `content.xml` and `styles.xml` |
| Word | `.docx` `.docm` `.dotx` `.dotm` | Real style cascade (document defaults, named styles, basedOn chains, direct formatting), headings, lists with real numbering and counters, tables with merged cells and per-cell borders and fills, inline and floating images, inline drawing groups (text boxes, pictures, preset silhouettes, freeform lines and connectors), page size and margins, headers and footers with live `PAGE` and `NUMPAGES`, footnotes and endnotes, hyperlinks, typeset OMML equations, character effects, theme colours |
| OpenDocument | `.odt` | Paragraphs, headings, lists, spans with character styles, tables, page setup |
| Rich text | `.rtf` | Character styles, paragraph alignment, spacing and indents, colours, fonts, tables |
| PowerPoint | `.pptx` `.pptm` `.ppsx` `.ppsm` `.potx` `.potm` | Slide canvas with shapes, groups, freeform paths, connectors with arrowheads, pictures with crop, tables, text with wrapping, bullets, indents and autofit, theme colours, gradients and patterns, shadows, embedded fonts, equations, SmartArt drawings, speaker notes, thumbnail rail, slide search, slideshow mode |
| OpenDocument | `.odp` | Frames with text and pictures |
| Word 97 | `.doc` | Text only: paragraphs and words read from the compound document, no formatting |

**Not supported:** `.xls` and `.ppt`. Their content lives in a legacy binary format this reader does not walk; opening one shows a clear message with a button to hand the file to the system viewer, and converting to `.xlsx` or `.pptx` is the fix.

---

## Installing

1. Copy `main.js`, `manifest.json`, `styles.css` and `versions.json` into `<Vault>/.obsidian/plugins/styled-ms-office-viewer/`.
2. In Obsidian, open **Settings → Community plugins** and turn Restricted mode off.
3. Reload the installed plugins, then enable **Styled MS Office Viewer**.

From a source checkout:

```bash
./install.sh /path/to/YourVault
```

builds the bundle and copies the four files into place.

---

## Read only, by design

There is no save path in this plugin. It reads bytes from the vault, parses them in memory and draws them. The only thing it writes is its own settings, which remember which sheet, filters and slide you were last on. If you want to edit a file, use the application that made it.

---

## How it works

The plugin is a small shared core with one reader and one renderer per format family. Everything runs in Obsidian's own pane; there is no companion application, no conversion step and no network access.

```
src/
  main.js                  plugin entry: settings, commands, view registration
  app/
    view.js                the FileView: toolbar, states, format routing
    format-router.js       extension and byte-signature detection
    document-cache.js      parsed models, keyed by path, size and mtime
    spreadsheet-controller.js
    document-controller.js
    presentation-controller.js
  shared/
    package.js             ZIP reader with validation, lazy parts, relationships
    xml.js                 namespace-tolerant XML traversal
    color.js               theme colours, tints, DrawingML transforms
    units.js               EMU, twips, half-points, ODF lengths
    math-dom.js            OMML equations, shared by Word and PowerPoint
  spreadsheet/             xlsx reader, style table, number formats, grid renderer
  docx/                    styles, numbering, document parser, page renderer
  pptx/                    presentation parser, slide renderer, embedded fonts
  odf/                     ods, odt and odp readers
  rtf/                     RTF reader
  delimited/               csv and tsv
  media/                   image decoding: TIFF, JPEG XR (WebAssembly), EMF and WMF
```

The path from a file on disk to pixels on screen:

1. **Detection.** The extension picks a candidate reader, and the first bytes confirm it: the ZIP header for Office packages, the OLE compound header for Word 97. A mislabelled file still opens, and a file that lies about being a package is refused with an explanation.
2. **Package.** The ZIP central directory is read and validated first: entry caps, decoded size caps, compression ratio, encryption, path traversal and XML entities are all checked before anything is inflated. Only the XML parts are decompressed at open; media waits until a page or slide asks for it.
3. **Parse.** Each family has its own parser. OOXML and ODF are read as parts with relationships and a style table; RTF and delimited text are read directly; `.doc` walks the compound document and its piece table. Pagination for documents is planned from measured block heights.
4. **Render.** Renderers build plain DOM inside the pane and let the browser lay out text, so line breaking and font fallback match the host. Tables, images and shapes are positioned from the file's own geometry. A virtualized grid draws only the rows near the viewport; a deck draws one slide plus the thumbnails around the rail.
5. **Media.** Chromium handles PNG, JPEG, GIF, WebP, BMP and SVG directly. TIFF is decoded in JavaScript; JPEG XR goes through a bundled WebAssembly build of the libjxr codec; EMF and WMF are drawn by a record player. Results are cached per part, and pictures are served as blob URLs rather than base64 so the DOM stays small.
6. **Cache.** The last few parsed models stay in memory keyed by path, size and modification time, and the vault's modify, delete and rename events drop them. Every open takes a token, so a slow read that finishes after you moved to another file is discarded instead of mounting the wrong document.

Two design choices decide whether a 190 MB deck opens in half a second or five:

- **Media is inflated and transcoded on first use**, so a deck never pays for pictures it does not draw.
- **Only XML parts are inflated at open**, and they inflate once for the lifetime of the cached model.

---

## Using the viewer

A toolbar sits above every file. It carries search, zoom, and the actions that fit the open format: background and filters for a workbook, outline for a document, slide navigation and slideshow for a deck, plus file details, reload and the command shortcut. The search box finder arrows appear only while a search has matches. The zoom percentage sits with the magnifying glasses, and a touchpad pinch over the view zooms as well.

**Workbooks.** Drag column and row header edges to size them, double click an edge to fit it to its content, and the sizes are remembered per sheet. Filter chips above the grid hide whole groups in one click, and the full filter dropdown stays available on every filterable column. Conditional formatting and frozen panes are drawn as the file declares them.

**Documents.** The body is split into pages with a quiet page number in the bottom margin when the document has no footer of its own. Text and images are selectable and copyable. The outline panel lists the headings. The toolbar shows the page you are on beside the zoom, updated as you scroll, and changing the zoom keeps the same page in view.

**Presentations.** A slideshow button (and the **Start or stop the slideshow** command) fills the pane with the slide, takes the arrow keys, space, Home and End, and leaves on Escape. Slides carry their final state: the renderer draws every shape and never plays the animation tree, so nothing is hidden behind an effect. The file's own transitions and animations are reported in the details panel instead of being played. A deck that embeds its fonts renders them, so it keeps its typography on a machine that never had the face installed.

**File details.** The info button opens a panel with name, path, size and date, then everything the file says about itself: author, application, revision, timestamps, slide and sheet counts, words, characters and pages, languages, media and more. A file or deck can also carry its own context menu for copying the selection, searching, selecting all text, file details, reload, and opening in the system editor.

---

## Settings

| Section | Setting | What it does |
|---|---|---|
| Spreadsheets | Sheet background | Always white, like Excel, or follow the Obsidian theme for cells without a fill |
| | Show gridlines | Draw light gridlines in cells without a border of their own |
| | Show row and column headers | Show the A B C letters and the row numbers |
| | Default zoom | Starting zoom for new views |
| | Row render limit | How many rows to draw at once, for very large sheets |
| | Filter chips limit | Show chips for columns with at most this many values |
| | Minimum column width and row height | A floor under every sheet, off at 0 |
| Documents | Document zoom | Starting zoom for new views |
| | Page background | Always white, like Word, or follow the theme |
| | Show page numbers | Substitute PAGE and NUMPAGES where the document has them |
| Presentations | Slide zoom and fit | Fit the whole slide, or keep its saved size |
| | Show speaker notes | Show the notes pane under the slide |
| View | Show the counts inside the view | Move the word, character and page counts to the bottom left of the view instead of the app's status bar |
| | Pinch to zoom | A touchpad pinch over the view steps the zoom |
| | Invert pinch zoom | Swap the pinch direction |

---

## Measured against real documents

Every number below comes from the harness in `test/`, over a library of university lecture handouts and slide decks: mostly `.docx`, mostly `.pptx`, from a few kilobytes to 190 MB, with photographs, vector figures, speaker notes and forms.

| Check | Result |
|---|---|
| Text fidelity over 396 documents | 396 pass, 0 fail, 45.9 s |
| Parse and render over 397 documents (228 Word, 163 decks, 1 workbook, 5 Word 97) | 0 failures, 34.9 s with rendering |
| Largest deck, 191 MB, 51 slides | opens in about 600 ms |
| Reopening a cached document | no parse at all |
| Hostile input (truncated, encrypted, traversal, ZIP bomb, entity payloads, bounded fuzz) | every case refused with a reason, no hang |

Text fidelity means every text run in the source XML appears in the rendered output. The harness is `test/fidelity.mjs`, and it can be pointed at any folder through `.testenv`.

---

## Testing

The synthetic suites need nothing but the repository:

```bash
bun esbuild.config.mjs production   # build main.js
bun test/smoke.mjs                  # the bundle loads under an Obsidian stub
bun test/filters.mjs                # autofilter semantics and borders
bun test/copy.mjs                   # selection and clipboard output
bun test/virtual.mjs                # windowed rows, scrolling, search
bun test/sizing.mjs                 # column and row sizing
bun test/meta.mjs                   # file details, thumbnails, slide text
bun test/numfmt.mjs                 # serial dates, times, fractions, sections
bun test/cf.mjs                     # conditional formatting arithmetic
bun test/robust.mjs                 # damaged, encrypted and hostile archives
bun test/cache.mjs                  # the document cache and stale loads
bun test/doc-legacy.mjs             # the compound reader and Word 97 text
bun test/math.mjs                   # OMML equations, tree and layout
bun test/pptx-layout.mjs            # synthetic deck: scripts, connectors, groups, paths
bun test/docx-layout.mjs            # synthetic document: blank lines, pages, zoom, tables, fields
bun test/drawing.mjs                # Word drawing groups: transforms, dashes, pictures, text
bun test/tiff.mjs                   # the TIFF predictor, byte for byte
bun test/embed-font.mjs             # embedded EOT/MTX fonts decode and register
bun test/jxr.mjs                    # JPEG XR parts decode to real PNGs
bun test/privacy.mjs                # no personal paths in tracked files
```

The suites that read real documents take their paths from a gitignored `.testenv`, so no personal path is ever committed. Copy the example and fill in your own files:

```bash
cp .testenv.example .testenv
# set OV_TEST_XLSX, OV_TEST_DOCX, OV_TEST_PPTX, OV_TEST_JXR or OV_TEST_FOLDER
bun test/fidelity.mjs               # text completeness over a folder
bun test/stress.mjs --folder DIR    # parse and render every file, with timings
bun test/speed.mjs                  # timing for the largest configured deck
bun test/vault.mjs                  # open through the real view code
bun test/render-html.mjs FILE --out out.html   # a standalone page to inspect
python3 test/compare-reference.py   # diff against a reference conversion
```

Anything left unset in `.testenv` is skipped with a notice, and a path that is set but missing stops the run, so a typo is not mistaken for a pass. `test/harness.mjs` gives linkedom the Obsidian element extensions and a canvas stub, so the parsers and renderers run outside Obsidian with no mocks inside the plugin code.

---

## Known limits

These are honest gaps, not bugs waiting to be reported:

- **`.xls` and `.ppt`** are detected and explained, never half rendered. `.doc` is read as text: paragraphs and words out of the compound document, without fonts, colours, tables or images.
- **Real chart parts** show a labelled placeholder. The numbers live in an embedded workbook that is not plotted. The charts that appear in most documents are plain vector shapes and text, and those are drawn.
- **SmartArt** renders the drawing part PowerPoint stores beside the data, shapes and labels included. A diagram with no drawing part keeps a labelled placeholder.
- **Animations and transitions** are never played: a reader must not lose content to an effect. The details panel reports what the file declares.
- **Word's pagination is estimated.** Page breaks come from measured block heights, so a document can come out a page longer or shorter than Word makes it. Page size, margins, headers, footers and page numbers are real, and `PAGE` and `NUMPAGES` fields show the live numbers.
- **A floating Word shape or text box is drawn in flow** at its anchor paragraph. Its size, fill, outline and text are its own, but it does not push nearby text aside the way Word's absolute anchor does.
- **Legacy VML drawings** are read from the newer `wps` branch Word stores beside them; a file that carries only VML keeps its pictures, not its vector silhouettes.
- **Unusual OMML structures** fall back to their own text rather than disappearing.
- **Spreadsheet text rotation, images and comments** are not drawn. Rich text runs inside a cell keep the cell's own formatting.
- **A long sheet builds only the rows near the viewport.** Copy, search and filters still see every row.
- **Archives are validated before anything is decompressed**, which is why a hostile file gets a message instead of a hang.
- **Pattern fills** in slides are approximated with CSS gradients.
- **EMF and WMF metafiles** are drawn by a partial record player: paths, polygons, rectangles, ellipses, text, clipping and stock GDI objects. Bitmaps inside a metafile are not drawn.
- **Autofit in slides** estimates line counts from measured text, so dense slides can break lines slightly differently from PowerPoint.
- **Deeply nested groups with their own scaling and rotation** can be a few pixels out.
- **Shape effects** are approximated: a shadow keeps its distance and blur, a 3-D bevel is not modelled.

---

## Funding

If this plugin saves you time, consider supporting its development. Every contribution goes toward maintenance, new format coverage and the long tail of edge cases that make real documents painful to render.

**Monero (XMR):**

```
8BdxmQSniku4dBJXWPXeXvgjztmj5nmvWQqeCrVvCtYciusbAyo4rqrGCefTfQ4gGaVZmLN7VgLiYUYyBdYFEwHn1UWPjWs
```

> **Tip:** You can easily purchase Litecoin using [Cake Wallet](https://cakewallet.com/) and then, within the app, create a Monero wallet and exchange the Litecoin into it, pointed at the address above.

Crypto isn't your thing? Starring the repository, filing clear bug reports with a sample file, and telling other Obsidian users about the plugin all help just as much.

---

## License

GPL-3.0-or-later. Copyright (C) 2026 Zoroaster1x. See [`LICENSE`](LICENSE).

Bundled libraries and their licenses are listed in [`docs/references.md`](docs/references.md), in the notices beside the code, and in the source headers.
