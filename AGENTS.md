# AGENTS.md

Rules for any agent working in this repository. Read this before touching anything.

## 1. What this is

Styled MS Office Viewer is an Obsidian plugin that reads `.xlsx`, `.xlsm`, `.csv`, `.tsv`, `.ods`, `.docx`, `.docm`, `.odt`, `.rtf`, `.pptx`, `.pptm` and `.odp` and draws them with their real styling.

## 2. Hard rules

* **Read only.** There is no save path for an opened document, and none may be added. The plugin reads bytes from the vault, parses them in memory, and draws them. The only file it writes is its own `data.json`, which remembers the active sheet, filters, slide position, and hand-set column widths and row heights.
* **Never write to the user's vault from a test.** Tests read documents, they never touch them.
* **No personal data in tracked files.** No home paths, user names, course codes, institution names, or document names. Test paths come from `.testenv`, which is gitignored. `test/privacy.mjs` fails the run when a tracked file looks like it carries one.
* **No em dashes** in text written here, including docs, commits and comments. Rewrite the sentence with a comma, a colon, a semicolon, or a full stop.
* **Verify before claiming done.** Run the suites that cover the change and report measured numbers, not intentions. Do not describe behaviour you did not run.

## 3. Build and run

`bun` is the runtime here, `node` may not be installed.

```bash
bun esbuild.config.mjs production   # builds main.js from src/main.js
bun test/smoke.mjs                  # loads the built bundle under an Obsidian stub
```

The bundle is CommonJS for Obsidian's renderer. `obsidian`, `electron` and node builtins are external. Dependencies are `fflate` for zip and (dev) `linkedom` for the test DOM.

## 4. Test suites

Synthetic, they need nothing configured:

| Suite | Covers |
|---|---|
| `test/smoke.mjs` | the built bundle loads, registers the view, the extensions, the commands and the settings tab, and paints its empty state |
| `test/filters.mjs` | autofilter semantics (values, blanks, custom operators, AND across columns) and that no border of a hidden row is drawn |
| `test/copy.mjs` | block selection, both clipboard flavours, quoted TSV, hidden rows skipped |
| `test/virtual.mjs` | windowed row rendering, scroll position after a rebuild, frozen band, selection tint, search outside the window |
| `test/sizing.mjs` | column and row drag, auto fit, remembered sizes, the minimum settings |
| `test/meta.mjs` | the file details panel per format, thumbnails, slide text not clipped |
| `test/numfmt.mjs` | Excel serial dates (including the 1900 leap bug), times, fractions, sections |
| `test/robust.mjs` | truncated, encrypted, traversal, zip bomb and entity payloads, plus a bounded fuzz |
| `test/cf.mjs` | conditional formatting: thresholds, data bars, colour scales, icon sets, top ten and averages |
| `test/doc-legacy.mjs` | the compound document reader and Word 97 text, on files the test builds byte by byte |
| `test/math.mjs` | OMML structures, their plain text, and the layout the renderer builds |
| `test/pptx-layout.mjs` | a synthetic deck pins the presentation rules: subscript baseline, a flipped connector, level line spacing, normAutofit reduction, layout date prototypes, slide number fields, a custom path with an arc, a scaled group's mapped geometry |
| `test/docx-layout.mjs` | a synthetic document pins the empty-paragraph line, no invented spacing, blank lines counting toward pages, and the zoom wiring |
| `test/tiff.mjs` | the TIFF predictor restarts at every row: a byte-built file proves RGB and alpha do not drift |
| `test/cache.mjs` | the parsed document cache, its invalidation and stale loads |
| `test/privacy.mjs` | no personal paths, emails or hostnames in tracked files |

File backed, they read paths from `.testenv` and skip with a notice when it is missing:

| Suite | Covers |
|---|---|
| `test/fidelity.mjs` | every text run of every document appears in the render, over a folder |
| `test/stress.mjs` | parse and render every document in a folder, with timings and failures |
| `test/speed.mjs` | open time for the largest configured document |
| `test/vault.mjs` | opens files through the real view code with a stubbed vault |
| `test/inspect.mjs` | one document, with structure counts and an optional HTML dump |
| `test/inventory.mjs` | prints what the viewer would draw as JSON, for one file |
| `test/jxr.mjs` | a package's JPEG XR parts decode to real PNGs through the WebAssembly codec |
| `test/embed-font.mjs` | a presentation's embedded EOT/MTX fonts decode to named TrueType faces and register; the fixture carries its OFL licence |
| `test/emf-probe.mjs` | records what the metafile player asks the canvas to do, so a vector logo can be checked without a rasteriser |
| `test/compare-reference.py` | converts the file with a headless office suite and diffs text and image placement with PyMuPDF |
| `test/bench.mjs` | unzip, XML and media cost for one package |
| `test/render-html.mjs` | renders a document to a standalone HTML file with pictures as data URLs, for looking at the output in a browser or screenshotting it next to the reference |
| `test/scan-docs.mjs` | reads every supported file in a folder and reports the metadata fields and content features inside, with the gaps ranked |

`test/harness.mjs` gives linkedom the Obsidian element extensions, pointer and
scroll events, and a canvas stub, so parsers and renderers run outside Obsidian
without mocks inside the plugin code.

## 5. Layout

```
src/main.js                 entry: settings, commands, view registration
src/app/view.js             the FileView: toolbar, states, routing, details panel
src/app/format-router.js    extension and byte signature detection
src/app/*-controller.js     one controller per format family
src/shared/package.js       zip reader: lazy part inflation, rels, content types
src/shared/xml.js           namespace tolerant XML traversal
src/shared/color.js         theme colours, tints, DrawingML transforms
src/shared/units.js         EMU, twips, half points, ODF lengths
src/shared/addr.js          cell addresses
src/spreadsheet/            xlsx reader, style table, number formats, grid renderer
src/docx/                   styles, numbering, parser, page renderer
src/pptx/                   presentation parser, slide renderer
src/pptx/presets.js         a:prstGeom silhouettes (ECMA-376 formulas, Apache-2.0 source, see docs/references.md)
src/odf/                    ods, odt, odp readers
src/rtf/                    RTF reader
src/delimited/              csv and tsv
src/media/media.js          TIFF, EMF and WMF transcodes
docs/references.md          where techniques came from and under which licence
```

## 6. Format knowledge worth keeping

These cost real time to find. They are all handled in the current code, so treat them as regressions to avoid, not as work to do.

### WordprocessingML

* A style keeps its text formatting in a top level `w:rPr`. It must merge into the paragraph's run defaults, in the order document defaults, style chain, paragraph mark `rPr`, then the run's own `rPr`. Precedence that drops the style is why headings lose their bold and size.
* Nested property bags merge per key. Assigning one bag over another discards everything the first one set.
* The run resolution cache key must include the paragraph context, or a body run's cached props leak into a heading run.
* Paragraph and character style ids are case sensitive (`Heading1`, not `heading1`).
* An empty paragraph still occupies one line at the paragraph mark's size. A form with 79 blank paragraphs depends on it.
* `w:docGrid` snaps lines only when its `type` is `lines` or `linesAndChars`. Word writes a `linePitch` into most documents with the default type, where the pitch is informational. Snapping there makes documents a page too long.
* A table row has a minimum height of one line of the default font, even when every cell is empty.
* Pictures arrive as `wp:inline` or `wp:anchor`, legacy ones as `w:pict` (VML), and a newer shape may sit inside `mc:AlternateContent`, where the `mc:Choice` branch is the one to read.
* A shape that holds a text box carries its paragraphs in `wps:txbx` / `w:txbxContent`. Dropping them loses text.
* `PAGE` and `NUMPAGES` fields need the page number at draw time: the viewer paginates, so substitute the live numbers rather than the saved ones.
* Footnote and endnote bodies live in their own parts; headers and footers are referenced from `w:sectPr`.
* A Word drawing group is `mc:AlternateContent/mc:Choice` with `Requires="wpg"`, holding a `wpg:wgp` of `wps:wsp` shapes, `pic:pic` pictures and one nested `wpg:grpSp` level. Dropping the group (or reading the VML `mc:Fallback`) leaves a blank gap where a diagram and its photographs belong. `src/docx/drawing.js` walks the group: `wpg:grpSpPr/a:xfrm` maps `chOff`/`chExt` onto `off`/`ext`, an `off`/`ext` of zero is an identity transform, shapes are `wps:spPr` geometry plus fill and line, and text lives in `wps:txbx/w:txbxContent`. A group or shape with no stated line colour still takes the `wps:style` `a:lnRef` colour; an `a:ln` that only names a width does not blank the outline.
* A standalone anchored shape (`mc:Choice Requires="wps"` with `wp:anchor`, no `wpg` around it) is a drawing too. A shape with geometry but no text box used to fall out of the parser and vanish, which is how an anchored orange trapezoid disappeared from a nose bridge.
* `a:custDash` is not `a:prstDash`: it lists dash and gap lengths as thousandths of the line width (`d="300000" sp="225000"` is 3 and 2.25 line widths). Reading only `prstDash` draws every custom-dashed line solid.
* DrawingML colour aliases `tx1`, `tx2`, `bg1`, `bg2` are not theme slot names; they resolve through the default colour map (`tx1` is `dk1`). Without them a line stated as `schemeClr tx1` falls through to a style colour and turns accent blue.
* A `PAGE` or `NUMPAGES` field arrives as `fldChar` begin/instr/separate/end spread across separate `w:r` runs, so its state must outlive one recursion of the inline renderer. The saved result can be stale (a footer in a 31-page manual carries 53), so the live number is substituted for the result text.
* A `vMerge` continuation cell must not be emitted at all. Writing it as an empty `<td>` on top of the restart cell's `rowspan` shifts every later cell one column right; the restart cell's rowspan is found by grid column, not by index in `row.cells`, because a `gridSpan` earlier in the row moves every later cell.
* `td.colSpan` and `td.rowSpan` are properties that some hosts do not reflect onto the attribute. Serialising the rendered HTML (the test renderer, a copy of the DOM) then silently loses every merged cell, so set the attributes.
* A section's header and footer references live in its own `w:sectPr`, and `parseBlockChildren` assigns a `block.section` index while walking, so the body up to a paragraph that carries a `sectPr` belongs to that section. Rendering every page with the final section's refs is wrong that way: a handbook with a cover section and a body section loses its running footer. `w:titlePg` means the first page of the section uses the `first` part and shows nothing when the file has none; `even` parts apply only when `word/settings.xml` carries `w:evenAndOddHeaders`, which most files do not.
* A page break is `w:br w:type="page"` inside a run. `splitAtPageBreaks` cuts the paragraph into a block that ends the page and the remainder, because the planner only breaks between blocks.
* `w:smartTag`, `w:ins` and `w:sdt` nest freely: a page number inside `w:sdt/w:sdtContent/w:sdt/w:sdtContent/w:r` is normal Word output. A one-level walk drops the text, so the inline walk recurses.
* `w:sym` names a glyph by private-use code point in a font (`char="F0B7" font="Symbol"`). The code point alone draws a blank box; the common Symbol, Wingdings and Webdings low-byte tables map to the Unicode characters they mean.
* Legacy VML (`w:pict` with `v:shape`, `v:rect`, `v:oval`, `v:line`, `v:polyline`, `v:group`) is what older files and the `mc:Fallback` branch carry. Coordinates inside a `v:group` are in the group's `coordsize` space (often twips), while a top level shape writes points in its style; `mso-wrap-style:square` reserves room in the line and `none` floats. `src/docx/vml.js` renders it into the same shape primitives as the DrawingML path.
* A table's `w:tblStyle` carries base borders and cell margins plus `w:tblStylePr` conditionals (firstRow, lastRow, firstCol, lastCol, band1Horz, band2Horz, band1Vert, band2Vert). Reading only the direct `w:tblPr` leaves a styled table borderless and unbanded.
* An embedded object (`w:object`) often keeps a preview image; when it has none, a labelled placeholder keeps the space and the reader informed.

### Equations (OMML)

* An equation is content, not decoration: `m:oMath` holds runs (`m:r`) and structures (`m:f` with `m:num`/`m:den`, `m:sSup`, `m:sSub`, `m:sSubSup`, `m:rad`, `m:d`, `m:nary`, `m:func`, `m:limLow`/`m:limUpp`, `m:acc`, `m:bar`, `m:m`, `m:eqArr`). Skipping `m:oMath` in the paragraph loop drops the numbers and the operators with it, which is exactly what happened: six documents lost their equations entirely.
* `src/docx/math.js` parses those structures into a small tree, `mathText` turns the tree back into plain text (search, word counts, the fidelity check), and the renderer builds fractions as stacked boxes, scripts as scripts, radicals with a sign and a bar, and matrices as a small table.
* A text run inside an equation is italic unless `m:nor` marks it: Word writes operators and units that way. Rendering `+` and `DS` in italics looks wrong next to proper maths.

### PresentationML

* `<p:bg>` lives inside `<p:cSld>`, not directly under the part root, on slides, layouts and masters alike. Searching the root finds nothing and every background is white.
* A group keeps its transform in `p:grpSpPr/a:xfrm`; every other shape keeps it in `p:spPr/a:xfrm`. Reading `spPr` for a group finds nothing, the group falls into the zero sized box path, and every child lands at `chOff` instead of the group's own box.
* Shape order in `p:spTree` is the z-order. Collecting the children by tag type (all shapes, then all pictures) draws every picture behind every shape, so an arrow that should cross a picture disappears under it.
* `<a:xfrm>` carries its offsets in `<a:off>` and `<a:ext>` children, not as attributes.
* A group with `off(0,0)` and `ext(0,0)` is an identity transform: its box is `chOff` / `chExt` and its children keep their own coordinates. Taking the zero extent seriously collapses the whole group onto the origin.
* Group children are placed from `chOff` and scaled by `ext / chExt`. Nested groups compose.
* A text style resolves through the shape's own `lstStyle`, then the layout placeholder's, then the master placeholder's, and only then the master's `txStyles` set. A subtitle belongs to `otherStyle`, not `bodyStyle`, or it inherits the body bullet and indent. Text boxes belong to `otherStyle` too, whose indent is zero.
* `bodyPr wrap="none"` means the text never wraps. A word that does not fit must be allowed to spill; clipping it turns "PROBLEMS" into "PROBLE".
* Autofit compares the text against the shape with the *same* font stack the span is drawn with. Measuring with `"Calibri, sans-serif"` while the span uses Carlito makes the estimate disagree with the layout and shrinks text that fits. Search the scale by bisection rather than five percent steps, or a title lands one word too wide.
* `bodyPr` insets, wrap, vertical text, autofit and the vertical anchor inherit from the layout placeholder then the master's. The anchor used to be dropped because a master that centres its title disagreed with the reference renderers; the file is the authority and PowerPoint inherits it. Centring only ever moves a text block up, so the old worry about a line hiding behind a picture that follows could not happen; the measured effect on that title was 11 px up, which is what the file asks for.
* A list that ends with an empty paragraph prints no marker for it in PowerPoint or the reference renderers.
* `a:rPr baseline` is 1000ths of a percent of the font size. The glyphs are also drawn at 58% of the run's size, so the CSS shift is `baseline / 58` em (the usual `-25000` is -0.43em). Writing `baseline + "em"` drops a subscript 25em below the box and it disappears while its width still reserves space.
* A flipped connector bakes the flip into its SVG coordinates; the element must NOT also carry the generic CSS flip, or every arrow is mirrored vertically.
* Paragraph levels carry paragraph properties as well as run properties. `lnSpc` and `spcBef` from the layout placeholder or the master's txStyles must merge into a paragraph that does not state its own, or titles lose the master's 90% line spacing and body lists their spacing before.
* `normAutofit lnSpcReduction` reduces line spacing by that fraction for every paragraph of the shape, including the ones with no explicit spacing.
* Date, footer and slide-number placeholders on layouts and masters are prototypes for the slide's own content. A slide shows one only when it carries that placeholder itself. The deck's own PowerPoint thumbnail and the reference renderers all draw an uninstantiated date and slide number nowhere; drawing the layout's put a ghost date on every page.
* `src/shared/metadata.js` reads everything a package says about itself. When a panel shows metadata, add the field there rather than in a controller, so every format gets it.
* `test/scan-docs.mjs` is the way to decide what to build next: it inventories a folder's metadata, languages, comment authors, media types and content features, and ranks the gaps by prevalence. A capability that looks important can turn out to appear once in four hundred documents, and the scanner says so before the work starts.
* A bullet marker needs the paragraph's level font size and colour, since it has no run of its own. Sizing it from the shape default gives level 1 and level 2 the same dot, too small beside 28pt text; take the size from the paragraph's first run (the level's resolved size). The marker box takes the hanging indent's width so the text after it starts at marL, and a font relative min-width must not stretch it.
* A shape or a background filled with a picture uses `a:blipFill`, not `p:pic`. Ignoring it drops every pasted screenshot in the deck.
* `<a:tblGrid>` holds `<a:gridCol>`; tables live in a `graphicFrame`.
* `<p:sldId>` carries both `id` and `r:id`. The relationship is the namespaced one.
* A picture with a vector version has `<asvg:svgBlip>` beside the raster `<a:blip>`. Draw the SVG.
* A chart in a slide is often just vector rectangles and text boxes. Real chart parts get a labelled placeholder.
* `mc:AlternateContent` in a shape tree holds a newer construct in `mc:Choice` and an older picture of it in `mc:Fallback`. Reading the Fallback (or skipping both) is how a slide's equations became empty rectangles: the Choice carries the real thing.
* An equation in a slide is an `a14:m` wrapper around the same OMML Word writes. Parse it with `src/docx/math.js` and draw it with `src/shared/math-dom.js`, the module both renderers share; add `mathText` to the paragraph text or search and the fidelity sweep cannot see it.
* `a:custGeom` paths have their own coordinate space (`a:path w/h`), so scale them to the shape. An `arcTo` centre follows from the current point and its angles, so compute the endpoint instead of guessing. A solid fill draws the silhouette as an SVG path; a clipped rectangle cannot show a concave outline. Picture and gradient fills clip the element instead.
* A horizontal or vertical connector has a zero extent on one axis, and an SVG with a zero width or height collapses and draws nothing. Give the viewport one line width on that axis and centre the line in it, or a sunburst of vertical connectors disappears. A line can also be a plain `p:sp` with the line preset, not only a `p:cxnSp`; check the geometry, or the shape draws only a border on a zero width box.
* A group scales its children by `ext / chExt` over `chOff / off`, and that belongs in the child geometry, not in a CSS transform. A group whose child space is 1000 units mapped to 2000000 EMU made a 0.0006 px child; a CSS scale turned the wrapper into a 348000 px element and Chrome stopped rasterising the SVGs inside it. Worse, the transform multiplied font sizes and stroke widths that PowerPoint keeps absolute. The renderer threads a map `{ax, ay, bx, by}` through `drawGroup`; `positionElement` and every shape-box reader apply it, while run sizes and line widths do not.
* CSS clamps a gradient stop that sits before its predecessor to that position, which collapses the rest of the list. A file can list stops unordered; sort them by position before building the gradient or the fill washes out.
* An equation is a sibling of the run, not a run: `m:oMath` sits directly inside `w:p`. A paragraph walker that only looks at `w:r` drops every equation silently, which is exactly what happened to six real documents. Any other content that is a sibling of the run has the same trap.
* OMML structures worth reading: `m:f` (num, den), `m:sSup`, `m:sSub`, `m:sSubSup`, `m:rad` (deg, degHide), `m:d` (begChr, endChr, sepChr), `m:nary` (chr, subHide, supHide), `m:func` (fName), `m:limLow`, `m:limUpp`, `m:acc` (accPr/chr), `m:bar` (pos), `m:m`/`m:mr`, `m:eqArr`. Text inside an equation is italic unless `m:nor` marks the run as an operator or a unit.
* An equation's plain text must reach the model text as well as the layout, or search, the word counts and the fidelity check cannot see it.
* The timing tree (`p:timing`) is never consulted, and no shape is ever gated on it: an animation must not be able to hide content from a reader. Slideshow mode is a view state on top of the same drawing.
* Slideshow mode hides the viewer's chrome only (toolbar, rail, notes, status) and takes the arrow keys, space, Home, End and Escape. The entrance between slides is a CSS animation; the file's own transition and animation are reported in the details, not played.
* A document or a deck can carry its own context menu: a capture phase `contextmenu` listener that stops propagation, then Obsidian's `Menu` with selection aware items. Without capturing, the app's plain text menu wins. The grid asks the spreadsheet controller for its items instead.
* SmartArt keeps the rendered shapes in a sibling part: the slide's relationships include one `diagramDrawing` per diagram (`ppt/diagrams/drawing*.xml`) whose `dsp:spTree` holds ordinary shapes with their text. Frames on a slide pair with those relationships in document order. Render the drawing and the SmartArt is real; skip it and all that is left is a placeholder.
* Shape silhouettes come from `a:prstGeom` plus `a:avLst` adjustment values. A fixed table of silhouettes with the ECMA-376 formulas lives in `src/pptx/presets.js`; without it an arrow is a rectangle.
* The slide number placeholder holds a glyph (`<a:fld type="slidenum">`), so the number is substituted when drawing.

### SpreadsheetML

* A border edge may only come from a neighbour that is actually drawn. Taking it from a hidden row or column leaves ghost lines under a filtered block.
* Autofilter semantics: a value set hides everything not in it, an empty set hides every data row and keeps the header, custom operators compare numbers, blanks get their own entry, and multiple columns combine with AND.
* Copy writes quoted fields when a value holds a tab, a newline or a quote, which is what Excel writes. An unquoted newline splits the row in the receiving application.
* Hidden and filtered rows are skipped when copying.
* `scrollToCell` must use the grid's own row and column offsets, not `offsetTop`, which a CSS grid with sticky cells does not report usefully.

### Media

* TIFF LZW uses early change: the code width grows at 511, 1023 and 2047 entries, not at 512. Getting this wrong truncates the image at about a fifth of its pixels.
* TIFF pixels become a PNG through `pngDataUrl` in `src/media/media.js`, not through a canvas. That keeps the decoder testable outside a browser and lets the EXIF orientation be applied to the buffer. A canvas round trip cannot be checked by a test and costs two large string copies.
* EMF record layouts: `rclBounds` is 16 bytes, the count follows it, then the points. `POLYLINE16` and its relatives start a new contour, `POLYLINE16X` and `POLYLINETO16` continue the current one. Welding a new contour onto the old one fills a letter's counter and turns a logo into a solid block.
* `EMR_FILLPATH` fills with the brush and outlines with the pen, which is where a vector logo's fine detail comes from.
* Stock GDI objects (`0x80000000` upward) supply a pen when the metafile never creates one.
* `EMR_EXTTEXTOUTW` carries its string offset after the scale fields.
* Bitmaps inside a metafile are not drawn. Say so rather than drawing a wrong block.
* JPEG XR (`.wdp`, `.hdp`, `.jxr`) is what Office stores for HD Photo pictures. Chromium cannot decode it, so `src/media/jxr.js` runs the libjxr codec compiled to WebAssembly, lazily and once per session, and `src/media/media.js` encodes the pixels to a PNG (the module is bundled; the notices are in `src/media/jxr/NOTICE.txt`). In the lecture collection every one of the 17 parts is an `a14:imgLayer` effect layer (brightness/contrast, sharpen, background removal) over a base picture. The base PNG is the rendered result PowerPoint saved and is what every viewer draws; the layer is the unadjusted original kept for re-editing, so it is not drawn. A `.wdp` that arrives as a plain picture goes through the decoder and draws. The named placeholder remains for a host that refuses WebAssembly or a package the codec rejects.
* DOCX embedded fonts (`.odttf`) are obfuscated by XORing the first 32 bytes with the GUID in the part name, reversed, twice; the key can also be stored in `w:fontKey`. PPTX `.fntdata` is not obfuscated at all: it is an EOT container whose payload is MicroType Express compressed, so the earlier "no key anywhere" note was wrong. `src/pptx/fonts.js` decodes the part with mtx-decompressor (MPL-2.0) and registers it through the FontFace API before the mount, so autofit measures the face it draws. `p:embeddedFontLst` names the typeface and maps regular, bold, italic and boldItalic to parts; slides only ever name the typeface. Release a cached model's faces with `releaseEmbeddedFonts`, or a session of decks accumulates font sets.

### Legacy binary formats

* A compound document is a file system whose every pointer comes from the file. Each sector number must go through the header offset (sector 0 starts at byte 512); treating a sector number as a byte offset reads the file header back as the stream, which is what a first attempt at Word 97 did.
* A chain ends at any reserved value, 0xfffffffa and up: DIFSECT, FATSECT, ENDOFCHAIN and FREESECT are not sectors. Walking a mini FAT chain as if FATSECT were a sector index leaves the allocation table.
* Streams under 4096 bytes live in the mini stream, which is the root entry's own chain cut into 64 byte mini sectors addressed by the mini FAT. Big streams use the normal FAT.
* Word 97 keeps its text in a piece table, not a string: a list of character ranges, each pointing at 8 bit (cp1252) or 16 bit (UTF-16) text in the WordDocument stream. Bit 30 of a piece's fc marks the 8 bit form, and its offset is then doubled. A file that was never edited has fComplex clear and one contiguous run.
* The text is full of control characters: \u0013 instruction \u0014 result \u0015 field markup, \u0007 cell ends, \u000b soft breaks, \u0001 picture anchors. Keep the results, drop the mechanics.
* Plain runs must be handed to the docx renderer as `{ type: "text" }`. A `{ type: "run" }` means a nested run list, and the renderer recurses into nothing.

### Spreadsheet number formats

* Excel's 1900 calendar contains a day that never existed: serial 60 is 1900-02-29. The epoch is 1899-12-31 and every serial from 61 up loses a day again. The 1904 system has no gap. `test/numfmt.mjs` pins all of this.
* Fraction formats (`# ?/?`, `??/??`, `?/8`) need a real search for the closest denominator; a formatter that falls through to the decimal path prints 1.5 where a timetable expects 1 1/2.
* `m` means month before an `h`/`s` token and minutes after it. A formatter that guesses wrong turns 12:00 into a month.

## 6b. UI patterns that earned their place

* The status line of a view costs a row of the document's height. Put the counts in the app's own status bar (`Plugin.addStatusBarItem`) and keep the in-view element hidden: a host without one still has somewhere to write. The `statusInView` setting (off by default) moves the same text to the bottom left of the view for readers who want it there; the text lives in the element either way, so tests read one source.
* A capture-phase `contextmenu` listener on the host gives documents and decks their own menu (copy the selection, search for it, select all, file details, reload, open in the system editor) and stops the app's plain text menu from winning. The grid already had its own menu through the controller callback.
* The details panel groups rows under headings, keeps its action buttons together at the right edge, offers a copy-as-text button, and hands the file to the system editor with a notice that names a free editor when that fails. This viewer never writes to the file, so any editing wish has to leave the plugin.
* Undecodable media gets a placeholder that names the format ("JPEG XR (HP Photo) image, not decodable here") and a `Not drawn` row in the details, so a reader is never left guessing why a picture is missing.
* A paginated document carries a live page indicator in the toolbar ("Page 14 / 31"), driven by the renderer's scroll listener and reported through `onPageChange`. The footer's own number is not enough: it sits at the bottom of the sheet the reader has already scrolled past.
* Changing the document zoom must scale the scroll offset by the same ratio. CSS `zoom` scales the page layout but leaves `scrollTop` in unscaled pixels, so keeping the offset walks the reader down the document when zooming out and up when zooming in. `renderer.setSettings({ zoom })` adjusts the offset before the browser repaints.

## 7. Performance rules

* The central directory is read and validated first, then parts inflate on demand. Never call `unzipSync` over the whole archive: a 200 MB deck then costs its directory, not a walk over every media entry. This took the largest real deck from 738 ms to 531 ms and a 381 document sweep from 17.4 s to 13.2 s.
* Every part inflates on first use, not at open (`Package.bytes`), and a view prepares the archive through `preparePackage` while the first frame paints: `Package.preload` inflates the XML parts in parallel with `DecompressionStream`, so the parser finds them ready. `preparePackage` keys the prepared package by the bytes object and hands it to `Package.open` once.
* A parse must never resolve media. `parseDrawing` and `parseLegacyPicture` keep the relationship id and leave `url` null; `renderImage`, `drawing.js` and `vml.js` call `model.mediaUrl` when the page is drawn. Resolving at parse transcoded every metafile and JPEG XR picture in the document, which cost the CLB handbook 800 ms per parse in headless Chromium.
* Documents render windowed (`virtualize`). Every page gets a shell at its planned height and only the pages near the viewport carry DOM; pages far away give theirs back. Search reads a model text index built at mount and renders a page when a hit needs it. The reference renderers and the test harness do not set the option, so their output stays complete. Measured on an 11.1 MB, 36 page manual: render 684-934 ms full, 32-60 ms windowed, with 3924 nodes down to 612.
* `tagName` caches the stripped name on the element and `attr` builds a local-name attribute map per element on first miss. A parse reads the same nodes thousands of times and DOM property access dominated the walk; the caches took the same manual's parse from about 250 ms to about 130 ms.
* RGBA pixels become a PNG through `pixelsToDataUrl`: Chromium's canvas encoder when it produces a real image, `pngDataUrl` (fflate plus base64) otherwise. The canvas stub in the test harness fails the length check on purpose, so the headless tests keep exercising the checkable JavaScript encoder.
* `measureWidth` is cached per renderer, and `planPages` caches each block's metrics for the keepNext and keep-with-table lookahead.
* A parsed document is cached (three entries, 32 MB each) keyed by path, size and mtime, and the vault's modify, delete and rename events drop it. The cache owns the model's object URLs and revokes them on eviction.
* Every open takes a token. A read that finishes after the reader moved to another file is dropped instead of mounting the wrong document.
* Pictures are served as blob URLs. Base64 in the DOM made one deck's markup 92 MB.
* The grid builds the frozen rows plus the rows near the viewport, and keeps the full row template so the scroll height and row positions stay right. A 250 row timetable carries about sixty rows of elements.
* Re-render the grid on scroll only when the visible row range changes, and never more than once a frame.
* Search, copy and filters read the model, not the rendered cells.
* Measure before changing anything here, and keep the numbers in the README.

## 8. Verification discipline

* The reference conversion is the standard for text and image placement. Text fidelity means every text run in the source XML appears in the output.
* When a reference disagrees with the file, follow the file and say so. Three known disagreements: the reference converter scales some pictures by their DPI metadata and clips them to the frame, so an image the OOXML stretches to its frame looks cropped; it drops the alpha of some PNGs and paints the transparent area black; and it places one deck's chart bars 170 px from where the OOXML puts them. Another reference renderer and this plugin agree with each other and with the file.
* Pagination is an estimate. A document whose pages are only about seventy percent full with hard page breaks cannot be compared page for page with the reference conversion, which lays the same document out more loosely. Keep the file's breaks and do not tune the page count to the reference.
* Report numbers with their conditions, for example "191 MB deck, 33 slides, opens in 738 ms".
* Measure what the user's own documents contain before building a feature: `test/scan-docs.mjs` over 386 lecture files ranked animations (114), transitions (66), SmartArt (12), .doc (5), embedded fonts (2) and real charts (1). SmartArt and .doc were worth the work; a real chart renderer would have been built for one deck. Reading the package's own `docProps/thumbnail.jpeg` settles what PowerPoint intends when the references disagree.
* Keep the known limits in the README current. Pagination is a measured estimate. Bitmaps inside metafiles are not drawn. Real chart parts show a placeholder. Deeply nested groups with their own scaling and rotation can be a few pixels out.

## 9. Local test configuration

If a file named `PRIVATE_AGENTS.md` exists beside this one, read it first: it
holds this machine's real paths (the vault, the document folders, the reference
tools) and is gitignored. This file and every other tracked file must stay free
of them.

`.testenv` at the repository root, gitignored, never committed. Copy `.testenv.example` and fill in your own paths.

```
OV_TEST_XLSX=      # a workbook for the grid, filter, copy and sizing checks
OV_TEST_DOCX=      # a document for the page renderer
OV_TEST_PPTX=      # a presentation for the slide renderer
OV_TEST_FOLDER=    # a folder to sweep with the stress and fidelity runs
OV_PRIVACY_PATTERNS=  # extra regexes that must never appear in a tracked file
```

A key that is unset skips its checks with a notice. A key that is set but points at nothing stops the run, so a typo is noticed.

## 10. Commits

* Subject: plain and descriptive, no `feat:` or `docs:` prefixes, no emoji.
* Body: terse bullets with a component prefix, one line each, only when the change needs one.
* One logical change per commit. Push only when asked, and then push.
