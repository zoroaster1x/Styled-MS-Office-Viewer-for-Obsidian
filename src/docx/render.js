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

// Renders the parsed document into DOM: a paginated white page, the style
// cascade applied, tables laid out from their grid, images placed inline or
// floating. Nothing here writes to the document; the model is read only.

const { twipToPx, halfPtToPx, emuToPx } = require("../shared/units");
const { tagName } = require("../shared/xml");
const { mathText } = require("./math");
const { renderMathDom } = require("../shared/math-dom");

const DEFAULT_FONT_PT = 11;
const DEFAULT_PAGE_WIDTH_TW = 11906;
const DEFAULT_PAGE_HEIGHT_TW = 16838;
const DEFAULT_MARGIN_TW = 1440;

function createDocxRenderer(opts) {
  const container = opts.container;
  const model = opts.model;
  const doc = container.ownerDocument || document;
  const settings = Object.assign({
    zoom: 1,
    pageBackground: "white",
    showHeaders: true,
    showPageNumbers: true,
    paginate: true,
    maxPageHeightPx: 0,
  }, opts.settings || {});
  const callbacks = {
    onInternalLink: opts.onInternalLink || function () {},
    onExternalLink: opts.onExternalLink || function (href) { window.open(href, "_blank"); },
    onReady: opts.onReady || function () {},
  };

  let destroyed = false;
  let rootEl = null;
  let scrollEl = null;
  let pagesEl = null;
  const headingAnchors = new Map();
  const noteRefs = [];
  const textItems = [];

  // Text measurement for pagination. The same font stack the runs use, so the
  // line count matches what the browser will lay out.
  let measureCtx = null;
  const measureCache = new Map();
  function measureWidth(text, font) {
    const cacheKey = font + "\u0000" + text;
    if (measureCache.has(cacheKey)) return measureCache.get(cacheKey);
    try {
      if (!measureCtx) {
        const canvas = doc.createElement("canvas");
        measureCtx = canvas.getContext ? canvas.getContext("2d") : null;
      }
      if (!measureCtx) return text.length * 6;
      measureCtx.font = font;
      const width = measureCtx.measureText(text).width;
      measureCache.set(cacheKey, width);
      return width;
    } catch (err) {
      return text.length * 6;
    }
  }

  const section = model.section || {};
  const pageWidthTw = section.pageWidthTw || DEFAULT_PAGE_WIDTH_TW;
  const pageHeightTw = section.pageHeightTw || DEFAULT_PAGE_HEIGHT_TW;
  const marginLeftTw = section.marginLeftTw == null ? DEFAULT_MARGIN_TW : section.marginLeftTw;
  const marginRightTw = section.marginRightTw == null ? DEFAULT_MARGIN_TW : section.marginRightTw;
  const marginTopTw = section.marginTopTw == null ? DEFAULT_MARGIN_TW : section.marginTopTw;
  const marginBottomTw = section.marginBottomTw == null ? DEFAULT_MARGIN_TW : section.marginBottomTw;

  const pageWidthPx = twipToPx(pageWidthTw);
  const pageHeightPx = twipToPx(pageHeightTw);
  const contentWidthPx = Math.max(120, pageWidthPx - twipToPx(marginLeftTw) - twipToPx(marginRightTw));
  const contentHeightPx = Math.max(120, pageHeightPx - twipToPx(marginTopTw) - twipToPx(marginBottomTw));
  // Word can snap every line to the document grid, which makes lines taller
  // than the font alone and is what turns a blank form into several pages.
  const gridApplies = section.docGridType === "lines" || section.docGridType === "linesAndChars";
  const gridLinePx = section.docGrid && gridApplies ? twipToPx(section.docGrid) : 0;

  function snappedLine(linePx, props) {
    if (!gridLinePx || (props && props.snapToGrid === false)) return linePx;
    const steps = Math.max(1, Math.ceil((linePx - 0.01) / gridLinePx));
    return steps * gridLinePx;
  }

  // ---------- style cascade ----------

  const paraCache = new Map();
  const runCache = new Map();

  function resolveParagraphProps(styleId, direct) {
    const key = (styleId || "") + "|" + hashProps(direct);
    if (paraCache.has(key)) return paraCache.get(key);
    const chain = styleChain(styleId, "paragraph");
    const out = {};
    applyDefaults(out, model.styles.docDefaults ? model.styles.docDefaults.pPr : null);
    for (const entry of chain) {
      applyDefaults(out, entry.pPr);
      // A style keeps its text formatting in a top level rPr, which is the
      // default for every run in the paragraph. Without this a heading loses
      // its bold and size as soon as a run sets a font of its own.
      if (entry.rPr) out.rPr = Object.assign({}, out.rPr || {}, entry.rPr);
    }
    applyDefaults(out, direct);
    paraCache.set(key, out);
    return out;
  }

  function resolveRunProps(styleId, direct, paraResolved) {
    // The paragraph context is part of the identity: two runs with the same
    // direct formatting resolve differently in a body paragraph and a heading.
    const key = (styleId || "") + "|" + hashProps(direct) + "|"
      + hashProps(paraResolved && paraResolved.rPr ? paraResolved.rPr : null);
    if (runCache.has(key)) return runCache.get(key);
    const chain = styleChain(styleId, "character");
    const out = {};
    applyDefaults(out, model.styles.docDefaults ? model.styles.docDefaults.rPr : null);
    for (const entry of chain) applyDefaults(out, entry.rPr);
    if (paraResolved && paraResolved.rPr) applyDefaults(out, paraResolved.rPr);
    applyDefaults(out, direct);
    runCache.set(key, out);
    return out;
  }

  function styleChain(styleId, type) {
    const out = [];
    const table = type === "character" ? model.styles.character : type === "table" ? model.styles.table : model.styles.paragraph;
    let id = styleId;
    let guard = 0;
    const seen = new Set();
    while (id && guard++ < 30 && !seen.has(id)) {
      seen.add(id);
      const entry = table.get(id);
      if (!entry) break;
      out.unshift(entry);
      id = entry.basedOn;
    }
    return out;
  }

  function isBag(value) {
    return value && typeof value === "object" && !Array.isArray(value);
  }

  function applyDefaults(target, source) {
    if (!source) return target;
    for (const key of Object.keys(source)) {
      const value = source[key];
      if (value === undefined || value === null) continue;
      if (key === "type" || key === "id" || key === "name") continue;
      if (isBag(value) && isBag(target[key])) {
        // Merge nested property bags (run properties, shading, borders) so
        // direct formatting overrides only the keys it actually sets.
        target[key] = Object.assign({}, target[key], value);
        continue;
      }
      target[key] = value;
    }
    return target;
  }

  function hashProps(props) {
    if (!props) return "";
    let out = "";
    for (const key of Object.keys(props).sort()) {
      const v = props[key];
      if (v && typeof v === "object") out += key + ":" + JSON.stringify(v) + ";";
      else out += key + ":" + v + ";";
    }
    return out;
  }

  // ---------- build ----------

  function render() {
    if (destroyed) return;
    rootEl = container.createDiv("ov-docx");
    rootEl.style.setProperty("--ov-docx-zoom", String(settings.zoom));
    if (settings.pageBackground === "theme") rootEl.addClass("ov-theme-surface");
    scrollEl = rootEl.createDiv("ov-docx-scroll");
    pagesEl = scrollEl.createDiv("ov-docx-pages");
    pagesEl.style.width = pageWidthPx + "px";

    const pages = planPages(model.body);
    totalPages = pages.length;
    pages.forEach((blocks, index) => {
      currentPage = index + 1;
      const page = createPage();
      renderBlocks(blocks, page.content, { state: { page, flow: page.content, usedHeight: 0 }, ctx: { listCounters: new Map(), pageCount: currentPage } });
      finalizePage(page, currentPage, totalPages);
    });
    currentPage = 1;
    renderNotes();
    callbacks.onReady({ outline: buildOutline(), pageCount: totalPages });
  }

  function createPage() {
    const page = doc.createElement("div");
    page.className = "ov-docx-page";
    page.style.width = pageWidthPx + "px";
    page.style.minHeight = pageHeightPx + "px";
    page.style.paddingTop = twipToPx(marginTopTw) + "px";
    page.style.paddingBottom = twipToPx(marginBottomTw) + "px";
    page.style.paddingLeft = twipToPx(marginLeftTw) + "px";
    page.style.paddingRight = twipToPx(marginRightTw) + "px";
    const header = page.createDiv("ov-docx-pageheader");
    const content = page.createDiv("ov-docx-pagecontent");
    content.style.minHeight = Math.max(60, pageHeightPx - twipToPx(marginTopTw) - twipToPx(marginBottomTw)) + "px";
    const footer = page.createDiv("ov-docx-pagefooter");
    pagesEl.appendChild(page);
    return { el: page, header, content, footer };
  }

  function finalizePage(page, pageNumber, pageTotal) {
    if (settings.showHeaders) renderHeaderFooter(page.header, "header", pageNumber, pageTotal);
    const hasFooter = renderHeaderFooter(page.footer, "footer", pageNumber, pageTotal);
    // Documents often carry no footer at all. A quiet marker in the bottom
    // margin tells the reader which sheet they are on, the way a PDF viewer
    // does, without pretending the document printed it.
    if (!hasFooter && settings.showPageNumbers) {
      const marker = page.footer.createDiv("ov-docx-pagenumber");
      marker.setText(pageNumber + " / " + pageTotal);
    }
  }

  // Returns true when the part had content of its own.
  function renderHeaderFooter(el, kind, pageNumber, pageTotal) {
    if (!settings.showHeaders && kind === "header") {
      el.addClass("is-hidden");
      return false;
    }
    const refs = kind === "header" ? section.headerRefs : section.footerRefs;
    if (!refs) return false;
    const rid = refs.default || refs.first || Object.values(refs)[0];
    if (!rid) return false;
    const part = (kind === "header" ? model.headers : model.footers).get(rid);
    if (!part) return false;
    if (!part.blocks) {
      const docObj = model.pkg.xml(part.path);
      part.blocks = docObj ? parsePartBlocks(docObj.documentElement, model) : [];
    }
    const previousPage = currentPage;
    const previousTotal = totalPages;
    currentPage = pageNumber;
    totalPages = pageTotal;
    const ctx = { listCounters: new Map(), pageCount: pageNumber };
    renderBlocks(part.blocks, el, { state: { page: null, flow: el, usedHeight: 0 }, ctx, inHeader: true });
    currentPage = previousPage;
    totalPages = previousTotal;
    return true;
  }

  function parsePartBlocks(root, model2) {
    const body = root;
    const out = [];
    for (const child of body.children || []) {
      const tag = tagName(child);
      if (tag === "p") out.push(require("./parse").parseParagraph(child, model2, { listCounters: new Map() }));
      else if (tag === "tbl") out.push(require("./parse").parseTable(child, model2, { listCounters: new Map() }));
    }
    return out;
  }

  // ---------- blocks ----------

  function renderBlocks(blocks, parent, state) {
    for (const block of blocks) {
      if (destroyed) return;
      if (block.type === "p") renderParagraph(block, parent);
      else if (block.type === "table") renderTable(block, parent, state);
      if (block.bookmarks && block.bookmarks.length) {
        for (const name of block.bookmarks) headingAnchors.set(name, parent.lastElementChild || parent);
      }
    }
  }

  function renderParagraph(block, parent) {
    const props = resolveParagraphProps(block.style, block.props);
    const el = parent.createDiv("ov-docx-p");
    applyParagraphStyle(el, props, block);
    if (block.sectionBreak) el.addClass("ov-docx-section-break");

    if (block.numbering) {
      applyListLayout(el, props, block.numbering);
      const marker = el.createSpan("ov-docx-marker");
      marker.setText(block.numbering.marker);
      const levelRPr = block.numbering.level.rPr || null;
      const markerProps = resolveRunProps(null, levelRPr, props);
      applyRunStyle(marker, markerProps);
    }

    const ctx = { listCounters: new Map() };
    renderInline(block.runs, el, props, ctx);
    if (isHeading(props) && block.runs.length) {
      const text = plainText(block.runs);
      if (text) headingAnchors.set("heading-" + headingAnchors.size, el);
    }
    registerText(el);
  }

  function applyParagraphStyle(el, props, block) {
    const styles = {};
    // The paragraph mark's run properties size the line itself, which is what
    // an empty paragraph's line box and the paragraph's base metrics use.
    const markProps = resolveRunProps(null, block.markRunProps || null, props);
    if (markProps && markProps.sizeHalfPt != null) styles.fontSize = ptToPxLocal(markProps.sizeHalfPt / 2) + "px";
    if (props.align) styles.textAlign = props.align;
    if (props.spaceBeforePt != null) styles.marginTop = ptToPxLocal(props.spaceBeforePt) + "px";
    if (props.spaceAfterPt != null) styles.marginBottom = ptToPxLocal(props.spaceAfterPt) + "px";
    if (props.lineHeightPt != null) styles.lineHeight = snappedLine(ptToPxLocal(props.lineHeightPt), props) + "px";
    else if (props.lineHeight && props.lineRule === "auto" && props.lineHeight > 1.01) styles.lineHeight = String(props.lineHeight);
    else if (gridLinePx) {
      // Snap the natural line height up to the document grid.
      const sizePt = props.rPr && props.rPr.sizeHalfPt ? props.rPr.sizeHalfPt / 2 : DEFAULT_FONT_PT;
      styles.lineHeight = snappedLine(ptToPxLocal(sizePt) * 1.18, props) + "px";
    }
    const indentLeft = props.indentLeftTw != null ? twipToPx(props.indentLeftTw) : null;
    const indentRight = props.indentRightTw != null ? twipToPx(props.indentRightTw) : null;
    const hanging = props.indentHangingTw != null ? twipToPx(props.indentHangingTw) : 0;
    const firstLine = props.indentFirstLineTw != null ? twipToPx(props.indentFirstLineTw) : 0;
    const pad = (indentLeft || 0) + hanging;
    if (pad) styles.paddingLeft = pad + "px";
    if (indentRight) styles.paddingRight = indentRight + "px";
    const textIndent = -hanging + firstLine;
    if (textIndent) styles.textIndent = textIndent + "px";
    if (props.keepNext) el.addClass("ov-keep-next");
    if (props.pageBreakBefore) el.addClass("ov-page-break-before");
    if (props.contextualSpacing) styles.marginTop = "0px";
    el.setCssStyles(styles);
    if (props.shading && props.shading.color) {
      el.style.backgroundColor = props.shading.color;
    }
    if (props.borders) {
      applyBoxBorders(el, props.borders, "padding");
    }
  }

  function applyListLayout(el, props, numbering) {
    const lvl = numbering.level || {};
    const lvlPr = lvl.pPr || {};
    const left = lvlPr.indentLeftTw != null ? lvlPr.indentLeftTw : props.indentLeftTw;
    const hang = lvlPr.indentHangingTw != null ? lvlPr.indentHangingTw : props.indentHangingTw;
    const leftPx = left != null ? twipToPx(left) : 0;
    const hangPx = hang != null ? twipToPx(hang) : 0;
    el.style.paddingLeft = (leftPx + hangPx) + "px";
    el.style.textIndent = (-hangPx) + "px";
  }

  function isHeading(props) {
    if (!props) return false;
    if (props.outlineLevel != null && props.outlineLevel >= 0 && props.outlineLevel <= 8) return true;
    return false;
  }

  // ---------- inline ----------

  function renderInline(runs, parent, paraProps, ctx) {
    let fieldState = null;
    let fieldRuns = [];
    let fieldKind = null;
    for (const run of runs) {
      if (!run) continue;
      if (run.type === "text") {
        appendTextRun(parent, run, paraProps, fieldState ? fieldRuns : null);
      } else if (run.type === "run") {
        renderInline(run.runs, parent, paraProps, ctx);
      } else if (run.type === "link") {
        renderLink(parent, run.link, paraProps, ctx);
      } else if (run.type === "break") {
        appendBreak(parent);
      } else if (run.type === "tab") {
        const tab = parent.createSpan("ov-docx-tab");
        tab.setText("\t");
        const tabProps = resolveRunProps(null, run.props, paraProps);
        applyRunStyle(tab, tabProps);
      } else if (run.type === "textbox") {
        renderTextBox(parent, run, paraProps);
      } else if (run.type === "image") {
        renderImage(parent, run);
      } else if (run.type === "math") {
        renderMath(run.node, parent, paraProps);
      } else if (run.type === "placeholder") {
        renderPlaceholder(parent, run);
      } else if (run.type === "noteRef") {
        const ref = parent.createEl("sup", { cls: "ov-docx-noteref" });
        ref.setText(String(noteCounter(run.kind)));
        ref.dataset.noteId = run.noteId;
        ref.dataset.noteKind = run.kind;
        ref.addEventListener("click", () => scrollToNote(run.kind, run.noteId));
        noteRefs.push(ref);
      } else if (run.type === "commentRef") {
        const ref = parent.createSpan("ov-docx-commentref");
        ref.setText("[c]");
        ref.title = "Comment " + (run.id || "");
      } else if (run.type === "fieldChar") {
        if (run.stage === "begin") {
          fieldState = "instr";
          fieldRuns = [];
          fieldKind = null;
        } else if (run.stage === "separate") {
          fieldState = "result";
        } else if (run.stage === "end") {
          if (fieldState === "instr" && fieldRuns.length === 0) {
            const instr = plainText(fieldRuns);
            fieldKind = /PAGE/.test(instr) ? "page" : /NUMPAGES/.test(instr) ? "pages" : null;
          }
          if (fieldKind === "page" || fieldKind === "pages") {
            const marker = parent.createSpan("ov-docx-field");
            marker.setText(String(fieldKind === "page" ? currentPage : totalPages));
            marker.dataset.field = fieldKind;
          }
          fieldState = null;
          fieldRuns = [];
          fieldKind = null;
        }
      } else if (run.type === "instr") {
        fieldRuns.push({ type: "text", text: run.text });
      } else if (run.type === "field") {
        const marker = parent.createSpan("ov-docx-field");
        if (run.kind === "page") marker.setText(String(currentPage));
        else if (run.kind === "pages") marker.setText(String(totalPages));
        else marker.setText(plainText(run.runs));
        marker.dataset.field = run.kind;
        if (run.runs && run.runs.length) {
          const props = resolveRunProps(null, run.runs[0].props, paraProps);
          applyRunStyle(marker, props);
        }
      }
    }
  }

  let currentPage = 1;
  let totalPages = 1;

  let noteCount = 0;
  let noteCounts = { footnote: 0, endnote: 0 };
  function noteCounter(kind) {
    noteCount++;
    noteCounts[kind] = (noteCounts[kind] || 0) + 1;
    return noteCounts[kind];
  }

  function renderLink(parent, link, paraProps, ctx) {
    if (!link) return;
    const anchor = doc.createElement("span");
    anchor.className = "ov-docx-link";
    const href = link.href;
    if (href) {
      anchor.addClass("is-external");
      anchor.addEventListener("click", (ev) => {
        ev.preventDefault();
        callbacks.onExternalLink(href);
      });
    } else if (link.anchor) {
      anchor.addClass("is-internal");
      anchor.addEventListener("click", (ev) => {
        ev.preventDefault();
        callbacks.onInternalLink(link.anchor);
      });
    }
    renderInline(link.runs, anchor, paraProps, ctx);
    parent.appendChild(anchor);
  }

  function appendTextRun(parent, run, paraProps, fieldRuns) {
    const props = resolveRunProps(null, run.props, paraProps);
    if (!run.text) {
      if (fieldRuns) fieldRuns.push(run);
      return;
    }
    if (fieldRuns) fieldRuns.push(run);
    const span = parent.createSpan("ov-docx-r");
    if (run.deleted) span.addClass("is-deleted");
    span.setText(run.text);
    applyRunStyle(span, props);
  }

  function appendBreak(parent) {
    parent.createEl("br");
  }

  function applyRunStyle(el, props) {
    const styles = {};
    if (props.fontFamilyTheme && model.theme && model.theme.fonts) {
      const minor = model.theme.fonts.minor;
      if (minor) styles.fontFamily = minor + ", Calibri, sans-serif";
    }
    if (props.fontFamily) styles.fontFamily = props.fontFamily;
    const sizePt = props.sizeHalfPt != null ? props.sizeHalfPt / 2 : DEFAULT_FONT_PT;
    styles.fontSize = ptToPxLocal(sizePt) + "px";
    if (props.bold) styles.fontWeight = "700";
    if (props.italic) styles.fontStyle = "italic";
    const decorations = [];
    if (props.underline) decorations.push("underline");
    if (props.strike) decorations.push("line-through");
    if (decorations.length) styles.textDecoration = decorations.join(" ");
    if (props.color) styles.color = props.color;
    if (props.highlight) styles.backgroundColor = props.highlight;
    if (props.shading && props.shading.color) styles.backgroundColor = props.shading.color;
    if (props.letterSpacingPt) styles.letterSpacing = ptToPxLocal(props.letterSpacingPt) + "px";
    if (props.positionHalfPt) styles.verticalAlign = (props.positionHalfPt / 2) + "px";
    if (props.vertAlign === "superscript") styles.verticalAlign = "super";
    else if (props.vertAlign === "subscript") styles.verticalAlign = "sub";
    if (props.vertAlign === "superscript" || props.vertAlign === "subscript") styles.fontSize = "0.72em";
    if (props.caps) styles.textTransform = "uppercase";
    if (props.smallCaps) styles.fontVariant = "small-caps";
    if (props.hidden) styles.display = "none";
    if (props.emphasis === "dot") styles.textEmphasis = "dot";
    el.setCssStyles(styles);
    if (props.underline && props.underline !== "single" && props.underline !== "none") {
      el.style.textDecorationStyle = props.underline === "dotted" ? "dotted" : props.underline === "dash" ? "dashed" : "solid";
    }
  }

  function renderImage(parent, run) {
    const wrap = doc.createElement("span");
    wrap.className = "ov-docx-image";
    const width = run.widthPx && run.widthPx > 0 ? run.widthPx : null;
    const height = run.heightPx && run.heightPx > 0 ? run.heightPx : null;
    if (width) wrap.style.width = width + "px";
    if (height) wrap.style.height = height + "px";
    if (run.anchor) wrap.addClass("ov-docx-image-float");
    if (run.url) {
      const img = wrap.createEl("img");
      img.src = run.url;
      img.alt = run.alt || run.name || "";
      img.loading = "lazy";
      if (width) img.style.width = "100%";
      if (height) img.style.height = "100%";
      img.setAttribute("draggable", "false");
    } else {
      wrap.addClass("is-missing");
      wrap.setText("Image could not be drawn");
      wrap.title = run.name || "";
    }
    parent.appendChild(wrap);
  }

  // A text box from a shape: its own runs, drawn where the shape is anchored.
  function renderTextBox(parent, run, paraProps) {
    const box = parent.createDiv("ov-docx-textbox");
    if (run.widthPx) box.style.maxWidth = Math.min(run.widthPx, contentWidthPx) + "px";
    const props = paraProps || null;
    renderInline(run.runs, box, props, { listCounters: new Map() });
    if (run.alt) box.title = run.alt;
  }

  function renderPlaceholder(parent, run) {
    const box = parent.createDiv("ov-docx-placeholder");
    if (run.widthPx) box.style.width = Math.min(run.widthPx, contentWidthPx) + "px";
    if (run.heightPx) box.style.height = run.heightPx + "px";
    box.setText(run.label || "Diagram");
    box.title = "This drawing is stored in a form the viewer does not render";
  }

  // ---------- tables ----------

  function renderTable(block, parent, state) {
    const props = block.props || {};
    const table = parent.createEl("table", { cls: "ov-docx-table" });
    if (props.align === "center") table.style.marginLeft = "auto", table.style.marginRight = "auto";
    else if (props.align === "right") table.style.marginLeft = "auto";
    if (props.widthTw) table.style.width = twipToPx(props.widthTw) + "px";
    else table.style.width = "100%";

    const grid = block.grid && block.grid.length ? block.grid : null;
    const colCount = computeColumnCount(block, grid);
    if (grid && grid.length) {
      const widths = grid.map((w) => (w > 0 ? twipToPx(w) : null));
      const total = widths.reduce((a, b) => a + (b || 0), 0);
      table.style.tableLayout = "fixed";
      const colgroup = doc.createElement("colgroup");
      for (let i = 0; i < colCount; i++) {
        const col = doc.createElement("col");
        if (widths[i]) col.style.width = widths[i] + "px";
        else if (total > 0) col.style.width = Math.max(24, contentWidthPx / colCount) + "px";
        colgroup.appendChild(col);
      }
      table.appendChild(colgroup);
    }

    const tbody = doc.createElement("tbody");
    table.appendChild(tbody);
    for (let r = 0; r < block.rows.length; r++) {
      const row = block.rows[r];
      const tr = doc.createElement("tr");
      if (row.props && row.props.heightTw) {
        tr.style.height = twipToPx(row.props.heightTw) + "px";
      }
      if (row.props && row.props.header) tr.addClass("ov-docx-table-header");
      for (const cell of row.cells) renderCell(tr, cell, block, r);
      tbody.appendChild(tr);
    }
    parent.appendChild(table);
    registerText(table);
  }

  function computeColumnCount(block, grid) {
    if (grid && grid.length) return grid.length + 1;
    let max = 0;
    for (const row of block.rows) {
      let count = 0;
      for (const cell of row.cells) count += cell.gridSpan || 1;
      max = Math.max(max, count);
    }
    return max || 1;
  }

  function renderCell(tr, cell, block, rowIndex) {
    const props = cell.props || {};
    const td = doc.createElement("td");
    if (cell.gridSpan > 1) td.colSpan = cell.gridSpan;
    const rowspan = continuationRowspan(block, rowIndex, cell);
    if (rowspan > 1) td.rowSpan = rowspan;
    if (cell.vMerge === "continue") {
      // Covered by the cell above; still emitted so the column count stays right.
    }
    if (props.widthTw) td.style.width = twipToPx(props.widthTw) + "px";
    if (props.vAlign) td.style.verticalAlign = props.vAlign === "center" ? "middle" : props.vAlign;
    if (props.shading && props.shading.color) td.style.backgroundColor = props.shading.color;
    const tableBorders = block.props && block.props.borders ? block.props.borders : null;
    const borders = mergeBorders(tableBorders, props.borders);
    applyCellBorders(td, borders);
    if (props.margin) {
      if (props.margin.left != null) td.style.paddingLeft = twipToPx(props.margin.left) + "px";
      else td.style.paddingLeft = "5.4pt";
      if (props.margin.top != null) td.style.paddingTop = twipToPx(props.margin.top) + "px";
      if (props.margin.bottom != null) td.style.paddingBottom = twipToPx(props.margin.bottom) + "px";
      if (props.margin.right != null) td.style.paddingRight = twipToPx(props.margin.right) + "px";
    }
    const cellState = { state: null, ctx: { listCounters: new Map() } };
    renderBlocks(cell.blocks, td, cellState);
    tr.appendChild(td);
  }

  function continuationRowspan(block, rowIndex, cell) {
    if (cell.vMerge !== "restart") return 1;
    let span = 1;
    for (let r = rowIndex + 1; r < block.rows.length; r++) {
      const nextRow = block.rows[r];
      const match = findCellAt(nextRow, cellIndexInRow(block.rows[rowIndex], cell));
      if (match && match.vMerge === "continue") span++;
      else break;
    }
    return span;
  }

  function cellIndexInRow(row, cell) {
    return row.cells.indexOf(cell);
  }

  function findCellAt(row, index) {
    if (index < 0) return null;
    let count = 0;
    for (const cell of row.cells) {
      const span = cell.gridSpan || 1;
      if (index >= count && index < count + span) return cell;
      count += span;
    }
    return null;
  }

  function mergeBorders(tableBorders, cellBorders) {
    const out = {};
    const sides = { top: "top", left: "left", bottom: "bottom", right: "right" };
    for (const key of Object.keys(sides)) {
      const cellEdge = cellBorders && cellBorders[key];
      const tableEdge = tableBorders && (tableBorders[key] || tableBorders[key === "left" ? "insideV" : key === "top" ? "insideH" : key]);
      out[key] = cellEdge || tableEdge || null;
    }
    return out;
  }

  function applyCellBorders(td, borders) {
    if (!borders) return;
    for (const side of ["top", "left", "bottom", "right"]) {
      const edge = borders[side];
      if (!edge) continue;
      td.style["border" + side[0].toUpperCase() + side.slice(1)] =
        edge.width + "px " + edge.style + " " + edge.color;
    }
  }

  function applyBoxBorders(el, borders) {
    for (const side of ["top", "left", "bottom", "right"]) {
      const edge = borders[side];
      if (!edge) continue;
      el.style["border" + side[0].toUpperCase() + side.slice(1)] =
        edge.width + "px " + edge.style + " " + edge.color;
    }
  }

  // ---------- footnotes and endnotes ----------

  function renderNotes() {
    const footnotes = [];
    for (const [id] of model.footnotes) footnotes.push(id);
    if (!footnotes.length && (!model.endnotes || !model.endnotes.size)) return;
    const section = pagesEl.createDiv("ov-docx-notes");
    if (model.footnotes && model.footnotes.size) {
      section.createDiv("ov-docx-notes-title", (el) => el.setText("Footnotes"));
      for (const [id, note] of model.footnotes) renderNote(note, section);
    }
    if (model.endnotes && model.endnotes.size) {
      section.createDiv("ov-docx-notes-title", (el) => el.setText("Endnotes"));
      for (const [id, note] of model.endnotes) renderNote(note, section);
    }
  }

  function renderNote(note, host) {
    const el = host.createDiv("ov-docx-note");
    el.dataset.noteId = note.id;
    el.dataset.noteKind = note.type;
    const blocks = note.blocks && note.blocks.length ? note.blocks : [];
    for (const block of blocks) {
      if (block.type === "p") {
        const p = el.createDiv("ov-docx-note-p");
        const props = resolveParagraphProps(block.style, block.props);
        renderInline(block.runs, p, props, { listCounters: new Map() });
      }
    }
    registerText(el);
  }

  function scrollToNote(kind, id) {
    const target = pagesEl.querySelector('.ov-docx-note[data-note-id="' + cssEscape(id) + '"]');
    if (target) target.scrollIntoView({ block: "center", behavior: "smooth" });
  }

  function scrollToAnchor(name) {
    const el = headingAnchors.get(name);
    if (el) el.scrollIntoView({ block: "start", behavior: "smooth" });
    return Boolean(el);
  }

  // ---------- text index for search ----------

  function registerText(el) {
    textItems.push({ el, text: (el.textContent || "").replace(/\s+/g, " ").trim() });
  }

  // ---------- equations ----------
  //
  // The structure becomes real layout: a fraction is a small stacked box with a
  // rule, a script is a script, a radical carries its sign and a bar. Text runs
  // inside an equation keep Word's italic convention unless m:nor says the run
  // is an operator or a unit.

  // The equation tree is shared with the slide renderer; see shared/math-dom.js.
  function renderMath(node, parent, paraProps) {
    renderMathDom(node, parent, paraProps);
  }

  function plainText(runs) {
    let out = "";
    const walk = (list) => {
      for (const run of list || []) {
        if (!run) continue;
        if (run.type === "text") out += run.text;
        else if (run.type === "math") out += mathText(run.node);
        else if (run.type === "run") walk(run.runs);
        else if (run.type === "link") walk(run.link.runs);
        else if (run.type === "tab") out += "\t";
      }
    };
    walk(runs);
    return out;
  }

  // ---------- outline ----------

  function buildOutline() {
    const outline = [];
    const walk = (blocks, depth) => {
      for (const block of blocks || []) {
        if (block.type === "p") {
          const props = resolveParagraphProps(block.style, block.props);
          const text = plainText(block.runs).trim();
          if (text && isHeading(props)) {
            outline.push({ text, level: Math.min(4, (props.outlineLevel || 0) + 1), anchor: block.bookmarks && block.bookmarks[0] });
          }
        } else if (block.type === "table") {
          for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks, depth + 1);
        }
      }
    };
    walk(model.body, 0);
    return outline;
  }

  // ---------- search ----------

  let searchState = { hits: [], index: -1 };

  function search(query) {
    clearSearch();
    const q = String(query || "").toLowerCase();
    if (!q) return 0;
    let count = 0;
    for (const item of textItems) {
      if ((item.text || "").toLowerCase().indexOf(q) !== -1) {
        item.el.addClass("ov-search-hit");
        count++;
      }
    }
    return count;
  }

  function clearSearch() {
    for (const item of textItems) {
      item.el.removeClass("ov-search-hit");
      item.el.removeClass("ov-search-current");
    }
    searchState = { hits: [], index: -1 };
  }

  function searchNext(dir) {
    const hits = [];
    for (const item of textItems) {
      if (item.el.hasClass("ov-search-hit")) hits.push(item.el);
    }
    if (!hits.length) return null;
    searchState.index = (searchState.index + dir + hits.length * 2) % hits.length;
    for (const el of hits) el.removeClass("ov-search-current");
    const el = hits[searchState.index];
    el.addClass("ov-search-current");
    el.scrollIntoView({ block: "center", behavior: "smooth" });
    return { index: searchState.index + 1, count: hits.length };
  }

  function ptToPxLocal(pt) {
    return Math.round(pt * 96) / 72;
  }

  // ---------- page planning ----------

  // The height one paragraph needs: its lines at the resolved font size plus
  // the spacing around it. An empty paragraph still occupies one line, which is
  // why a form full of blank lines has to keep its shape.
  function estimateParagraph(block) {
    const props = resolveParagraphProps(block.style, block.props);
    let sizePt = null;
    const runs = block.runs || [];
    for (const run of runs) {
      const candidate = run.props || (run.runs && run.runs[0] ? run.runs[0].props : null);
      if (candidate && candidate.sizeHalfPt) {
        sizePt = candidate.sizeHalfPt / 2;
        break;
      }
    }
    if (sizePt == null) {
      const markProps = resolveRunProps(null, block.markRunProps || null, props);
      if (markProps && markProps.sizeHalfPt != null) sizePt = markProps.sizeHalfPt / 2;
    }
    const fontPt = sizePt == null ? DEFAULT_FONT_PT : sizePt;
    const lineFactor = props.lineHeight && props.lineHeight > 1 ? props.lineHeight : 1.18;
    const baseLine = props.lineHeightPt ? ptToPxLocal(props.lineHeightPt) : ptToPxLocal(fontPt) * lineFactor;
    const linePx = snappedLine(baseLine, props);
    const text = plainText(runs).replace(/\u00ad/g, "");
    const leftIndent = (props.indentLeftTw ? twipToPx(props.indentLeftTw) : 0) + (props.indentHangingTw ? twipToPx(props.indentHangingTw) : 0);
    const rightIndent = props.indentRightTw ? twipToPx(props.indentRightTw) : 0;
    const available = Math.max(40, contentWidthPx - leftIndent - rightIndent);
    const font = (props.rPr && props.rPr.fontFamily ? props.rPr.fontFamily : "Calibri") + ", sans-serif";
    let lines = 1;
    if (text) {
      lines = 0;
      for (const piece of text.split("\n")) {
        const width = measureWidth(piece, ptToPxLocal(fontPt) + "px " + font);
        lines += Math.max(1, Math.ceil(width / available));
      }
    }
    // A paragraph can hold a picture or a text box, and those occupy their own
    // height. Without this a page of images would be counted as a page of text.
    let contentHeight = lines * linePx;
    for (const run of runs) {
      const blocks = run.type === "run" ? run.runs : [run];
      for (const inner of blocks || []) {
        if (!inner) continue;
        if (inner.type === "image" || inner.type === "textbox") {
          const imageHeight = Number(inner.heightPx) > 0 ? Number(inner.heightPx) : 0;
          const imageLines = Number(inner.heightPx) > 0 ? Math.ceil(imageHeight / Math.max(1, linePx)) : 0;
          contentHeight = Math.max(contentHeight, imageHeight + 4, imageLines * linePx);
        }
      }
    }
    const before = props.spaceBeforePt != null && props.spaceBeforePt > 0 ? ptToPxLocal(props.spaceBeforePt) : 0;
    const after = props.spaceAfterPt != null && props.spaceAfterPt > 0 ? ptToPxLocal(props.spaceAfterPt) : 0;
    return { height: contentHeight, before, after, linePx, lines };
  }

  function estimateTable(block) {
    // Word gives every row at least one line of the default font, however
    // empty the form is, which is what makes a blank form several pages long.
    const defaultLine = estimateParagraph({ type: "p", runs: [], props: null, style: null }).linePx;
    let total = 0;
    for (const row of block.rows) {
      let rowHeight = row.props && row.props.heightTw ? twipToPx(row.props.heightTw) : 0;
      for (const cell of row.cells) {
        let cellHeight = 0;
        for (const inner of cell.blocks) {
          if (inner.type === "p") {
            const m = estimateParagraph(inner);
            cellHeight += m.before + m.height + m.after;
          } else if (inner.type === "table") {
            cellHeight += estimateTable(inner);
          }
        }
        rowHeight = Math.max(rowHeight, cellHeight + 6, defaultLine + 6);
      }
      total += rowHeight;
    }
    return total;
  }

  function estimateBlock(block) {
    if (block.type === "table") {
      const before = block.props && block.props.pageBreakBefore ? 0 : 0;
      return { height: estimateTable(block), before, after: 0 };
    }
    return estimateParagraph(block);
  }

  // Splits the body into pages of content that each fit the printable height.
  function planPages(blocks) {
    const pages = [];
    let current = [];
    let used = 0;
    const nextPage = () => {
      pages.push(current);
      current = [];
      used = 0;
    };
    for (const block of blocks) {
      const metrics = estimateBlock(block);
      const total = metrics.before + metrics.height + metrics.after;
      const forced = block.props && block.props.pageBreakBefore;
      if (forced && (current.length || used > 0)) nextPage();
      if (used > 0 && used + total > contentHeightPx) nextPage();
      current.push(block);
      used += total;
    }
    if (current.length || pages.length === 0) nextPage();
    return pages;
  }

  function getScrollPosition() {
    return scrollEl ? { top: scrollEl.scrollTop, left: 0 } : { top: 0, left: 0 };
  }

  function setScrollPosition(pos) {
    if (scrollEl && pos) scrollEl.scrollTop = pos.top || 0;
  }

  function setSettings(patch) {
    Object.assign(settings, patch);
    if (rootEl) rootEl.style.setProperty("--ov-docx-zoom", String(settings.zoom));
  }

  function destroy() {
    destroyed = true;
    if (container) container.textContent = "";
    rootEl = scrollEl = pagesEl = null;
    textItems.length = 0;
    headingAnchors.clear();
  }

  function cssEscape(value) {
    return String(value).replace(/["\\]/g, "\\$&");
  }

  render();

  return {
    destroy,
    setSettings,
    getScrollPosition,
    setScrollPosition,
    search,
    searchNext,
    clearSearch,
    scrollToAnchor,
    scrollToNote,
    getOutline: buildOutline,
  };
}

module.exports = { createDocxRenderer };
