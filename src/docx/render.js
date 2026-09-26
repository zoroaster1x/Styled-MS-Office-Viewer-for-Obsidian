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
const { renderShapeGroup: drawShapeGroup, applyAnchorStyles } = require("./drawing");
const { renderVml } = require("./vml");

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
    onPageChange: opts.onPageChange || function () {},
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
    let previousSection = -1;
    pages.forEach((blocks, index) => {
      currentPage = index + 1;
      const sectionIndex = blocks.length && blocks[0].section != null ? blocks[0].section : 0;
      const pageSection = (model.sectionList && model.sectionList[sectionIndex]) || section;
      const firstOfSection = sectionIndex !== previousSection;
      previousSection = sectionIndex;
      const page = createPage();
      page.el.dataset.section = String(sectionIndex);
      renderBlocks(blocks, page.content, { state: { page, flow: page.content, usedHeight: 0 }, ctx: { listCounters: new Map(), pageCount: currentPage } });
      finalizePage(page, currentPage, totalPages, pageSection, firstOfSection);
    });
    currentPage = 1;
    renderNotes();
    observePage();
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

  function finalizePage(page, pageNumber, pageTotal, pageSection, firstOfSection) {
    if (settings.showHeaders) renderHeaderFooter(page.header, "header", pageNumber, pageTotal, pageSection, firstOfSection);
    const hasFooter = renderHeaderFooter(page.footer, "footer", pageNumber, pageTotal, pageSection, firstOfSection);
    // Documents often carry no footer at all. A quiet marker in the bottom
    // margin tells the reader which sheet they are on, the way a PDF viewer
    // does, without pretending the document printed it.
    if (!hasFooter && settings.showPageNumbers) {
      const marker = page.footer.createDiv("ov-docx-pagenumber");
      marker.setText(pageNumber + " / " + pageTotal);
    }
  }

  // Returns true when the part had content of its own.
  function renderHeaderFooter(el, kind, pageNumber, pageTotal, pageSection, firstOfSection) {
    if (!settings.showHeaders && kind === "header") {
      el.addClass("is-hidden");
      return false;
    }
    const refs = kind === "header" ? pageSection.headerRefs : pageSection.footerRefs;
    if (!refs) return false;
    const evenOdd = Boolean(model.settings && model.settings.evenAndOddHeaders);
    const even = evenOdd && pageNumber % 2 === 0;
    let rid = null;
    if (firstOfSection && pageSection.titlePage) {
      // A title page shows the first-page part only when the file has one.
      rid = refs.first || null;
    } else {
      rid = (even ? refs.even : null) || refs.default || refs.first || Object.values(refs)[0];
    }
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
    if (hasAnchoredRun(block.runs)) el.addClass("ov-has-anchor");

    if (block.numbering) {
      applyListLayout(el, props, block.numbering);
      const marker = el.createSpan("ov-docx-marker");
      marker.setText(block.numbering.marker);
      const levelRPr = block.numbering.level.rPr || null;
      const markerProps = resolveRunProps(null, levelRPr, props);
      applyRunStyle(marker, markerProps);
    }

    const ctx = {
      listCounters: new Map(),
      tabStops: paragraphTabStops(props),
      cursor: 0,
    };
    // A paragraph that tabs past its own width is a layout line (a cover logo,
    // a signature rule). Word lets it overflow the margin; the browser would
    // wrap it at the last tab, so the line is kept whole.
    if (ctx.tabStops.some((stop) => stop.pos > contentWidthPx)) el.style.whiteSpace = "nowrap";
    renderInline(block.runs, el, props, ctx);
    if (isHeading(props) && block.runs.length) {
      const text = plainText(block.runs);
      if (text) headingAnchors.set("heading-" + headingAnchors.size, el);
    }
    registerText(el);
  }

  // Word lays tabs out on explicit stops, or on the default half-inch grid. A
  // right tab puts the end of the following content at the stop, which is how
  // a cover logo lands at the right of the page. The positions are measured
  // from the text margin, so a left indent shifts them in our box.
  function paragraphTabStops(props) {
    const indentPx = props.indentLeftTw != null ? twipToPx(props.indentLeftTw) : 0;
    const stops = [];
    for (const tab of props.tabs || []) {
      stops.push({ pos: twipToPx(tab.posTw) - indentPx, align: tab.align || "left", leader: tab.leader || "none" });
    }
    stops.sort((a, b) => a.pos - b.pos);
    return stops;
  }

  function tabStopAfter(ctx, cursor) {
    for (const stop of ctx.tabStops) {
      if (stop.pos > cursor + 0.5) return stop;
    }
    return { pos: (Math.floor(cursor / 48) + 1) * 48, align: "left", leader: "none" };
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

  function renderInline(runs, parent, paraProps, ctx, fieldState) {
    // Field characters arrive in separate w:r elements, so their state has to
    // outlive one pass of this loop.
    const state = fieldState || { stack: [] };
    if (!ctx) ctx = { listCounters: new Map(), tabStops: [], cursor: 0 };
    if (!ctx.tabStops) ctx.tabStops = [];
    if (typeof ctx.cursor !== "number") ctx.cursor = 0;
    for (let i = 0; i < runs.length; i++) {
      const run = runs[i];
      if (!run) continue;
      if (run.type === "text") {
        if (suppressFieldResult(state, parent)) continue;
        ctx.cursor += appendTextRun(parent, run, paraProps);
      } else if (run.type === "run") {
        renderInline(run.runs, parent, paraProps, ctx, state);
      } else if (run.type === "link") {
        renderLink(parent, run.link, paraProps, ctx, state);
      } else if (run.type === "break") {
        appendBreak(parent);
        ctx.cursor = 0;
      } else if (run.type === "tab") {
        i = renderTabGroup(parent, paraProps, ctx, runs, i, state);
      } else if (run.type === "textbox") {
        renderTextBox(parent, run, paraProps);
        ctx.cursor += flowWidth(run);
      } else if (run.type === "image") {
        renderImage(parent, run);
        ctx.cursor += flowWidth(run);
      } else if (run.type === "shapegroup") {
        renderShapeGroup(parent, run);
        ctx.cursor += flowWidth(run);
      } else if (run.type === "vml") {
        renderVmlShape(parent, run);
        ctx.cursor += flowWidth(run);
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
        handleFieldChar(run, state, parent);
      } else if (run.type === "instr") {
        const frame = state.stack[state.stack.length - 1];
        if (frame) frame.instr += run.text;
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

  // PAGE and NUMPAGES substitute the number at draw time. The saved result in
  // the file can be stale (one footer carries the value 53 in a 31-page
  // document), so it is dropped in favour of the live figure.
  function handleFieldChar(run, state, parent) {
    if (run.stage === "begin") {
      state.stack.push({ instr: "", kind: null, mode: "instr", emitted: false });
      return;
    }
    if (run.stage === "separate") {
      const frame = state.stack[state.stack.length - 1];
      if (frame) {
        frame.kind = fieldKindFromInstr(frame.instr);
        frame.mode = "result";
      }
      return;
    }
    if (run.stage === "end") {
      const frame = state.stack.pop();
      if (!frame) return;
      if (!frame.kind) frame.kind = fieldKindFromInstr(frame.instr);
      if (frame.kind && !frame.emitted) emitFieldNumber(frame, parent);
    }
  }

  function fieldKindFromInstr(instr) {
    const text = String(instr || "").toUpperCase();
    if (!text) return null;
    if (/NUMPAGES|SECTIONPAGES/.test(text)) return "pages";
    if (/\bPAGE\b/.test(text)) return "page";
    return null;
  }

  function suppressFieldResult(state, parent) {
    const frame = state.stack[state.stack.length - 1];
    if (!frame || frame.mode !== "result" || !frame.kind) return false;
    if (!frame.emitted) emitFieldNumber(frame, parent);
    return true;
  }

  function emitFieldNumber(frame, parent) {
    frame.emitted = true;
    const marker = parent.createSpan("ov-docx-field");
    marker.setText(String(frame.kind === "page" ? currentPage : totalPages));
    marker.dataset.field = frame.kind;
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

  function renderLink(parent, link, paraProps, ctx, fieldState) {
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
    renderInline(link.runs, anchor, paraProps, ctx, fieldState);
    parent.appendChild(anchor);
  }

  function appendTextRun(parent, run, paraProps) {
    const props = resolveRunProps(null, run.props, paraProps);
    if (!run.text) return 0;
    const span = parent.createSpan("ov-docx-r");
    if (run.deleted) span.addClass("is-deleted");
    span.setText(run.text);
    applyRunStyle(span, props);
    return measureWidth(run.text, runFontString(props));
  }

  function runFontString(props) {
    const sizePt = props && props.sizeHalfPt != null ? props.sizeHalfPt / 2 : DEFAULT_FONT_PT;
    let family = props && props.fontFamily;
    if (!family && props && props.fontFamilyTheme && model.theme && model.theme.fonts) {
      family = model.theme.fonts.minor;
    }
    return ptToPxLocal(sizePt) + "px " + (family || "Calibri") + ", Calibri, sans-serif";
  }

  function renderTabGroup(parent, paraProps, ctx, runs, index, fieldState) {
    const stop = tabStopAfter(ctx, ctx.cursor);
    let end = index + 1;
    while (end < runs.length && runs[end] && runs[end].type !== "break" && !containsTab(runs[end])) end++;
    const following = runs.slice(index + 1, end);

    if (stop.align === "left" || following.length === 0) {
      // A left stop is a plain advance; the content after it flows on.
      const width = Math.max(0, stop.pos - ctx.cursor);
      const span = parent.createSpan("ov-docx-tabstop");
      span.style.width = width + "px";
      applyLeader(span, stop);
      ctx.cursor = stop.pos;
      return index;
    }

    // A right or centre stop aligns the content that follows it. An atomic box
    // of the stop's width, with the content aligned inside, keeps the layout
    // exact and the line unbreakable; a span plus a separate content element
    // cannot know the content width before it is laid out.
    const width = Math.max(0, stop.pos - ctx.cursor);
    const box = parent.createSpan("ov-docx-tabbox");
    box.style.width = width + "px";
    box.style.textAlign = stop.align;
    renderInline(following, box, paraProps, { listCounters: ctx.listCounters, tabStops: [], cursor: 0 }, fieldState);
    ctx.cursor = stop.pos;
    return end - 1;
  }

  // A tab inside a run wrapper is still a tab for the grouping walk.
  function containsTab(run) {
    if (!run) return false;
    if (run.type === "tab") return true;
    if (run.type === "run") return (run.runs || []).some(containsTab);
    return false;
  }

  function applyLeader(span, stop) {
    if (!stop.leader || stop.leader === "none") return;
    span.addClass("is-leader");
    span.setAttribute("data-leader", stop.leader);
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
    if (run.anchor && run.anchor.outOfFlow) {
      applyAnchorStyles(wrap, run.anchor, contentWidthPx, width || 0);
    } else if (run.anchor) {
      wrap.addClass("ov-docx-image-float");
    }
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

  // A Word drawing group from wpg:wgp. The drawing module places every child
  // from its own transform; paragraph rendering stays here so a text box uses
  // the same style cascade as the rest of the document.
  function renderShapeGroup(parent, run) {
    drawShapeGroup(parent, run, {
      theme: model.theme,
      contentWidthPx,
      mediaUrl: (rid) => model.mediaUrl(rid),
      parseParagraph: (el) => require("./parse").parseParagraph(el, model, { listCounters: new Map() }),
      drawParagraph: (block, host) => renderParagraph(block, host),
    });
  }

  // A legacy VML drawing (w:pict). Coordinates and style are read by the vml
  // module; paragraphs keep the document style cascade.
  function renderVmlShape(parent, run) {
    renderVml(parent, run, {
      theme: model.theme,
      contentWidthPx,
      pageMarginLeftPx: twipToPx(marginLeftTw),
      pageMarginTopPx: twipToPx(marginTopTw),
      mediaUrl: (rid) => model.mediaUrl(rid),
      parseParagraph: (el) => require("./parse").parseParagraph(el, model, { listCounters: new Map() }),
      drawParagraph: (block, host) => renderParagraph(block, host),
    });
  }

  // A text box from a shape: its own runs, drawn where the shape is anchored.
  function renderTextBox(parent, run, paraProps) {
    const box = parent.createDiv("ov-docx-textbox");
    if (run.widthPx) box.style.maxWidth = Math.min(run.widthPx, contentWidthPx) + "px";
    if (run.anchor && run.anchor.outOfFlow) {
      box.style.width = Math.min(run.widthPx || contentWidthPx, contentWidthPx) + "px";
      applyAnchorStyles(box, run.anchor, contentWidthPx, run.widthPx || 0);
    }
    const props = paraProps || null;
    renderInline(run.runs, box, props, { listCounters: new Map() });
    if (run.alt) box.title = run.alt;
  }

  // True when a paragraph carries an out-of-flow drawing. The paragraph then
  // becomes the positioning context for the anchors inside it.
  function hasAnchoredRun(runs) {
    for (const run of runs || []) {
      if (!run) continue;
      if (run.anchor && run.anchor.outOfFlow) return true;
      if (run.type === "run" && hasAnchoredRun(run.runs)) return true;
      if (run.type === "link" && hasAnchoredRun(run.link.runs)) return true;
    }
    return false;
  }

  function flowWidth(run) {
    if (run.anchor && run.anchor.outOfFlow) return 0;
    return Number(run.widthPx) > 0 ? Number(run.widthPx) : 0;
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
    // A table style carries the base borders, cell margins and the conditional
    // formatting (banded rows, a bold first row) the file relies on.
    const tableStyle = block.styleId ? model.styles.table.get(block.styleId) : null;
    const styleProps = tableStyle && tableStyle.tblPr ? tableStyle.tblPr : null;
    const baseBorders = props.borders || (styleProps && styleProps.borders) || null;
    const baseMargin = props.cellMargin || (styleProps && styleProps.cellMargin) || null;
    const look = props.look || (styleProps && styleProps.look) || {};
    const conditionals = styleProps && tableStyle.conditionals ? tableStyle.conditionals : [];

    const table = parent.createEl("table", { cls: "ov-docx-table" });
    if (props.align === "center") table.style.marginLeft = "auto", table.style.marginRight = "auto";
    else if (props.align === "right") table.style.marginLeft = "auto";
    if (props.widthTw) table.style.width = twipToPx(props.widthTw) + "px";
    else table.style.width = "100%";
    // tblInd shifts the whole table from the margin.
    if (props.indentTw && props.align !== "center" && props.align !== "right") {
      table.style.marginLeft = twipToPx(props.indentTw) + "px";
    }

    const grid = block.grid && block.grid.length ? block.grid : null;
    const colCount = computeColumnCount(block, grid);
    if (grid) {
      const widths = grid.map((w) => (w > 0 ? twipToPx(w) : null));
      const known = widths.reduce((a, b) => a + (b || 0), 0);
      const missing = widths.filter((w) => !w).length;
      const fallback = missing ? Math.max(24, (contentWidthPx - known) / missing) : 0;
      // The grid is the truth for a fixed layout. Chrome lets the column
      // widths grow the table past its declared width when they disagree, so
      // the two are kept equal here.
      if (!props.widthTw && known) table.style.width = (known + missing * fallback) + "px";
      table.style.tableLayout = "fixed";
      const colgroup = doc.createElement("colgroup");
      for (let i = 0; i < colCount; i++) {
        const col = doc.createElement("col");
        col.style.width = (widths[i] || fallback || Math.max(24, contentWidthPx / colCount)) + "px";
        colgroup.appendChild(col);
      }
      table.appendChild(colgroup);
    }

    const rowCount = block.rows.length;
    const tbody = doc.createElement("tbody");
    table.appendChild(tbody);
    for (let r = 0; r < rowCount; r++) {
      const row = block.rows[r];
      const tr = doc.createElement("tr");
      if (row.props && row.props.heightTw) {
        tr.style.height = twipToPx(row.props.heightTw) + "px";
      }
      if (row.props && row.props.header) tr.addClass("ov-docx-table-header");
      const rowConditionals = conditionalsForRow(conditionals, r, rowCount, look);
      if (rowConditionals.some((entry) => entry.type === "firstRow" || entry.rPr && entry.rPr.bold)) {
        tr.addClass("ov-docx-table-strong");
      }
      let column = 0;
      for (const cell of row.cells) {
        const span = cell.gridSpan || 1;
        // A vMerge continuation is covered by the restart cell's rowspan above.
        // Emitting it as well would shift every following cell one column right
        // and split the row.
        if (cell.vMerge === "continue") {
          column += span;
          continue;
        }
        renderCell(tr, cell, block, r, Boolean(grid), {
          baseBorders,
          baseMargin,
          conditionals: conditionalsForCell(rowConditionals, conditionals, column, colCount, look),
        });
        column += span;
      }
      tbody.appendChild(tr);
    }
    parent.appendChild(table);
    registerText(table);
  }

  // Which tblStylePr entries apply to a row: the first and last row, and the
  // horizontal bands. Bands alternate over the data rows, which is where the
  // first and last rows are excluded when the look says so.
  function conditionalsForRow(conditionals, rowIndex, rowCount, look) {
    const out = [];
    const has = (type) => conditionals.find((entry) => entry.type === type);
    const first = look.firstRow !== false && rowIndex === 0;
    const last = look.lastRow !== false && rowIndex === rowCount - 1;
    if (first) {
      const entry = has("firstRow");
      if (entry) out.push(entry);
    }
    if (last) {
      const entry = has("lastRow");
      if (entry) out.push(entry);
    }
    if (look.noHBand !== true) {
      const bandIndex = first ? -1 : rowIndex - (look.firstRow === false ? 0 : 1);
      if (!last && bandIndex >= 0) {
        const entry = has(bandIndex % 2 === 0 ? "band1Horz" : "band2Horz");
        if (entry) out.push(entry);
      }
    }
    return out;
  }

  function conditionalsForCell(rowConditionals, conditionals, column, colCount, look) {
    const out = rowConditionals.slice();
    const has = (type) => conditionals.find((entry) => entry.type === type);
    if (look.firstColumn !== false && column === 0) {
      const entry = has("firstCol");
      if (entry) out.push(entry);
    }
    if (look.lastColumn !== false && column === colCount - 1) {
      const entry = has("lastCol");
      if (entry) out.push(entry);
    }
    if (look.noVBand !== true) {
      const entry = has(column % 2 === 0 ? "band1Vert" : "band2Vert");
      if (entry) out.push(entry);
    }
    return out;
  }

  function computeColumnCount(block, grid) {
    if (grid) return grid.length;
    let max = 0;
    for (const row of block.rows) {
      let count = 0;
      for (const cell of row.cells) count += cell.gridSpan || 1;
      max = Math.max(max, count);
    }
    return max || 1;
  }

  function renderCell(tr, cell, block, rowIndex, hasGrid, table) {
    const props = cell.props || {};
    const context = table || {};
    const conditional = pickConditional(context.conditionals);
    const td = doc.createElement("td");
    // Attributes, not properties: a host that does not reflect colSpan onto the
    // attribute drops the span out of any serialised HTML.
    if (cell.gridSpan > 1) td.setAttribute("colspan", String(cell.gridSpan));
    const rowspan = continuationRowspan(block, rowIndex, cell);
    if (rowspan > 1) td.setAttribute("rowspan", String(rowspan));
    // The colgroup owns the widths when there is a grid; a cell width on top of
    // it can make Chrome widen the table past the declared size.
    if (!hasGrid && props.widthTw) td.style.width = twipToPx(props.widthTw) + "px";
    const vAlign = props.vAlign || (conditional && conditional.tcPr && conditional.tcPr.vAlign);
    if (vAlign) td.style.verticalAlign = vAlign === "center" ? "middle" : vAlign;
    const shading = props.shading || (conditional && conditional.tcPr && conditional.tcPr.shading);
    if (shading && shading.color) td.style.backgroundColor = shading.color;
    // Direct cell borders win over the conditional style, which wins over the
    // table's own borders and the table style's base.
    const conditionalBorders = conditional && conditional.tcPr ? conditional.tcPr.borders : null;
    const merged = mergeBorders(context.baseBorders || null, props.borders || null);
    applyCellBorders(td, mergeBorders(merged, conditionalBorders));
    const tableMargin = context.baseMargin || null;
    const cellMargin = props.margin || (conditional && conditional.tcPr && conditional.tcPr.margin) || null;
    if (tableMargin || cellMargin) {
      // The table's tblCellMar is the default for every cell; tcMar overrides
      // it per side. Without this a cell that states zero margins still gets
      // the stylesheet's 3px 5px and a 31-row form overflows its page.
      for (const side of ["top", "left", "bottom", "right"]) {
        const value = cellMargin && cellMargin[side] != null ? cellMargin[side]
          : tableMargin && tableMargin[side] != null ? tableMargin[side]
            : 0;
        td.style["padding" + side[0].toUpperCase() + side.slice(1)] = twipToPx(value) + "px";
      }
    }
    const cellState = { state: null, ctx: { listCounters: new Map() } };
    renderBlocks(cell.blocks, td, cellState);
    tr.appendChild(td);
  }

  // Word applies conditionals in a fixed order; the last entry that states a
  // property is the one that holds.
  function pickConditional(conditionals) {
    if (!conditionals || !conditionals.length) return null;
    return conditionals[conditionals.length - 1];
  }

  function continuationRowspan(block, rowIndex, cell) {
    if (cell.vMerge !== "restart") return 1;
    // The covering cell lives in the same grid column, not the same cell index:
    // a gridSpan earlier in the row shifts every later cell.
    const column = gridColumnOf(block.rows[rowIndex], cell);
    let span = 1;
    for (let r = rowIndex + 1; r < block.rows.length; r++) {
      const match = findCellAt(block.rows[r], column);
      if (match && match.vMerge === "continue") span++;
      else break;
    }
    return span;
  }

  function gridColumnOf(row, cell) {
    let column = 0;
    for (const candidate of row.cells) {
      if (candidate === cell) return column;
      column += candidate.gridSpan || 1;
    }
    return column;
  }

  function findCellAt(row, column) {
    let count = 0;
    for (const cell of row.cells) {
      const span = cell.gridSpan || 1;
      if (column >= count && column < count + span) return cell;
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
        else if ((run.type === "shapegroup" || run.type === "vml") && run.texts) out += run.texts.join(" ") + " ";
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
        if (inner.type === "image" || inner.type === "textbox" || inner.type === "shapegroup" || inner.type === "vml") {
          const top = (inner.offsetPx && inner.offsetPx.top) || 0;
          const imageHeight = Number(inner.heightPx) > 0 ? Number(inner.heightPx) + top : 0;
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
    const metricsCache = new Map();
    const nextPage = () => {
      pages.push(current);
      current = [];
      used = 0;
    };
    const blockMetrics = (block) => {
      let metrics = metricsCache.get(block);
      if (!metrics) {
        metrics = estimateBlock(block);
        metricsCache.set(block, metrics);
      }
      return metrics;
    };
    const blockTotal = (block) => {
      const metrics = blockMetrics(block);
      return metrics.before + metrics.height + metrics.after;
    };
    const headingBlock = (block) => {
      if (!block || block.type !== "p") return false;
      return isHeading(resolveParagraphProps(block.style, block.props));
    };
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      const metrics = blockMetrics(block);
      const total = metrics.before + metrics.height + metrics.after;
      const forced = block.props && block.props.pageBreakBefore;
      const nextTotal = blocks[i + 1] ? blockTotal(blocks[i + 1]) : 0;
      if (forced && (current.length || used > 0)) nextPage();
      // keepNext: this block must share its page with the one after it.
      if (block.props && block.props.keepNext && current.length
        && used + total + nextTotal > contentHeightPx && used + total <= contentHeightPx) {
        nextPage();
      }
      if (used > 0 && used + total > contentHeightPx) {
        // A heading directly before a table travels with the table, which is
        // what Word and the reference renderers do with a form's title.
        const last = current.length ? current[current.length - 1] : null;
        if (block.type === "table" && headingBlock(last)) {
          current.pop();
          used -= blockTotal(last);
          nextPage();
          current.push(last);
          used += blockTotal(last);
        } else {
          nextPage();
        }
      }
      current.push(block);
      used += total;
      if (block.pageBreakAfter) nextPage();
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

  // ---------- page tracking ----------

  let lastReportedPage = 0;

  function observePage() {
    if (!scrollEl || !scrollEl.addEventListener) return;
    scrollEl.addEventListener("scroll", onScroll, { passive: true });
    reportCurrentPage();
  }

  let scrollPending = false;
  function onScroll() {
    if (scrollPending) return;
    scrollPending = true;
    const raf = typeof requestAnimationFrame === "function" ? requestAnimationFrame : (cb) => setTimeout(cb, 16);
    raf(() => {
      scrollPending = false;
      reportCurrentPage();
    });
  }

  function pagesInOrder() {
    const out = [];
    for (const el of pagesEl.children) {
      if (el.classList && el.classList.contains("ov-docx-page")) out.push(el);
    }
    return out;
  }

  // The page the reader is on: the first page whose bottom edge is still below
  // the top of the scroll viewport.
  function currentPageNumber() {
    const list = pagesInOrder();
    if (!list.length) return 1;
    const top = scrollEl.getBoundingClientRect().top + 8;
    for (let i = 0; i < list.length; i++) {
      if (list[i].getBoundingClientRect().bottom > top) return i + 1;
    }
    return list.length;
  }

  function reportCurrentPage() {
    const count = pagesInOrder().length;
    if (!count) return;
    const page = currentPageNumber();
    if (page === lastReportedPage && lastReportedPage !== 0) return;
    lastReportedPage = page;
    callbacks.onPageChange({ page, count });
  }

  function setSettings(patch) {
    const previousZoom = settings.zoom || 1;
    Object.assign(settings, patch);
    if (!rootEl) return;
    rootEl.style.setProperty("--ov-docx-zoom", String(settings.zoom));
    // CSS zoom scales the pages, and the scroll offset stays in unscaled
    // pixels. Keeping the raw offset walks the reader down the document when
    // zooming out; following the same ratio keeps the page in view.
    if (scrollEl && patch && patch.zoom != null && patch.zoom !== previousZoom && previousZoom > 0) {
      scrollEl.scrollTop = scrollEl.scrollTop * (patch.zoom / previousZoom);
      reportCurrentPage();
    }
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
    currentPage: currentPageNumber,
    pageCount: () => pagesInOrder().length,
  };
}

module.exports = { createDocxRenderer };
