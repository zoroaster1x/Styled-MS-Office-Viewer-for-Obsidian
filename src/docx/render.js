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

// Text measurement is shared by every renderer, so the measuring canvas and the
// cache live here rather than per mount.
let SHARED_MEASURE_CTX = null;
const MEASURE_CACHE = new Map();
const MEASURE_CACHE_LIMIT = 40000;

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
    // Windowed rendering is off unless the view asks for it, so a reference
    // render (and the test harness) always draws every page.
    virtualize: false,
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
  const noteRefs = [];
  // Windowed rendering: every page gets a shell at its planned height, and only
  // the pages near the viewport carry their content. A long document opens at
  // the cost of its first pages instead of all of them, and the pages far away
  // give their DOM back.
  let pageRecords = [];
  const blockElements = new Map();
  const anchorBlocks = new Map();
  const headingNames = new Map();
  let textIndex = [];
  let searchHits = [];
  let searchBlockSet = new Set();
  let searchIndex = -1;
  let searchCurrentEl = null;

  // Text measurement for pagination. The same font stack the runs use, so the
  // line count matches what the browser will lay out. The canvas and its cache
  // are shared by every renderer: the same strings measure the same in every
  // document, so re-planning a cached file or opening the same text again is
  // free.
  function measureWidth(text, font) {
    const cacheKey = font + "\u0000" + text;
    const hit = MEASURE_CACHE.get(cacheKey);
    if (hit !== undefined) return hit;
    try {
      if (!SHARED_MEASURE_CTX && typeof document !== "undefined" && document.createElement) {
        const canvas = document.createElement("canvas");
        SHARED_MEASURE_CTX = canvas.getContext ? canvas.getContext("2d") : null;
      }
      if (!SHARED_MEASURE_CTX) return text.length * 6;
      SHARED_MEASURE_CTX.font = font;
      const width = SHARED_MEASURE_CTX.measureText(text).width;
      if (MEASURE_CACHE.size >= MEASURE_CACHE_LIMIT) MEASURE_CACHE.clear();
      MEASURE_CACHE.set(cacheKey, width);
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
    // The run defaults from the style chain, before the paragraph mark's own
    // rPr is merged. A run inherits the style, not the formatting of the
    // paragraph mark: Word writes a paragraph mark in Times while its runs
    // stay in the document's Calibri.
    out.styleRPr = Object.assign({}, out.rPr || {});
    applyDefaults(out, direct);
    paraCache.set(key, out);
    return out;
  }

  function resolveRunProps(styleId, direct, paraResolved) {
    // The paragraph context is part of the identity: two runs with the same
    // direct formatting resolve differently in a body paragraph and a heading.
    const key = (styleId || "") + "|" + hashProps(direct) + "|"
      + hashProps(paraResolved && paraResolved.styleRPr ? paraResolved.styleRPr : null);
    if (runCache.has(key)) return runCache.get(key);
    const chain = styleChain(styleId, "character");
    const out = {};
    applyDefaults(out, model.styles.docDefaults ? model.styles.docDefaults.rPr : null);
    for (const entry of chain) applyDefaults(out, entry.rPr);
    if (paraResolved && paraResolved.styleRPr) applyDefaults(out, paraResolved.styleRPr);
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

    const plan = planPages(model.body);
    totalPages = plan.length;
    let previousSection = -1;
    pageRecords = plan.map((entry, index) => {
      const blocks = entry.blocks;
      const sectionIndex = blocks.length && blocks[0].section != null ? blocks[0].section : 0;
      const pageSection = (model.sectionList && model.sectionList[sectionIndex]) || section;
      const firstOfSection = sectionIndex !== previousSection;
      previousSection = sectionIndex;
      const page = createPage(entry.height);
      page.el.dataset.section = String(sectionIndex);
      return {
        number: index + 1,
        blocks,
        pageSection,
        firstOfSection,
        height: entry.height,
        el: page.el,
        header: page.header,
        content: page.content,
        footer: page.footer,
        minHeight: page.minHeight,
        rendered: false,
      };
    });
    buildBlockIndex();
    if (settings.virtualize === true) {
      renderPageWindow(1);
    } else {
      for (const record of pageRecords) renderPage(record);
    }
    currentPage = 1;
    renderNotes();
    observePage();
    callbacks.onReady({ outline: buildOutline(), pageCount: totalPages });
  }

  // The DOM for one page, drawn on demand in windowed mode.
  function renderPage(record) {
    if (!record || record.rendered || destroyed) return;
    const previousPage = currentPage;
    const previousTotal = totalPages;
    currentPage = record.number;
    totalPages = pageRecords.length;
    currentPageContent = record.content;
    renderBlocks(record.blocks, record.content, {
      state: { page: record, flow: record.content, usedHeight: 0 },
      ctx: { listCounters: new Map(), pageCount: record.number },
    });
    finalizePage(
      { el: record.el, header: record.header, content: record.content, footer: record.footer },
      record.number,
      totalPages,
      record.pageSection,
      record.firstOfSection
    );
    currentPage = previousPage;
    totalPages = previousTotal;
    currentPageContent = null;
    record.rendered = true;
  }

  // A page far from the viewport keeps its height but gives its DOM back.
  function unrenderPage(record) {
    if (!record || !record.rendered) return;
    record.header.textContent = "";
    record.content.textContent = "";
    record.footer.textContent = "";
    record.el.style.minHeight = record.minHeight + "px";
    record.rendered = false;
    forEachBlock(record.blocks, (block) => blockElements.delete(block));
  }

  const RENDER_BEHIND = 1;
  const RENDER_AHEAD = 3;
  const KEEP_BEHIND = 4;
  const KEEP_AHEAD = 8;

  function renderPageWindow(center) {
    if (settings.virtualize !== true || destroyed) return;
    const page = Math.max(1, Math.min(totalPages, center || 1));
    for (const record of pageRecords) {
      if (record.number >= page - RENDER_BEHIND && record.number <= page + RENDER_AHEAD) {
        renderPage(record);
      } else if (record.number < page - KEEP_BEHIND || record.number > page + KEEP_AHEAD) {
        unrenderPage(record);
      }
    }
  }

  function forEachBlock(blocks, fn) {
    for (const block of blocks || []) {
      fn(block);
      if (block.type === "table") {
        for (const row of block.rows) for (const cell of row.cells) forEachBlock(cell.blocks, fn);
      }
    }
  }

  function mapBlock(block, el) {
    blockElements.set(block, el);
    if (searchBlockSet.has(block)) el.addClass("ov-search-hit");
  }

  // The search index is the model's text, not the DOM, so it covers pages that
  // are not rendered yet. A block maps to its element as soon as it is drawn.
  function buildBlockIndex() {
    textIndex = [];
    for (const record of pageRecords) {
      forEachBlock(record.blocks, (block) => {
        if (block.type !== "p") return;
        const text = plainText(block.runs).replace(/\s+/g, " ").trim();
        if (!text) return;
        textIndex.push({ block, page: record.number, text, lower: text.toLowerCase() });
        if (block.bookmarks) {
          for (const name of block.bookmarks) anchorBlocks.set(name, { block, page: record.number });
        }
        const props = resolveParagraphProps(block.style, block.props);
        if (isHeading(props)) {
          const name = "heading-" + headingNames.size;
          headingNames.set(block, name);
          anchorBlocks.set(name, { block, page: record.number });
        }
      });
    }
  }

  function findRecordForBlock(block) {
    for (const record of pageRecords) {
      let found = false;
      forEachBlock(record.blocks, (candidate) => {
        if (candidate === block) found = true;
      });
      if (found) return record;
    }
    return null;
  }

  function createPage(placeholderHeight) {
    const page = doc.createElement("div");
    page.className = "ov-docx-page";
    page.style.width = pageWidthPx + "px";
    // A page that is not drawn yet still holds its planned height, so the
    // scrollbar and the page number stay honest.
    const planned = (placeholderHeight || 0) + twipToPx(marginTopTw) + twipToPx(marginBottomTw) + 4;
    const minHeight = Math.max(pageHeightPx, planned);
    page.style.minHeight = minHeight + "px";
    page.style.paddingTop = twipToPx(marginTopTw) + "px";
    page.style.paddingBottom = twipToPx(marginBottomTw) + "px";
    page.style.paddingLeft = twipToPx(marginLeftTw) + "px";
    page.style.paddingRight = twipToPx(marginRightTw) + "px";
    const header = page.createDiv("ov-docx-pageheader");
    const content = page.createDiv("ov-docx-pagecontent");
    content.style.minHeight = Math.max(60, pageHeightPx - twipToPx(marginTopTw) - twipToPx(marginBottomTw)) + "px";
    const footer = page.createDiv("ov-docx-pagefooter");
    // The header and footer live in the page margins, not in the body flow:
    // Word measures the header from the page edge and the footer from the
    // bottom edge. Letting them take part in the flex column grew the sheet
    // past its paper size and pushed the first line of every page down.
    const headerDistanceTw = section.headerTw != null ? section.headerTw : 708;
    const footerDistanceTw = section.footerTw != null ? section.footerTw : 708;
    header.style.top = twipToPx(headerDistanceTw) + "px";
    footer.style.bottom = twipToPx(footerDistanceTw) + "px";
    header.style.left = twipToPx(marginLeftTw) + "px";
    header.style.right = twipToPx(marginRightTw) + "px";
    footer.style.left = twipToPx(marginLeftTw) + "px";
    footer.style.right = twipToPx(marginRightTw) + "px";
    pagesEl.appendChild(page);
    return { el: page, header, content, footer, minHeight };
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
    // A multi-column section is not drawn as columns yet: CSS columns force
    // equal widths, while the files that use them here (an equation sheet)
    // state a narrow and a wide column, and a float inside an equal column
    // would be pushed below its text. The blocks flow in one column, which
    // keeps every page-relative shape where the file puts it.
    for (const block of blocks) {
      if (destroyed) return;
      if (block.type === "p") renderParagraph(block, parent);
      else if (block.type === "table") renderTable(block, parent, state);
    }
  }

  function renderParagraph(block, parent) {
    const props = resolveParagraphProps(block.style, block.props);
    const el = parent.createDiv("ov-docx-p");
    applyParagraphStyle(el, props, block);
    if (block.sectionBreak) el.addClass("ov-docx-section-break");
    if (block.columnBreakBefore) el.addClass("ov-docx-column-break");
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
      plannedTop: block.plannedTop || 0,
      blockLeft: parseFloat(el.style.paddingLeft) || 0,
      reserve: null,
      takeReserve: null,
    };
    // A top-and-bottom (or through) VML box pushes the lines of its paragraph
    // below its bottom edge. The paragraph reserves the tallest box once: one
    // block spacer at the first box's place, so a shape that shares its
    // paragraph with a line of text keeps the text under the drawing.
    let reserved = maxVmlReserve(block.runs);
    ctx.reserve = (px) => { if (px > reserved) reserved = px; };
    ctx.takeReserve = () => {
      const value = reserved;
      reserved = 0;
      return value;
    };
    // The line the paragraph's own text will follow the spacer with, so the
    // spacer can stop short by one line and the paragraph still reserves the
    // box's bottom, not the bottom plus the text.
    ctx.lineHeightPx = (function () {
      const declared = el.style.lineHeight;
      const size = parseFloat(el.style.fontSize) || 16;
      if (declared) {
        const value = parseFloat(declared);
        if (declared.indexOf("px") !== -1) return value;
        if (value > 0 && value < 5) return value * size;
        return value;
      }
      return size * 1.18;
    })();
    // A paragraph that tabs past its own width is a layout line (a cover logo,
    // a signature rule). Word lets it overflow the margin; the browser would
    // wrap it at the last tab, so the line is kept whole.
    if (ctx.tabStops.some((stop) => stop.pos > contentWidthPx)) el.style.whiteSpace = "nowrap";
    renderInline(block.runs, el, props, ctx);
    if (reserved > 0) el.style.minHeight = reserved + "px";
    // A paragraph whose only content is an overlay still occupies its own line
    // in Word, which is what spaces a column of anchored boxes down the page.
    let hasFlow = false;
    for (const child of el.children) {
      if (!child.classList || !child.classList.contains("ov-docx-anchor")) {
        hasFlow = true;
        break;
      }
    }
    if (!hasFlow) el.addClass("ov-anchor-only");
    mapBlock(block, el);
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
    if (!ctx) ctx = { listCounters: new Map(), tabStops: [], cursor: 0, plannedTop: 0, blockLeft: 0 };
    if (!ctx.tabStops) ctx.tabStops = [];
    if (typeof ctx.plannedTop !== "number") ctx.plannedTop = 0;
    if (typeof ctx.blockLeft !== "number") ctx.blockLeft = 0;
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
        // A column break moves to a column of a section this viewer draws as a
        // single column; breaking the line there would push text down a line
        // the file never breaks.
        if (!run.column) {
          appendBreak(parent);
          ctx.cursor = 0;
        }
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
        renderVmlShape(parent, run, ctx);
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
  // The page content element of the page being drawn, so a page-relative VML
  // overlay can be positioned against the page rather than its paragraph.
  let currentPageContent = null;
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
    applyRunStyle(span, props, paraProps);
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

  function applyRunStyle(el, props, paraProps) {
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
    if (props.positionHalfPt) {
      // A raised or lowered run must not grow its line. Exact line spacing
      // clips it the way Word does, so shift the glyphs with a relative
      // offset, which is out of flow. Auto spacing keeps the vertical-align.
      if (paraProps && paraProps.lineRule === "exact") {
        styles.position = "relative";
        styles.top = (-props.positionHalfPt / 2) + "px";
      } else {
        styles.verticalAlign = (props.positionHalfPt / 2) + "px";
      }
    }
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
    // Media resolves when the picture is drawn, not when the document is
    // parsed, so a package only pays for the pages a reader sees.
    const url = run.url || (run.rid && model.mediaUrl ? model.mediaUrl(run.rid) : null);
    if (url) {
      const img = wrap.createEl("img");
      img.src = url;
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
  function renderVmlShape(parent, run, ctx) {
    renderVml(parent, run, {
      theme: model.theme,
      contentWidthPx,
      pageMarginLeftPx: twipToPx(marginLeftTw),
      pageMarginTopPx: twipToPx(marginTopTw),
      plannedLeft: (ctx && ctx.blockLeft) || 0,
      reserve: ctx && ctx.reserve ? ctx.reserve : null,
      takeReserve: ctx && ctx.takeReserve ? ctx.takeReserve : null,
      lineHeightPx: ctx && ctx.lineHeightPx ? ctx.lineHeightPx : 0,
      pageHost: currentPageContent,
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

  // The tallest text-relative VML box that reserves room in this paragraph.
  // The renderer draws it as one block spacer at the first such box, and the
  // page estimate uses the same height.
  function maxVmlReserve(runs) {
    let max = 0;
    const walk = (list) => {
      for (const run of list || []) {
        if (!run) continue;
        if (run.type === "vml" && (run.wrapStyle === "topandbottom" || run.wrapStyle === "through")
          && run.relativeV !== "page" && run.relativeV !== "margin") {
          const top = (run.offsetPx && run.offsetPx.top) || 0;
          const height = Number(run.heightPx) > 0 ? Number(run.heightPx) : 0;
          max = Math.max(max, top + height);
        } else if (run.type === "run") {
          walk(run.runs);
        } else if (run.type === "link") {
          walk(run.link.runs);
        }
      }
    };
    walk(runs);
    return max;
  }

  // True when a paragraph carries an out-of-flow drawing. The paragraph then
  // becomes the positioning context for the anchors inside it.
  function hasAnchoredRun(runs) {    for (const run of runs || []) {
      if (!run) continue;
      // A page-context anchor is positioned against the page content, so it
      // must not make the paragraph a containing block. A text-relative VML
      // box is positioned against its paragraph and does.
      if (run.anchor && run.anchor.outOfFlow && run.anchor.context !== "page") return true;
      if (run.type === "vml" && run.anchor && run.anchor.context !== "page") return true;
      if (run.type === "run" && hasAnchoredRun(run.runs)) return true;
      if (run.type === "link" && hasAnchoredRun(run.link.runs)) return true;
    }
    return false;
  }

  function flowWidth(run) {
    if (run.type === "vml") {
      if (run.anchor && run.anchor.context === "page") return 0;
      return Number(run.widthPx) > 0 ? Number(run.widthPx) : 0;
    }
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
    mapBlock(block, table);
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
    const text = (el.textContent || "").replace(/\s+/g, " ").trim();
    if (text) textIndex.push({ block: null, page: totalPages, text, lower: text.toLowerCase(), el });
  }

  function scrollToNote(kind, id) {
    const target = pagesEl.querySelector('.ov-docx-note[data-note-id="' + cssEscape(id) + '"]');
    if (target) target.scrollIntoView({ block: "center", behavior: "smooth" });
  }

  function scrollToAnchor(name) {
    const target = anchorBlocks.get(name);
    if (!target) return false;
    const record = pageRecords[target.page - 1];
    if (record && !record.rendered) renderPage(record);
    const el = blockElements.get(target.block);
    if (el) el.scrollIntoView({ block: "start", behavior: "smooth" });
    return Boolean(el);
  }

  // ---------- search ----------

  function search(query) {
    clearSearch();
    const q = String(query || "").toLowerCase();
    if (!q) return 0;
    searchHits = [];
    for (const item of textIndex) {
      if (item.lower.indexOf(q) !== -1) searchHits.push(item);
    }
    searchBlockSet = new Set(searchHits.map((item) => item.block).filter(Boolean));
    for (const item of searchHits) {
      const el = item.el || blockElements.get(item.block);
      if (el) el.addClass("ov-search-hit");
    }
    return searchHits.length;
  }

  function clearSearch() {
    for (const item of searchHits) {
      const el = item.el || blockElements.get(item.block);
      if (el) el.removeClass("ov-search-hit");
    }
    if (searchCurrentEl) {
      searchCurrentEl.removeClass("ov-search-current");
      searchCurrentEl = null;
    }
    searchHits = [];
    searchBlockSet = new Set();
    searchIndex = -1;
  }

  function searchNext(dir) {
    if (!searchHits.length) return null;
    searchIndex = (searchIndex + dir + searchHits.length * 2) % searchHits.length;
    const item = searchHits[searchIndex];
    const record = pageRecords[item.page - 1];
    if (record && !record.rendered) renderPage(record);
    const el = item.el || blockElements.get(item.block);
    if (searchCurrentEl) searchCurrentEl.removeClass("ov-search-current");
    searchCurrentEl = el || null;
    if (el) {
      el.addClass("ov-search-current");
      el.scrollIntoView({ block: "center", behavior: "smooth" });
    }
    return { index: searchIndex + 1, count: searchHits.length };
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

  // Only the text that flows in the paragraph: the text inside a drawing's own
  // text box is placed by the drawing, so it must not be counted as paragraph
  // lines by the pagination estimate.
  function flowText(runs) {
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
            const anchor = (block.bookmarks && block.bookmarks[0]) || headingNames.get(block) || undefined;
            outline.push({ text, level: Math.min(4, (props.outlineLevel || 0) + 1), anchor });
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
    const text = flowText(runs).replace(/\u00ad/g, "");
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
    let reserve = 0;
    let imageHeight = 0;
    for (const run of runs) {
      const blocks = run.type === "run" ? run.runs : [run];
      for (const inner of blocks || []) {
        if (!inner) continue;
        if (inner.type === "image" || inner.type === "textbox" || inner.type === "shapegroup" || inner.type === "vml") {
          // A DrawingML overlay takes no room in the page plan; an in-flow VML
          // box does, and its height is what spaces the page.
          if (inner.anchor && inner.anchor.outOfFlow) continue;
          if (inner.type === "vml") {
            // Match the renderer: only a top-and-bottom or through box
            // reserves its height. A float or a bare overlay adds nothing.
            if (inner.wrapStyle !== "topandbottom" && inner.wrapStyle !== "through") continue;
            if (inner.relativeV === "page" || inner.relativeV === "margin") continue;
          }
          const top = (inner.offsetPx && inner.offsetPx.top) || 0;
          const height = Number(inner.heightPx) > 0 ? Number(inner.heightPx) + top : 0;
          if (inner.type === "vml") reserve = Math.max(reserve, height);
          else imageHeight = Math.max(imageHeight, height);
        }
      }
    }
    if (reserve > 0) {
      // One block spacer reserves the tallest box; the paragraph's own lines
      // then follow it, and the first of them sits inside the reserved band.
      contentHeight = reserve + Math.max(0, lines - 1) * linePx;
    }
    if (imageHeight > 0) {
      const imageLines = Math.ceil(imageHeight / Math.max(1, linePx));
      contentHeight = Math.max(contentHeight, imageHeight + 4, imageLines * linePx);
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
    // Each page keeps the height its blocks need, so a page that is not
    // rendered yet can hold the scroll position with a placeholder.
    const nextPage = () => {
      pages.push({ blocks: current, height: used });
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
      // A new section starts on a fresh page unless the section says it is
      // continuous. Word stores that in the new section's own sectPr, and the
      // default when it is absent is nextPage.
      const sectionIndex = block.section != null ? block.section : 0;
      const previous = i > 0 ? blocks[i - 1] : null;
      const previousSection = previous ? (previous.section != null ? previous.section : 0) : -1;
      if (previous && sectionIndex !== previousSection) {
        const sectionProps = (model.sectionList && model.sectionList[sectionIndex]) || {};
        if ((sectionProps.type || "nextPage") !== "continuous" && (current.length || used > 0)) {
          nextPage();
        }
      }
      if (forced && (current.length || used > 0)) nextPage();
      // keepNext: this block must share its page with the one after it.
      if (block.props && block.props.keepNext && current.length
        && used + total + nextTotal > contentHeightPx && used + total <= contentHeightPx) {
        nextPage();
      }
      if (used > 0 && used + total > contentHeightPx) {
        // A short lead-in (a form's title, a heading) travels with a table that
        // would otherwise start a page alone, which is what Word and the
        // reference renderers do with a dispensing form.
        const last = current.length ? current[current.length - 1] : null;
        const lastMetrics = last ? blockMetrics(last) : null;
        const keepWithTable = last && last.type === "p" && block.type === "table"
          && total > contentHeightPx * 0.4
          && (headingBlock(last) || (lastMetrics && lastMetrics.lines <= 1));
        if (keepWithTable) {
          current.pop();
          used -= blockTotal(last);
          nextPage();
          current.push(last);
          used += blockTotal(last);
        } else {
          nextPage();
        }
      }
      // The block's offset in the page content box, which VML overlays and
      // other anchored content need to place themselves without measuring.
      block.plannedTop = used + metrics.before;
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
    // The first report measures every page. In windowed mode that is deferred
    // to the next frame so the mount still paints immediately.
    if (settings.virtualize === true) {
      const raf = typeof requestAnimationFrame === "function" ? requestAnimationFrame : (cb) => setTimeout(cb, 16);
      raf(() => reportCurrentPage());
    } else {
      reportCurrentPage();
    }
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
    if (!pagesEl) return out;
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
    if (destroyed || !pagesEl) return;
    const count = pagesInOrder().length;
    if (!count) return;
    const page = currentPageNumber();
    if (page !== lastReportedPage || lastReportedPage === 0) {
      lastReportedPage = page;
      callbacks.onPageChange({ page, count });
    }
    renderPageWindow(page);
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
    pageRecords = [];
    blockElements.clear();
    anchorBlocks.clear();
    headingNames.clear();
    textIndex = [];
    searchHits = [];
    searchBlockSet = new Set();
    searchCurrentEl = null;
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
