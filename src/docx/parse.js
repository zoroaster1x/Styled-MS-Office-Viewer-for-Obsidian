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

// Parses a WordprocessingML document into the block and inline model the
// renderer walks. Both .docx and .docm share this format; the only difference
// is the macro part, which is never read.

const { Package, relTypeIs } = require("../shared/package");
const {
  childrenOf, firstOf, lastOf, findAll, findFirst, attr, attrInt, attrBool, attrNs, flag, tagName,
} = require("../shared/xml");
const { createMediaCache, mediaUrl } = require("../media/media");
const { parseMath } = require("./math");
const { parseVmlRun } = require("./vml");
const styles = require("./styles");

function parseDocx(input) {
  const pkg = Package.open(input);
  const mainPath = findMainPart(pkg);
  if (!mainPath) {
    throw new Error("This file has no word/document.xml part, so it is not a readable Word document.");
  }
  const documentXml = pkg.xml(mainPath);
  if (!documentXml) {
    throw new Error("word/document.xml could not be parsed.");
  }
  const themeDoc = pkg.xml("word/theme/theme1.xml");
  const { parseTheme } = require("../shared/color");
  const theme = parseTheme(themeDoc);
  const styleTable = styles.parseStyles(pkg.xml("word/styles.xml"), theme);
  const numbering = styles.parseNumbering(pkg.xml("word/numbering.xml"), theme);
  const rels = pkg.rels(mainPath);
  const mediaCache = createMediaCache();

  const model = {
    kind: "docx",
    pkg,
    partPath: mainPath,
    theme,
    styles: styleTable,
    numbering,
    rels,
    mediaCache,
    mediaUrl(relId) {
      const rel = rels.get(relId);
      if (!rel) return null;
      const abs = rel.mode === "External" ? rel.target : pkg.resolve(mainPath, rel);
      if (!abs) return null;
      if (rel.mode === "External") return null;
      return mediaUrl(pkg, abs, mediaCache);
    },
    footnotes: parseNotes(pkg, "word/footnotes.xml", theme),
    endnotes: parseNotes(pkg, "word/endnotes.xml", theme),
    headers: parseHeaderFooterParts(pkg, mainPath, "header"),
    footers: parseHeaderFooterParts(pkg, mainPath, "footer"),
    hyperlinks: new Map(),
  };

  // Bookmark targets for internal links.
  const bookmarkTargets = new Map();
  collectBookmarks(documentXml.documentElement, bookmarkTargets);
  model.bookmarks = bookmarkTargets;

  const body = firstOf(documentXml.documentElement, "body");
  if (!body) throw new Error("The document body is missing.");
  const bodyCtx = { listCounters: new Map(), sectionIndex: 0, sections: [] };
  model.body = parseBlockChildren(body, model, bodyCtx);
  model.section = parseFinalSection(body, model);
  // Section properties in document order: the sectPr ending each section, then
  // the body's own final sectPr. A header or footer reference is per section.
  model.sectionList = bodyCtx.sections;
  model.sectionList[bodyCtx.sections.length] = model.section;
  model.sections = countSections(body);
  // settings.xml says whether even and odd pages carry different headers and
  // footers. Without it every page uses the default part, which is what Word
  // does, even when the section lists an even part.
  const settingsDoc = pkg.xml("word/settings.xml");
  model.settings = {
    evenAndOddHeaders: settingsDoc ? Boolean(findFirst(settingsDoc.documentElement, "evenAndOddHeaders")) : false,
  };
  model.properties = parseCoreProperties(pkg.xml("docProps/core.xml"), pkg.xml("docProps/app.xml"));
  return model;
}

function findMainPart(pkg) {
  const rels = pkg.rels("");
  for (const rel of rels.values()) {
    if (relTypeIs(rel, "officeDocument")) return pkg.resolve("", rel);
  }
  if (pkg.has("word/document.xml")) return "word/document.xml";
  return null;
}

// ---------- notes ----------

function parseNotes(pkg, path, theme) {
  const out = new Map();
  const doc = pkg.xml(path);
  if (!doc) return out;
  for (const note of doc.documentElement.children || []) {
    const tag = tagName(note);
    if (tag !== "footnote" && tag !== "endnote") continue;
    const id = attr(note, "id");
    if (id === "-1" || id === "0") continue;
    const type = attr(note, "type");
    if (type === "separator" || type === "continuationSeparator") continue;
    out.set(id, {
      id,
      type: tag,
      blocks: parseBlockChildren(note, { theme, styles: null, numbering: null, rels: new Map(), pkg, partPath: path, mediaCache: createMediaCache() }, { listCounters: new Map() }),
    });
  }
  return out;
}

// ---------- headers and footers ----------

function parseHeaderFooterParts(pkg, mainPath, kind) {
  const out = new Map();
  const rels = pkg.rels(mainPath);
  for (const [rid, rel] of rels) {
    if (!relTypeIs(rel, kind)) continue;
    const path = pkg.resolve(mainPath, rel);
    if (!path) continue;
    out.set(rid, { path, blocks: null });
  }
  return out;
}

// ---------- blocks ----------

function parseBlockChildren(parent, model, ctx, depth) {
  const out = [];
  const level = (depth || 0) + 1;
  // A table inside a table inside a table is legitimate; a thousand of them is
  // a stack overflow, so nesting stops at a depth no real document reaches.
  if (level > 12) return out;
  let counted = 0;
  for (const el of parent.children || []) {
    if (counted++ > 50000) break;
    const tag = tagName(el);
    if (tag === "p") {
      const parsed = parseParagraph(el, model, ctx);
      // Content up to and including a paragraph that carries a sectPr belongs
      // to that section; the sections after it get the next index.
      for (const block of splitAtPageBreaks(parsed)) {
        block.section = ctx.sectionIndex || 0;
        out.push(block);
      }
      if (parsed.sectionBreak && ctx.sections) {
        ctx.sections[ctx.sectionIndex || 0] = parsed.sectionBreak;
        ctx.sectionIndex = (ctx.sectionIndex || 0) + 1;
      }
    } else if (tag === "tbl") {
      const table = parseTable(el, model, ctx, level);
      table.section = ctx.sectionIndex || 0;
      out.push(table);
    } else if (tag === "sdt") {
      const content = firstOf(el, "sdtContent");
      if (content) out.push(...parseBlockChildren(content, model, ctx, level));
    } else if (tag === "sectPr") {
      // Captured by the caller as the section break that ends this block run.
    } else if (tag === "bookmarkStart") {
      // Handled through the bookmark table.
    }
  }
  return out;
}

// A page break in the middle of a paragraph (or, as Word usually writes it, in
// a paragraph of its own) ends the page there. The block is split so the page
// planner can honour both halves.
function splitAtPageBreaks(block) {
  if (!block || block.type !== "p" || !block.runs || !block.runs.length) return [block];
  if (!block.runs.some((run) => containsPageBreak(run))) return [block];
  const out = [];
  let current = [];
  const flush = (breakAfter) => {
    out.push(Object.assign({}, block, { runs: current, pageBreakAfter: breakAfter }));
    current = [];
  };
  for (const run of block.runs) {
    if (isPageBreak(run)) {
      flush(true);
      continue;
    }
    if (containsPageBreak(run)) {
      for (const inner of run.runs || []) {
        if (isPageBreak(inner)) flush(true);
        else current.push(inner);
      }
      continue;
    }
    current.push(run);
  }
  if (current.length || !out.length) flush(false);
  return out;
}

function isPageBreak(run) {
  return Boolean(run && run.type === "break" && run.page);
}

function containsPageBreak(run) {
  if (!run) return false;
  if (isPageBreak(run)) return true;
  if (run.type === "run") return (run.runs || []).some(containsPageBreak);
  return false;
}

function parseParagraph(el, model, ctx) {
  const pPrEl = firstOf(el, "pPr");
  const pPr = styles.parseParaProps(pPrEl, model.theme);
  const para = {
    type: "p",
    props: pPr,
    style: pPr && pPr.styleId ? pPr.styleId : null,
    runs: [],
    markRunProps: null,
    numbering: null,
    bookmarks: [],
    sectionBreak: null,
  };
  const pPrRPr = firstOf(pPrEl, "rPr");
  if (pPrRPr) para.markRunProps = styles.parseRunProps(pPrRPr, model.theme);

  if (pPr) {
    if (pPr.numId != null && model.numbering && model.numbering.numIdToAbstract.size) {
      para.numbering = resolveNumbering(pPr.numId, pPr.ilvl || 0, model, ctx);
    }
    const sectPrEl = firstOf(pPrEl, "sectPr");
    if (sectPrEl) para.sectionBreak = styles.parseSectionProps(sectPrEl, model.theme);
  }

  // Runs, hyperlinks, bookmarks, fields and drawings in document order.
  let hyperlink = null;
  for (const child of el.children || []) {
    const tag = tagName(child);
    if (tag === "r") {
      parseRunInto(para.runs, child, model, ctx, hyperlink);
    } else if (tag === "hyperlink") {
      const rid = attr(child, "id") || attrNs(child, "id");
      const anchor = attr(child, "anchor");
      const target = rid ? resolveHyperlink(model, rid) : null;
      const inner = { href: target ? target.href : null, anchor: anchor || (target && target.anchor) || null, runs: [] };
      for (const r of childrenOf(child, "r")) parseRunInto(inner.runs, r, model, ctx, null);
      para.runs.push({ type: "link", link: inner });
    } else if (tag === "oMath" || tag === "oMathPara") {
      // Equations are content, not decoration: an unread structure would drop
      // the numbers and the operators with them.
      const node = parseMath(child);
      if (node) para.runs.push({ type: "math", node });
    } else if (tag === "bookmarkStart") {
      const name = attr(child, "name");
      if (name && !name.startsWith("_")) para.bookmarks.push(name);
    } else if (tag === "sdt") {
      // Content controls nest, and a page number inside one is still content.
      parseInlineContainer(child, para.runs, model, ctx);
    } else if (tag === "fldSimple") {
      const instr = attr(child, "instr") || "";
      const isPage = /PAGE/.test(instr) && !/NUMPAGES/.test(instr);
      const runs = [];
      for (const r of childrenOf(child, "r")) parseRunInto(runs, r, model, ctx, null);
      para.runs.push({ type: "field", kind: isPage ? "page" : /NUMPAGES/.test(instr) ? "pages" : "field", runs });
    } else if (tag === "smartTag" || tag === "ins" || tag === "sdt") {
      // These wrappers nest freely around runs and other inline content, so
      // the walk has to recurse: a w:r two levels down still holds text.
      parseInlineContainer(child, para.runs, model, ctx);
    }
  }
  void hyperlink;
  return para;
}

// Walks an inline wrapper (smartTag, ins, sdt) for the content it carries.
function parseInlineContainer(el, list, model, ctx) {
  for (const child of el.children || []) {
    const tag = tagName(child);
    if (tag === "r") {
      parseRunInto(list, child, model, ctx, null);
    } else if (tag === "smartTag" || tag === "ins" || tag === "sdt" || tag === "sdtContent") {
      parseInlineContainer(child, list, model, ctx);
    } else if (tag === "hyperlink") {
      const rid = attr(child, "id") || attrNs(child, "id");
      const target = rid ? resolveHyperlink(model, rid) : null;
      const link = { href: target ? target.href : null, anchor: attr(child, "anchor"), runs: [] };
      for (const r of childrenOf(child, "r")) parseRunInto(link.runs, r, model, ctx, null);
      list.push({ type: "link", link });
    } else if (tag === "oMath" || tag === "oMathPara") {
      const node = parseMath(child);
      if (node) list.push({ type: "math", node });
    }
  }
}

function resolveHyperlink(model, rid) {
  const rel = model.rels.get(rid);
  if (!rel) return null;
  if (rel.mode === "External") return { href: rel.target, anchor: null };
  const target = model.pkg.resolve(model.partPath, rel);
  if (target && target.indexOf("file://") === 0) return { href: target, anchor: null };
  return { href: null, anchor: null };
}

function parseRunInto(list, el, model, ctx, parentLink) {
  const run = parseRun(el, model, ctx);
  if (!run) return;
  if (parentLink) parentLink.runs.push(run);
  else if (run.runs && run.isLink) list.push(run);
  else list.push(run);
}

function parseRun(el, model, ctx) {
  const rPrEl = firstOf(el, "rPr");
  const rPr = styles.parseRunProps(rPrEl, model.theme);
  const runs = [];
  for (const child of el.children || []) {
    const tag = tagName(child);
    if (tag === "t") {
      runs.push({ type: "text", text: child.textContent || "", props: rPr });
    } else if (tag === "delText") {
      runs.push({ type: "text", text: child.textContent || "", props: rPr, deleted: true });
    } else if (tag === "br") {
      const kind = attr(child, "type");
      if (kind === "page") runs.push({ type: "break", page: true });
      else if (kind === "column") runs.push({ type: "break", column: true });
      else runs.push({ type: "break" });
    } else if (tag === "cr") {
      runs.push({ type: "break" });
    } else if (tag === "tab") {
      runs.push({ type: "tab" });
    } else if (tag === "sym") {
      runs.push({ type: "text", text: symbolFromChar(attr(child, "char"), attr(child, "font")), props: rPr });
    } else if (tag === "noBreakHyphen") {
      runs.push({ type: "text", text: "\u2011", props: rPr });
    } else if (tag === "softHyphen") {
      runs.push({ type: "text", text: "\u00ad", props: rPr });
    } else if (tag === "drawing") {
      const drawing = parseDrawing(child, model, ctx);
      if (drawing) runs.push(drawing);
    } else if (tag === "pict" || tag === "object") {
      const vml = parseVmlRun(child);
      if (vml) {
        runs.push(vml);
      } else {
        const picture = parseLegacyPicture(child, model, ctx);
        if (picture) runs.push(picture);
        else runs.push({ type: "placeholder", kind: "object", label: "Embedded object" });
      }
    } else if (tag === "fldChar") {
      const type = attr(child, "fldCharType");
      runs.push({ type: "fieldChar", stage: type });
    } else if (tag === "instrText") {
      runs.push({ type: "instr", text: child.textContent || "" });
    } else if (tag === "footnoteReference" || tag === "endnoteReference") {
      const id = attr(child, tag === "footnoteReference" ? "id" : "id");
      runs.push({ type: "noteRef", noteId: id, kind: tag === "footnoteReference" ? "footnote" : "endnote", props: rPr });
    } else if (tag === "commentReference") {
      // Comments are annotations, not content; a small marker keeps the text
      // honest without inventing the comment body.
      runs.push({ type: "commentRef", id: attr(child, "id") });
    } else if (tag === "AlternateContent") {
      // Word marks newer shapes as a Choice with an older Fallback beside it.
      // The first branch that yields anything wins, which keeps the picture a
      // Choice carries instead of dropping the whole element.
      const choice = firstOf(child, "Choice");
      const fallback = firstOf(child, "Fallback");
      for (const branch of [choice, fallback]) {
        if (!branch) continue;
        const inner = parseRun(branch, model, ctx);
        if (inner) {
          if (inner.type === "run" && inner.runs) runs.push(...inner.runs);
          else runs.push(inner);
          break;
        }
      }
    } else if (tag === "rt") {
      // Ruby text; the base text is already emitted.
    }
  }
  // Field runs collapse into a single field node when they only carry a value.
  if (runs.length === 1 && runs[0].type === "text" && !runs[0].deleted) {
    return { type: "text", text: runs[0].text, props: rPr };
  }
  if (!runs.length) return null;
  // A run with one child is that child: a wrapper would hide a tab from the
  // content after it, which is what a right tab measures.
  if (runs.length === 1) return runs[0];
  return { type: "run", runs, props: rPr, isLink: false };
}

// w:sym addresses a glyph in a legacy symbol font by a private-use code point:
// F0xx is the old byte xx in that font's encoding. Passing the code point
// straight through draws a blank box, so the common symbol fonts are mapped to
// the Unicode characters they mean.
const SYMBOL_FONT_MAP = {
  symbol: {
    0x22: "\u2200", 0x24: "\u2203", 0x27: "\u220b", 0x2a: "\u2217", 0x2d: "\u2212",
    0x40: "\u2245",
    0x41: "\u0391", 0x42: "\u0392", 0x43: "\u03a7", 0x44: "\u0394", 0x45: "\u0395",
    0x46: "\u03a6", 0x47: "\u0393", 0x48: "\u0397", 0x49: "\u0399", 0x4a: "\u03d1",
    0x4b: "\u039a", 0x4c: "\u039b", 0x4d: "\u039c", 0x4e: "\u039d", 0x4f: "\u039f",
    0x50: "\u03a0", 0x51: "\u0398", 0x52: "\u03a1", 0x53: "\u03a3", 0x54: "\u03a4",
    0x55: "\u03a5", 0x56: "\u03c2", 0x57: "\u03a9", 0x58: "\u039e", 0x59: "\u03a8",
    0x5a: "\u0396", 0x5c: "\u2234", 0x5e: "\u22a5",
    0x60: "\u203e", 0x61: "\u03b1", 0x62: "\u03b2", 0x63: "\u03c7", 0x64: "\u03b4",
    0x65: "\u03b5", 0x66: "\u03c6", 0x67: "\u03b3", 0x68: "\u03b7", 0x69: "\u03b9",
    0x6a: "\u03d5", 0x6b: "\u03ba", 0x6c: "\u03bb", 0x6d: "\u03bc", 0x6e: "\u03bd",
    0x6f: "\u03bf", 0x70: "\u03c0", 0x71: "\u03b8", 0x72: "\u03c1", 0x73: "\u03c3",
    0x74: "\u03c4", 0x75: "\u03c5", 0x76: "\u03d6", 0x77: "\u03c9", 0x78: "\u03be",
    0x79: "\u03c8", 0x7a: "\u03b6", 0x7b: "{", 0x7c: "|", 0x7d: "}", 0x7e: "\u223c",
    0xa0: "\u20ac", 0xa1: "\u03d2", 0xa2: "\u2032", 0xa3: "\u2264", 0xa4: "\u2044",
    0xa5: "\u221e", 0xa6: "\u0192", 0xa7: "\u2663", 0xa8: "\u2666", 0xa9: "\u2665",
    0xaa: "\u2660", 0xab: "\u2194", 0xac: "\u2190", 0xad: "\u2191", 0xae: "\u2192",
    0xaf: "\u2193", 0xb0: "\u00b0", 0xb1: "\u00b1", 0xb2: "\u2033", 0xb3: "\u2265",
    0xb4: "\u00d7", 0xb5: "\u221d", 0xb6: "\u2202", 0xb7: "\u2022", 0xb8: "\u00f7",
    0xb9: "\u2260", 0xba: "\u2261", 0xbb: "\u2248", 0xbc: "\u2026", 0xbe: "\u23af",
    0xbf: "\u21b5", 0xc0: "\u2135", 0xc1: "\u2111", 0xc2: "\u211c", 0xc3: "\u2118",
    0xc4: "\u2297", 0xc5: "\u2295", 0xc6: "\u2205", 0xc7: "\u2229", 0xc8: "\u222a",
    0xc9: "\u2283", 0xca: "\u2287", 0xcb: "\u2284", 0xcc: "\u2282", 0xcd: "\u2286",
    0xce: "\u2208", 0xcf: "\u2209", 0xd0: "\u2220", 0xd1: "\u2207", 0xd5: "\u220f",
    0xd6: "\u221a", 0xd7: "\u22c5", 0xd8: "\u00ac", 0xd9: "\u2227", 0xda: "\u2228",
    0xdb: "\u21d4", 0xdc: "\u21d0", 0xdd: "\u21d1", 0xde: "\u21d2", 0xdf: "\u21d3",
    0xe0: "\u25ca", 0xe1: "\u2329", 0xe5: "\u2211", 0xf1: "\u232a", 0xf2: "\u222b",
    0xf3: "\u2320", 0xf4: "\u23ae", 0xf5: "\u2321", 0xf6: "\u23af", 0xfe: "\u25a0",
  },
  wingdings: {
    0x2a: "\u261b", 0x2b: "\u261e", 0x2d: "\u270d", 0x2e: "\u270e", 0x2f: "\u270f",
    0x36: "\u2714", 0x37: "\u2718", 0x38: "\u2720", 0x39: "\u2726", 0x3a: "\u2605",
    0x3b: "\u2606", 0x3c: "\u2736", 0x3f: "\u2739",
    0x4c: "\u25cf", 0x4d: "\u274d", 0x4e: "\u25a0", 0x4f: "\u25a1", 0x50: "\u2751",
    0x51: "\u2752", 0x52: "\u25b2", 0x53: "\u25bc", 0x54: "\u25c6", 0x55: "\u2756",
    0x56: "\u2605", 0x57: "\u2735", 0x58: "\u2734", 0x59: "\u2739",
    0x6c: "\u25cf", 0x6d: "\u274d", 0x6e: "\u25a0", 0x6f: "\u25a1", 0x70: "\u2751",
    0x71: "\u2752", 0x72: "\u25b2", 0x73: "\u25bc", 0x74: "\u25c6", 0x75: "\u2756",
    0x76: "\u2605", 0x77: "\u2736", 0x78: "\u2734", 0x79: "\u2739",
    0x9f: "\u2022", 0xa1: "\u261c", 0xa2: "\u261e", 0xa3: "\u261d", 0xa4: "\u261f",
    0xa5: "\u2763", 0xa6: "\u2764", 0xa7: "\u2764", 0xa8: "\u2765", 0xa9: "\u2766",
    0xaa: "\u2767", 0xab: "\u2660", 0xac: "\u2663", 0xad: "\u2665", 0xae: "\u2666",
    0xaf: "\u2022", 0xb0: "\u25d6", 0xb1: "\u25d7", 0xb2: "\u25d0", 0xb3: "\u25d1",
    0xb4: "\u25d3", 0xb5: "\u25d2", 0xb6: "\u25d5", 0xb7: "\u25d4", 0xb8: "\u25d8",
    0xb9: "\u25d9", 0xba: "\u263a", 0xbb: "\u263b", 0xbc: "\u2639",
    0xd8: "\u2191", 0xd9: "\u2193", 0xda: "\u2192", 0xdb: "\u2190",
    0xdc: "\u21d1", 0xdd: "\u21d3", 0xde: "\u21d2", 0xdf: "\u21d0",
    0xe0: "\u21e7", 0xe1: "\u21e9", 0xe2: "\u21e8", 0xe3: "\u21e6",
    0xe8: "\u2600", 0xe9: "\u2601", 0xea: "\u2602", 0xeb: "\u2603", 0xec: "\u2604",
    0xed: "\u2605", 0xee: "\u2606", 0xef: "\u2607", 0xf0: "\u2608",
    0xfc: "\u2713", 0xfd: "\u2717", 0xfe: "\u2612",
  },
  webdings: {
    0x21: "\u2713", 0x22: "\u2714", 0x25: "\u2605", 0x6f: "\u25a1", 0x70: "\u25a0",
    0x72: "\u25b2", 0x73: "\u25bc", 0x74: "\u25c6", 0x75: "\u2756", 0x76: "\u2605",
    0x77: "\u2606", 0xa1: "\u2708", 0xa2: "\u2709", 0xa5: "\u270e", 0xac: "\u2666",
    0xab: "\u2663", 0xa9: "\u2665", 0xaa: "\u2660",
  },
};

function symbolFromChar(char, font) {
  if (!char) return "";
  const code = parseInt(char, 16);
  if (isNaN(code)) return "";
  const key = String(font || "").toLowerCase().replace(/\s+/g, "");
  const table = SYMBOL_FONT_MAP[key];
  if (table) {
    const low = code & 0xff;
    if (table[low]) return table[low];
  }
  try {
    return String.fromCodePoint(code);
  } catch (err) {
    return "";
  }
}

// ---------- drawings ----------

function parseDrawing(el, model, ctx) {
  const inline = firstOf(el, "inline");
  const anchor = firstOf(el, "anchor");
  const container = inline || anchor;
  if (!container) return null;
  const extent = firstOf(container, "extent");
  const cx = extent ? attrInt(extent, "cx", 0) : 0;
  const cy = extent ? attrInt(extent, "cy", 0) : 0;
  const anchorInfo = anchor ? parseAnchor(anchor) : null;
  const graphic = firstOf(container, "graphic");  if (!graphic) return null;
  const graphicData = firstOf(graphic, "graphicData");
  if (!graphicData) return null;
  const docPrEarly = firstOf(container, "docPr");
  const shapeName = docPrEarly ? attr(docPrEarly, "name") : "";
  const altText = docPrEarly ? (attr(docPrEarly, "descr") || shapeName) : "";
  let blip = null;
  for (const child of graphicData.children || []) {
    const tag = tagName(child);
    if (tag === "pic" || tag === "picture") {
      blip = findDescendant(child, "blip");
    } else if (tag === "wsp" || tag === "wps" || tag === "shape") {
      // A shape can hold a text box, or stand alone as a filled outline (an
      // anchored trapezoid, a bar, a callout). Word positions it absolutely,
      // but losing it would be worse than placing it inline, so it becomes a
      // boxed element of its own runs or its own silhouette.
      const textbox = findDescendant(child, "txbxContent");
      const spPr = firstOf(child, "spPr");
      const geometry = spPr ? (firstOf(spPr, "prstGeom") || firstOf(spPr, "custGeom")) : null;
      if (!textbox && geometry && cx && cy) {
        const texts = [];
        for (const t of findAll(child, "t")) {
          const value = t.textContent || "";
          if (value.trim()) texts.push(value);
        }
        return {
          type: "shapegroup",
          node: child,
          widthPx: Math.round((cx / 9525) * 100) / 100,
          heightPx: Math.round((cy / 9525) * 100) / 100,
          name: shapeName,
          alt: altText,
          texts,
          anchor: anchorInfo,
        };
      }
      if (!textbox) return null;
      const runs = [];
      for (const paragraph of childrenOf(textbox, "p")) {
        for (const inner of childrenOf(paragraph, "r")) {
          const run = parseRun(inner, model, ctx);
          if (run) runs.push(run);
        }
        runs.push({ type: "break" });
      }
      if (!runs.length) return null;
      return {
        type: "textbox",
        runs,
        widthPx: cx ? cx / 9525 : 0,
        heightPx: cy ? cy / 9525 : 0,
        name: shapeName,
        alt: altText,
        anchor: anchorInfo,
      };
    } else if (tag === "chart" || tag === "chartex") {
      return { type: "placeholder", kind: "chart", label: "Chart", widthPx: cx / 9525, heightPx: cy / 9525 };
    } else if (tag === "tbl" || tag === "table") {
      return null;
    } else if (tag === "relIds") {
      return { type: "placeholder", kind: "smartart", label: "SmartArt diagram", widthPx: cx / 9525, heightPx: cy / 9525 };
    } else if (tag === "wgp") {
      // A Word drawing group: text boxes, pictures, freeform lines and
      // connectors in one inline box. The renderer walks the group at draw
      // time; the plain text is collected here for search, counts and the
      // fidelity sweep.
      if (!cx || !cy) return null;
      const texts = [];
      for (const t of findAll(child, "t")) {
        const value = t.textContent || "";
        if (value.trim()) texts.push(value);
      }
      return {
        type: "shapegroup",
        node: child,
        widthPx: Math.round((cx / 9525) * 100) / 100,
        heightPx: Math.round((cy / 9525) * 100) / 100,
        name: shapeName,
        alt: altText,
        texts,
        anchor: anchorInfo,
      };
    }
  }
  if (!blip) return null;
  const rid = attr(blip, "embed") || attrNs(blip, "embed") || attr(blip, "link");
  if (!rid) return null;
  const url = model.mediaUrl(rid);
  if (!url) return null;
  const docPr = firstOf(container, "docPr");
  const name = docPr ? attr(docPr, "name") : "";
  const alt = docPr ? (attr(docPr, "descr") || name) : "";
  const geometry = parseDrawingGeometry(container);
  return {
    type: "image",
    rid,
    url,
    widthPx: Math.round((cx / 9525) * 100) / 100,
    heightPx: Math.round((cy / 9525) * 100) / 100,
    name,
    alt,
    rotation: 0,
    anchor: anchorInfo,
    wrap: parseWrap(anchor),
    geometry,
  };
}

// A wp:anchor carries its own position. wrapNone objects take no space in the
// line: they are placed from these offsets, so a nose-bridge trapezoid sits
// under the ruler instead of at the start of its paragraph.
function parseAnchor(el) {
  const readPosition = (pos, fallbackFrom) => {
    if (!pos) return { from: fallbackFrom, offset: null, align: null };
    const offsetEl = firstOf(pos, "posOffset");
    const alignEl = firstOf(pos, "align");
    let offset = null;
    if (offsetEl) {
      // wp:posOffset carries the value as its text, not an attribute.
      const value = parseInt(offsetEl.textContent || "", 10);
      if (!isNaN(value)) offset = value;
    }
    return {
      from: attr(pos, "relativeFrom") || fallbackFrom,
      offset,
      align: alignEl ? attr(alignEl, "val") : null,
    };
  };
  const h = readPosition(firstOf(el, "positionH"), "column");
  const v = readPosition(firstOf(el, "positionV"), "paragraph");
  const wrap = parseWrap(el);
  return {
    h,
    v,
    behindDoc: attrBool(el, "behindDoc", false),
    wrap,
    outOfFlow: wrap === "wrapNone",
  };
}

function parseWrap(anchor) {
  if (!anchor) return null;
  for (const child of anchor.children || []) {
    const tag = tagName(child);
    if (tag === "wrapSquare" || tag === "wrapTight" || tag === "wrapThrough"
      || tag === "wrapTopAndBottom" || tag === "wrapNone") {
      return tag;
    }
  }
  return null;
}

function parseDrawingGeometry(container) {
  const xfrm = findDescendant(findDescendant(container, "spPr"), "xfrm") || findDescendant(container, "xfrm");
  if (!xfrm) return null;
  const off = firstOf(xfrm, "off");
  const ext = firstOf(xfrm, "ext");
  const rot = attrInt(xfrm, "rot", 0);
  return {
    xPx: off ? attrInt(off, "x", 0) / 9525 : 0,
    yPx: off ? attrInt(off, "y", 0) / 9525 : 0,
    cxPx: ext ? attrInt(ext, "cx", 0) / 9525 : 0,
    cyPx: ext ? attrInt(ext, "cy", 0) / 9525 : 0,
    rotationDeg: rot / 60000,
  };
}

function parseLegacyPicture(el, model) {
  const shape = findDescendant(el, "shape");
  const imagedata = findDescendant(el, "imagedata");
  if (!imagedata) return null;
  const rid = attrNs(imagedata, "id") || attr(imagedata, "id") || attrNs(imagedata, "embed");
  if (!rid) return null;
  const url = model.mediaUrl(rid);
  if (!url) return null;
  const style = shape ? attr(shape, "style") || "" : "";
  const widthMatch = /width:([\d.]+)pt/.exec(style);
  const heightMatch = /height:([\d.]+)pt/.exec(style);
  return {
    type: "image",
    rid,
    url,
    widthPx: widthMatch ? Math.round(parseFloat(widthMatch[1]) * 96 / 72) : 0,
    heightPx: heightMatch ? Math.round(parseFloat(heightMatch[1]) * 96 / 72) : 0,
    name: shape ? attr(shape, "alt") || "" : "",
    alt: shape ? attr(shape, "alt") || "" : "",
    rotation: 0,
    anchor: null,
    wrap: null,
    geometry: null,
  };
}

function findDescendant(el, tag) {
  if (!el) return null;
  for (const child of el.children || []) {
    if (tagName(child) === tag) return child;
    const found = findDescendant(child, tag);
    if (found) return found;
  }
  return null;
}

// ---------- tables ----------

function parseTable(el, model, ctx, depth) {
  const tblPrEl = firstOf(el, "tblPr");
  const props = styles.parseTableProps(tblPrEl, model.theme) || {};
  const grid = [];
  const gridEl = firstOf(el, "tblGrid");
  if (gridEl) {
    for (const col of childrenOf(gridEl, "gridCol")) {
      grid.push(attrInt(col, "w", 0));
    }
  }
  const rows = [];
  for (const rowEl of childrenOf(el, "tr")) {
    const rowPr = styles.parseRowProps(firstOf(rowEl, "trPr"));
    const cells = [];
    for (const cellEl of childrenOf(rowEl, "tc")) {
      const cellPr = styles.parseCellProps(firstOf(cellEl, "tcPr"), model.theme) || {};
      const blocks = parseBlockChildren(cellEl, model, ctx, depth || 0);
      cells.push({
        props: cellPr,
        blocks,
        gridSpan: cellPr.gridSpan || 1,
        vMerge: cellPr.vMerge || null,
      });
    }
    rows.push({ props: rowPr, cells });
  }
  return { type: "table", props, grid, rows, styleId: props.styleId || null };
}

// ---------- numbering resolution ----------

// Resolves the abstract level for a numId and computes the marker text for this
// occurrence, tracking counters so nested lists number like Word does.
function resolveNumbering(numId, ilvl, model, ctx) {
  const numbering = model.numbering;
  if (!numbering || !numbering.numIdToAbstract.size) return null;
  const entry = numbering.numIdToAbstract.get(numId);
  if (!entry) return null;
  let abstract = numbering.abstract.get(entry.abstractId);
  if (abstract && abstract.numStyleLink) {
    // The style link points at a number style in styles.xml; follow it by name.
    const linked = model.styles.numbering.get(abstract.numStyleLink);
    if (linked && linked.pPr && linked.pPr.numId != null) {
      const inner = numbering.numIdToAbstract.get(linked.pPr.numId);
      if (inner) abstract = numbering.abstract.get(inner.abstractId) || abstract;
    }
  }
  if (!abstract) return null;
  const override = entry.overrides.get(ilvl);
  let level = abstract.levels[ilvl];
  if (!level && override && override.level) {
    level = Object.assign({}, abstract.levels[0] || {}, override.level);
  }
  if (!level) level = abstract.levels[0];
  if (!level) return null;
  const counters = ctx.listCounters;
  const key = numId + ":" + (abstract.id != null ? abstract.id : "a") + ":" + ilvl;
  if (!counters.has(key)) {
    counters.set(key, override && override.startOverride != null ? override.startOverride : (level.start || 1));
  } else {
    counters.set(key, counters.get(key) + 1);
  }
  const value = counters.get(key);
  // Deeper levels restart when a shallower counter advances.
  for (const otherKey of counters.keys()) {
    const parts = otherKey.split(":");
    if (parts[0] === String(numId) && Number(parts[2]) > ilvl) counters.delete(otherKey);
  }
  const marker = formatMarker(level, value, counters, ilvl);
  return {
    level,
    value,
    marker,
    ilvl,
    isBullet: level.numFmt === "bullet" || level.numFmt === "none",
  };
}

function formatMarker(level, value, counters, ilvl) {
  const text = String(level.lvlText || "");
  if (!text) return "";
  let out = "";
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    if (ch === "%") {
      const next = text[i + 1];
      const levelNum = parseInt(next, 10);
      if (!isNaN(levelNum)) {
        const key = findCounterKey(counters, ilvl, levelNum);
        const n = key != null ? counters.get(key) : value;
        out += formatCounter(n, levelAt(counters, levelNum));
        i += 2;
        continue;
      }
    }
    out += ch;
    i++;
  }
  return out;
}

function findCounterKey(counters, ilvl, wanted) {
  for (const key of counters.keys()) {
    const parts = key.split(":");
    if (Number(parts[2]) === wanted) return key;
  }
  return null;
}

function levelAt() {
  return null;
}

function formatCounter(value, level) {
  const fmt = level && level.numFmt ? level.numFmt : "decimal";
  switch (fmt) {
    case "lowerLetter": return toLetter(value).toLowerCase();
    case "upperLetter": return toLetter(value).toUpperCase();
    case "lowerRoman": return toRoman(value).toLowerCase();
    case "upperRoman": return toRoman(value).toUpperCase();
    case "ordinal": return String(value);
    case "cardinalText":
    case "ordinalText":
    case "decimalZero": return value < 10 ? "0" + value : String(value);
    case "bullet":
    case "none": return "";
    default: return String(value);
  }
}

function toLetter(n) {
  let out = "";
  let v = Math.max(1, n);
  while (v > 0) {
    const rem = (v - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    v = Math.floor((v - 1) / 26);
  }
  return out;
}

function toRoman(n) {
  const table = [
    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"],
    [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
  ];
  let out = "";
  let v = Math.max(1, Math.min(3999, n));
  for (const [value, symbol] of table) {
    while (v >= value) {
      out += symbol;
      v -= value;
    }
  }
  return out;
}

// ---------- bookmarks ----------

function collectBookmarks(el, out) {
  for (const child of el.children || []) {
    if (tagName(child) === "bookmarkStart") {
      const name = attr(child, "name");
      if (name) out.set(name, true);
    }
    collectBookmarks(child, out);
  }
}

// ---------- sections and properties ----------

function parseFinalSection(body, model) {
  const sectPr = lastOf(body, "sectPr");
  if (sectPr) return styles.parseSectionProps(sectPr, model.theme);
  return {
    pageWidthTw: 11906,
    pageHeightTw: 16838,
    marginTopTw: 1440,
    marginRightTw: 1440,
    marginBottomTw: 1440,
    marginLeftTw: 1440,
    headerTw: 708,
    footerTw: 708,
    columns: 1,
    columnSpaceTw: 708,
  };
}

function countSections(body) {
  let count = 1;
  const walk = (el) => {
    for (const child of el.children || []) {
      const tag = tagName(child);
      if (tag === "sectPr") count++;
      else if (tag === "p") {
        const pPr = firstOf(child, "pPr");
        if (pPr && firstOf(pPr, "sectPr")) count++;
      } else if (tag === "tbl") {
        walk(child);
      }
    }
  };
  walk(body);
  return count;
}

function parseCoreProperties(coreDoc, appDoc) {
  const out = { title: null, author: null, created: null, modified: null, application: null, pages: null, words: null, characters: null, paragraphs: null, fonts: [] };
  if (coreDoc) {
    const root = coreDoc.documentElement;
    for (const child of root.children || []) {
      const tag = tagName(child);
      if (tag === "title") out.title = child.textContent;
      else if (tag === "creator") out.author = child.textContent;
      else if (tag === "created") out.created = child.textContent;
      else if (tag === "modified") out.modified = child.textContent;
    }
  }
  if (appDoc) {
    const root = appDoc.documentElement;
    for (const child of root.children || []) {
      const tag = tagName(child);
      if (tag === "Application") out.application = child.textContent;
      else if (tag === "Pages") out.pages = child.textContent;
      else if (tag === "Words") out.words = child.textContent;
      else if (tag === "Characters") out.characters = child.textContent;
      else if (tag === "Paragraphs") out.paragraphs = child.textContent;
    }
  }
  return out;
}

module.exports = {
  parseDocx,
  parseParagraph,
  parseRun,
  parseTable,
  parseBlockChildren,
  resolveNumbering,
  formatMarker,
};
