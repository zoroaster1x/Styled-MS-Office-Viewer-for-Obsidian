/*
 * Styled MS Office Viewer, an Obsidian plugin that renders office documents
 * (xlsx, docx, pptx and their relatives) with their real styling, read only.
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

// Everything a document says about itself, read out of the package.
//
// Editors fill Office files with far more than the visible content: who wrote
// it, who last touched it, the company, the application and its version, the
// editing time, the language, the custom properties a template defines, the
// names attached to comments and tracked changes, and whether the file is
// protected or has edits pending. This module collects that in sections of
// label and value rows, ready for the details panel. It reads only; nothing
// here changes a file.

const CORE_FIELDS = [
  ["title", "Title"],
  ["subject", "Subject"],
  ["creator", "Author"],
  ["lastModifiedBy", "Last modified by"],
  ["keywords", "Keywords"],
  ["description", "Description"],
  ["category", "Category"],
  ["contentStatus", "Content status"],
  ["revision", "Revision"],
  ["language", "Language"],
  ["created", "Created"],
  ["modified", "Modified"],
  ["lastPrinted", "Last printed"],
];

const APP_FIELDS = [
  ["Application", "Application"],
  ["AppVersion", "Application version"],
  ["Company", "Company"],
  ["Manager", "Manager"],
  ["Template", "Template"],
  ["DocSecurity", "Document security"],
  ["ScaleCrop", "Scale to fit"],
  ["SharedDoc", "Shared document"],
  ["HyperlinksChanged", "Hyperlinks changed"],
  ["LinksUpToDate", "Links up to date"],
  ["TotalTime", "Total editing time"],
  ["Pages", "Pages"],
  ["Words", "Words"],
  ["Characters", "Characters"],
  ["CharactersWithSpaces", "Characters with spaces"],
  ["Lines", "Lines"],
  ["Paragraphs", "Paragraphs"],
  ["Slides", "Slides"],
  ["Notes", "Notes pages"],
  ["HiddenSlides", "Hidden slides"],
  ["MMClips", "Media clips"],
  ["Worksheets", "Worksheets"],
];

function displayName(name) {
  // "dc:creator" and "dcterms:created" come out of the XML, so trim the prefix;
  // the exact element name is still useful for the unusual ones.
  return String(name || "").replace(/^[a-z]+:/, "");
}

function readColonSeparated(xmlText) {
  const out = [];
  if (!xmlText) return out;
  // Only leaf elements count: a container such as cp:coreProperties wraps other
  // tags in its body, so its value may not contain another tag. The prefix is
  // optional because app.xml puts its fields in a default namespace.
  const re = /<(?:([A-Za-z0-9]+):)?([A-Za-z0-9]+)(?:\s[^>]*)?>([^<]*)<\/(?:\1:)?\2>/g;
  let match;
  while ((match = re.exec(xmlText))) {
    const value = match[3].trim();
    if (value) out.push([match[2], value]);
  }
  return out;
}

function section(title, rows) {
  const kept = rows.filter((row) => row && row[1] !== null && row[1] !== undefined && row[1] !== "");
  return kept.length ? { title, rows: kept.map(([label, value]) => ({ label, value: String(value) })) } : null;
}

// The properties store: docProps/core.xml and docProps/app.xml.
function coreSections(pkg) {
  const out = [];
  const coreText = pkg ? pkg.text("docProps/core.xml") : null;
  const core = readColonSeparated(coreText);
  const coreMap = new Map(core.map(([name, value]) => [displayName(name), value]));
  out.push(section("Document properties", CORE_FIELDS.map(([key, label]) => [label, coreMap.get(key)])));

  const appText = pkg ? pkg.text("docProps/app.xml") : null;
  const app = readColonSeparated(appText);
  const appMap = new Map(app.map(([name, value]) => [displayName(name), value]));
  out.push(section("Application", APP_FIELDS.map(([key, label]) => [label, appMap.get(key)])));
  return out;
}

// docProps/custom.xml holds whatever the template or the user defined.
function customSection(pkg) {
  const rows = [];
  const doc = pkg ? pkg.xml("docProps/custom.xml") : null;
  if (doc && doc.documentElement) {
    const walk = (el) => {
      for (const child of el.children || []) {
        const name = child.getAttribute && child.getAttribute("name");
        if (name) {
          const value = (child.textContent || "").trim();
          if (value) rows.push([name, value]);
        } else {
          walk(child);
        }
      }
    };
    walk(doc.documentElement);
  }
  return section("Custom properties", rows);
}

// The language Office stamps on text, from the part that carries it.
function languageOf(pkg, partPath) {
  const text = pkg ? pkg.text(partPath) : null;
  if (!text) return null;
  const counts = new Map();
  const re = /(?:w:val="|lang=")([a-z]{2}(?:-[A-Za-z0-9]{2,})?)"/g;
  let match;
  while ((match = re.exec(text))) counts.set(match[1], (counts.get(match[1]) || 0) + 1);
  let best = null;
  let bestCount = 0;
  for (const [value, count] of counts) {
    if (count > bestCount) {
      bestCount = count;
      best = value;
    }
  }
  return best;
}

// Comments carry authors, initials and dates, which is often the most personal
// metadata a document holds.
function commentSection(pkg, kind) {
  const rows = [];
  const names = new Set();
  if (!pkg) return null;
  const parts = pkg.list().filter((name) => {
    if (kind === "document") return /^word\/comments(?:Extended)?\.xml$/.test(name);
    if (kind === "spreadsheet") return /^xl\/(?:comments\d*|threadedComments\/.*)\.xml$/.test(name);
    if (kind === "presentation") return /^ppt\/comments\/.*\.xml$/.test(name) || name === "ppt/commentAuthors.xml";
    return false;
  });
  let count = 0;
  for (const part of parts) {
    const text = pkg.text(part);
    if (!text) continue;
    if (/commentAuthors/.test(part)) {
      const re = /<p:cmAuthor[^>]*name="([^"]*)"[^>]*initials="([^"]*)"[^>]*/g;
      let match;
      while ((match = re.exec(text))) {
        const label = match[1] + (match[2] ? " (" + match[2] + ")" : "");
        if (!names.has(label)) {
          names.add(label);
          rows.push(["Comment author", label]);
        }
      }
      continue;
    }
    const authors = /author="([^"]*)"/g;
    let match;
    while ((match = authors.exec(text))) {
      const label = match[1];
      if (label && !names.has(label)) {
        names.add(label);
        rows.push(["Comment author", label]);
      }
    }
    const commentMatches = text.match(/<(?:w:comment|p:cm|comment)\b/g);
    if (commentMatches) count += commentMatches.length;
  }
  if (count) rows.unshift(["Comments", String(count)]);
  return section("Comments", rows);
}

function documentSections(pkg, model) {
  const out = [];
  const stylesText = pkg ? pkg.text("word/styles.xml") : null;
  const settingsText = pkg ? pkg.text("word/settings.xml") : null;
  const language = languageOf(pkg, "word/styles.xml") || languageOf(pkg, "word/document.xml");
  const rows = [];
  if (language) rows.push(["Language", language]);
  const settings = settingsText || "";
  if (/<w:documentProtection\b/.test(settings)) {
    const edit = /<w:documentProtection[^>]*w:edit="([^"]*)"/.exec(settings);
    const enforced = /<w:documentProtection[^>]*w:enforcement="([^"]*)"/.exec(settings);
    rows.push(["Protection", (edit ? edit[1] : "on") + (enforced && enforced[1] === "1" ? ", enforced" : "")]);
  }
  if (/<w:trackChanges\b/.test(settings)) rows.push(["Track changes", "on"]);
  const stylesCount = stylesText ? (stylesText.match(/<w:style\b/g) || []).length : 0;
  if (stylesCount) rows.push(["Styles", String(stylesCount)]);
  const numberingText = pkg ? pkg.text("word/numbering.xml") : null;
  if (numberingText) rows.push(["Numbering definitions", String((numberingText.match(/<w:abstractNum\b/g) || []).length)]);
  const comments = commentSection(pkg, "document");
  const revisions = new Set();
  const documentText = pkg ? pkg.text("word/document.xml") : null;
  if (documentText) {
    const re = /<w:(?:ins|del)\b[^>]*w:author="([^"]*)"/g;
    let match;
    while ((match = re.exec(documentText))) revisions.add(match[1]);
  }
  if (revisions.size) rows.push(["Tracked change authors", Array.from(revisions).join(", ")]);
  // Named pageSection so it does not shadow the section() helper.
  const pageSection = model && model.section;
  if (pageSection) {
    const pt = (tw) => (tw ? Math.round(tw / 20) + "pt" : null);
    const size = pt(pageSection.pageWidthTw) && pt(pageSection.pageHeightTw)
      ? pt(pageSection.pageWidthTw) + " x " + pt(pageSection.pageHeightTw)
      : null;
    if (size) rows.push(["Page size", size]);
    const margins = [pageSection.marginTopTw, pageSection.marginRightTw, pageSection.marginBottomTw, pageSection.marginLeftTw]
      .map(pt)
      .filter(Boolean);
    if (margins.length === 4) rows.push(["Margins", margins.join(", ")]);
    if (pageSection.orientation) rows.push(["Orientation", pageSection.orientation]);
    if (pageSection.columns) rows.push(["Columns", String(pageSection.columns)]);
  }
  if (model) {
    if (Array.isArray(model.footnotes)) rows.push(["Footnotes", String(model.footnotes.length)]);
    if (Array.isArray(model.endnotes)) rows.push(["Endnotes", String(model.endnotes.length)]);
  }
  out.push(section("Word document", rows));
  if (comments) out.push(comments);
  return out;
}

function presentationSections(pkg, model) {
  const out = [];
  const rows = [];
  const language = languageOf(pkg, "ppt/slideMasters/slideMaster1.xml");
  if (language) rows.push(["Language", language]);
  const parts = pkg ? pkg.list() : [];
  const counts = { layouts: 0, masters: 0, notes: 0, media: 0, fonts: 0, charts: 0, diagrams: 0 };
  for (const name of parts) {
    if (/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name)) counts.layouts++;
    else if (/^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(name)) counts.masters++;
    else if (/^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name)) counts.notes++;
    else if (/^ppt\/media\//.test(name)) counts.media++;
    else if (/\.fntdata$/.test(name)) counts.fonts++;
    else if (/^ppt\/charts\//.test(name)) counts.charts++;
    else if (/^ppt\/diagrams\//.test(name)) counts.diagrams++;
  }
  if (counts.layouts) rows.push(["Layouts", String(counts.layouts)]);
  if (counts.masters) rows.push(["Masters", String(counts.masters)]);
  if (counts.media) rows.push(["Media parts", String(counts.media)]);
  if (counts.fonts) rows.push(["Embedded fonts", String(counts.fonts)]);
  if (counts.charts) rows.push(["Chart parts", String(counts.charts)]);
  if (counts.diagrams) rows.push(["Diagram parts", String(counts.diagrams)]);
  if (model && model.slides) {
    let hidden = 0;
    let transitions = 0;
    let animations = 0;
    const titles = [];
    for (const slide of model.slides) {
      if (slide.hidden) hidden++;
      const text = pkg && slide.path ? pkg.text(slide.path) : null;
      if (text) {
        if (/<p:transition\b/.test(text)) transitions++;
        if (/<p:timing\b/.test(text)) animations++;
      }
      if (slide.title && titles.length < 3) titles.push(slide.title.replace(/\s+/g, " ").slice(0, 50));
    }
    if (hidden) rows.push(["Hidden slides", String(hidden)]);
    if (transitions) rows.push(["Slides with a transition", transitions + " of " + model.slides.length]);
    if (animations) rows.push(["Slides with animation", animations + " of " + model.slides.length]);
    if (titles.length) rows.push(["First slides", titles.join(" / ")]);
  }
  const theme = model && model.slides && model.slides[0] ? model.slides[0].theme : null;
  if (theme && theme.name) rows.push(["Theme", theme.name]);
  out.push(section("Presentation", rows));
  const comments = commentSection(pkg, "presentation");
  if (comments) out.push(comments);
  return out;
}

function spreadsheetSections(pkg, book) {
  const out = [];
  const rows = [];
  const language = languageOf(pkg, "xl/styles.xml");
  if (language) rows.push(["Language", language]);
  const workbookText = pkg ? pkg.text("xl/workbook.xml") : null;
  if (workbookText) {
    const codeName = /<workbookPr[^>]*codeName="([^"]*)"/.exec(workbookText);
    if (codeName) rows.push(["Workbook code name", codeName[1]]);
    const date1904 = /<workbookPr[^>]*date1904="([^"]*)"/.exec(workbookText);
    if (date1904) rows.push(["Date system", date1904[1] === "1" || date1904[1] === "true" ? "1904" : "1900"]);
    const names = (workbookText.match(/<definedName\b/g) || []).length;
    if (names) rows.push(["Defined names", String(names)]);
  }
  const parts = pkg ? pkg.list() : [];
  const counts = { tables: 0, comments: 0, charts: 0, drawings: 0, media: 0, external: 0, pivots: 0 };
  for (const name of parts) {
    if (/^xl\/tables\/.*\.xml$/.test(name)) counts.tables++;
    else if (/^xl\/(?:comments\d*|threadedComments\/)/.test(name)) counts.comments++;
    else if (/^xl\/charts\//.test(name)) counts.charts++;
    else if (/^xl\/drawings\//.test(name)) counts.drawings++;
    else if (/^xl\/media\//.test(name)) counts.media++;
    else if (/^xl\/externalLinks\//.test(name)) counts.external++;
    else if (/^xl\/pivotCache\//.test(name)) counts.pivots++;
  }
  if (counts.tables) rows.push(["Tables", String(counts.tables)]);
  if (counts.charts) rows.push(["Chart parts", String(counts.charts)]);
  if (counts.pivots) rows.push(["Pivot caches", String(counts.pivots)]);
  if (counts.media) rows.push(["Media parts", String(counts.media)]);
  if (counts.external) rows.push(["External link parts", String(counts.external)]);
  if (book && book.sheets) rows.push(["Sheets", book.sheets.map((sheet) => sheet.name).join(", ")]);
  const protection = [];
  for (const name of parts) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) continue;
    const text = pkg.text(name);
    if (text && /<sheetProtection\b/.test(text)) protection.push(name.replace(/.*\//, ""));
  }
  if (protection.length) rows.push(["Protected sheets", protection.join(", ")]);
  out.push(section("Workbook", rows));
  const comments = commentSection(pkg, "spreadsheet");
  if (comments) out.push(comments);
  return out;
}

// Statistics Office keeps in app.xml are only as fresh as the last save, so
// they are labelled as saved values where that matters.
function packageSection(pkg) {
  if (!pkg) return null;
  const names = pkg.list();
  const rows = [["Parts", String(names.length)]];
  let mediaCount = 0;
  let mediaBytes = 0;
  const sizes = [];
  for (const name of names) {
    const size = typeof pkg.partSize === "function" ? pkg.partSize(name) : 0;
    if (/\/media\//.test(name)) {
      mediaCount++;
      mediaBytes += size || 0;
    }
    sizes.push([name, size]);
  }
  if (mediaCount) {
    rows.push(["Media parts", String(mediaCount)]);
    rows.push(["Media size", formatBytes(mediaBytes)]);
    const byExt = new Map();
    for (const name of names) {
      if (!/\/media\//.test(name)) continue;
      const ext = (name.split(".").pop() || "").toLowerCase();
      byExt.set(ext, (byExt.get(ext) || 0) + 1);
    }
    rows.push(["Media types", Array.from(byExt).map(([ext, count]) => ext + " x" + count).join(", ")]);
    // Be honest about the formats that cannot be drawn, so a reader knows why
    // a picture shows a placeholder.
    const undecodable = ["wdp", "hdp", "jxr"].filter((ext) => byExt.has(ext));
    const unsupported = undecodable.map((ext) => ext + " x" + byExt.get(ext)).join(", ");
    if (unsupported) rows.push(["Not drawn", unsupported + " (JPEG XR: no decoder in Chromium)"]);
  }
  sizes.sort((a, b) => b[1] - a[1]);
  const top = sizes.slice(0, 3).filter(([, size]) => size > 0);
  if (top.length) rows.push(["Largest parts", top.map(([name, size]) => name + " (" + formatBytes(size) + ")").join(", ")]);
  return section("Package", rows);
}

function formatBytes(bytes) {
  const value = Number(bytes) || 0;
  if (value < 1024) return value + " B";
  if (value < 1024 * 1024) return (value / 1024).toFixed(1) + " KB";
  return (value / (1024 * 1024)).toFixed(1) + " MB";
}

// The whole set of sections for a parsed model. Missing parts simply produce
// no rows, so an old or minimal file still shows what it has.
function describePackage(pkg, options) {
  const opts = options || {};
  const kind = opts.kind || "";
  const out = [];
  const push = (entry) => {
    if (entry && entry.rows.length) out.push(entry);
  };
  for (const entry of coreSections(pkg)) push(entry);
  push(customSection(pkg));
  if (kind === "document") for (const entry of documentSections(pkg, opts.model)) push(entry);
  else if (kind === "presentation") for (const entry of presentationSections(pkg, opts.model)) push(entry);
  else if (kind === "spreadsheet") for (const entry of spreadsheetSections(pkg, opts.model)) push(entry);
  push(packageSection(pkg));
  return out;
}

// Plain text for the copy button: one line per field, sections as headings.
function sectionsToText(sections) {
  const lines = [];
  for (const entry of sections || []) {
    if (lines.length) lines.push("");
    lines.push(entry.title);
    for (const row of entry.rows) lines.push("  " + row.label + ": " + row.value);
  }
  return lines.join("\n");
}

module.exports = {
  describePackage,
  sectionsToText,
  formatBytes,
};
