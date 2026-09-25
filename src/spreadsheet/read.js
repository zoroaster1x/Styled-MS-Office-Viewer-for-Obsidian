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

const { unzipSync } = require("fflate");
const { Package } = require("../shared/package");
const { childrenOf, firstOf, attr, attrInt, parseXml } = require("../shared/xml");
const { parseStyles } = require("./styles");
const { parseSheet } = require("./sheet");

// Reads an .xlsx (zip) buffer into a workbook model. Sheets are parsed on
// demand and cached, so opening a big workbook only pays for the visible sheet.

function readWorkbook(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  let files;
  try {
    files = unzipSync(bytes);
  } catch (err) {
    throw new Error("Could not unzip this file. Is it a real .xlsx workbook? (" + err.message + ")");
  }
  // The metadata panel reads the package directory; opening it here costs one
  // pass over the central directory and nothing is inflated twice.
  let pkg = null;
  try {
    pkg = Package.open(bytes);
  } catch (err) {
    pkg = null;
  }

  const decoder = new TextDecoder("utf-8");
  const readText = (name) => (files[name] ? decoder.decode(files[name]) : null);
  const readRels = (name) => parseRels(readText(name));

  const workbookXml = readText("xl/workbook.xml");
  if (!workbookXml) {
    throw new Error("This file is missing xl/workbook.xml, so it is not a readable .xlsx workbook.");
  }

  const theme = parseTheme(readText("xl/theme/theme1.xml"));
  const styles = parseStyles(readText("xl/styles.xml"), theme);
  const sharedStrings = parseSharedStrings(readText("xl/sharedStrings.xml"));

  const doc = parseXml(workbookXml);
  const root = doc.documentElement;
  const workbookPr = firstOf(root, "workbookPr");
  const date1904 = workbookPr ? attr(workbookPr, "date1904") === "1" || attr(workbookPr, "date1904") === "true" : false;

  const workbookRels = readRels("xl/_rels/workbook.xml.rels");
  const sheets = [];
  const sheetList = firstOf(root, "sheets");
  if (sheetList) {
    const list = childrenOf(sheetList, "sheet");
    for (let i = 0; i < list.length; i++) {
      const el = list[i];
      const rid = el.getAttribute("r:id") || el.getAttribute("id");
      const rel = rid ? workbookRels.get(rid) : null;
      if (!rel) continue;
      sheets.push({
        name: attr(el, "name") || "Sheet" + (i + 1),
        index: i,
        state: attr(el, "state") || "visible",
        sheetId: attrInt(el, "sheetId", i + 1),
        path: resolvePath("xl", rel.target),
      });
    }
  }

  const cache = new Map();
  const api = {
    sheets,
    sharedStrings,
    styles,
    theme,
    date1904,
    pkg,
    properties: parseDocumentProperties(readText("docProps/core.xml"), readText("docProps/app.xml")),
    files: Object.keys(files),
    loadSheet(index) {
      if (cache.has(index)) return cache.get(index);
      const info = sheets[index];
      if (!info) throw new Error("Sheet index " + index + " is out of range.");
      const xml = readText(info.path);
      if (!xml) throw new Error("Sheet file " + info.path + " is missing from the workbook.");
      const relsPath = relsPathFor(info.path);
      const model = parseSheet(xml, {
        name: info.name,
        sharedStrings,
        styles,
        theme,
        rels: readRels(relsPath),
        date1904,
      });
      cache.set(index, model);
      return model;
    },
    sheetIndexByName(name) {
      for (let i = 0; i < sheets.length; i++) {
        if (sheets[i].name === name) return i;
      }
      return -1;
    },
  };
  return api;
}

// The property parts Excel writes: creator and dates in core.xml, the
// application and counts in app.xml.
function parseDocumentProperties(coreText, appText) {
  const out = { title: null, author: null, created: null, modified: null, application: null, company: null, sheets: null };
  if (coreText) {
    try {
      const doc = parseXml(coreText);
      for (const child of doc.documentElement.children || []) {
        const tag = (child.tagName || "").replace(/^.*:/, "");
        if (tag === "title") out.title = child.textContent;
        else if (tag === "creator") out.author = child.textContent;
        else if (tag === "created") out.created = child.textContent;
        else if (tag === "modified") out.modified = child.textContent;
      }
    } catch (err) {
      // A missing or malformed property part is not a read failure.
    }
  }
  if (appText) {
    try {
      const doc = parseXml(appText);
      for (const child of doc.documentElement.children || []) {
        const tag = (child.tagName || "").replace(/^.*:/, "");
        if (tag === "Application") out.application = child.textContent;
        else if (tag === "Company") out.company = child.textContent;
        else if (tag === "TitlesOfParts") out.sheets = (child.children || []).length;
      }
    } catch (err) {
      // Ignored, as above.
    }
  }
  return out;
}

function relsPathFor(partPath) {
  const slash = partPath.lastIndexOf("/");
  const dir = slash === -1 ? "" : partPath.slice(0, slash + 1);
  const name = slash === -1 ? partPath : partPath.slice(slash + 1);
  return dir + "_rels/" + name + ".rels";
}

function resolvePath(baseDir, target) {
  if (!target) return null;
  if (target.startsWith("/")) return target.slice(1);
  const parts = (baseDir ? baseDir + "/" + target : target).split("/");
  const out = [];
  for (const part of parts) {
    if (part === "" || part === ".") continue;
    if (part === "..") out.pop();
    else out.push(part);
  }
  return out.join("/");
}

function parseRels(text) {
  const map = new Map();
  if (!text) return map;
  const doc = parseXml(text);
  const list = doc.documentElement.children || [];
  for (let i = 0; i < list.length; i++) {
    const el = list[i];
    if ((el.tagName || "").indexOf("Relationship") === -1) continue;
    map.set(attr(el, "Id"), {
      type: attr(el, "Type"),
      target: attr(el, "Target"),
      mode: attr(el, "TargetMode"),
    });
  }
  return map;
}

function parseTheme(text) {
  const theme = { colors: [] };
  if (!text) return theme;
  const doc = parseXml(text);
  const scheme = findFirst(doc.documentElement, "clrScheme");
  if (!scheme) return theme;
  const byName = {};
  const list = scheme.children || [];
  for (let i = 0; i < list.length; i++) {
    const el = list[i];
    const name = (el.tagName || "").replace(/^.*:/, "");
    const srgb = firstOf(el, "srgbClr");
    const sys = firstOf(el, "sysClr");
    let hex = null;
    if (srgb) hex = attr(srgb, "val");
    else if (sys) hex = attr(sys, "lastClr") || "000000";
    if (hex) byName[name] = "#" + String(hex).replace(/^#/, "");
  }
  // Theme indexes used by colour references: lt1, dk1, lt2, dk2, accents, links.
  const order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  theme.colors = order.map((name) => byName[name] || null);
  return theme;
}

function findFirst(el, tag) {
  if (!el) return null;
  const list = el.children || [];
  for (let i = 0; i < list.length; i++) {
    const child = list[i];
    const name = (child.tagName || "").replace(/^.*:/, "");
    if (name === tag) return child;
    const found = findFirst(child, tag);
    if (found) return found;
  }
  return null;
}

function parseSharedStrings(text) {
  if (!text) return [];
  const doc = parseXml(text);
  const out = [];
  const list = doc.documentElement.children || [];
  for (let i = 0; i < list.length; i++) {
    const el = list[i];
    if ((el.tagName || "").indexOf("si") === -1) continue;
    out.push(textOf(el));
  }
  return out;
}

function textOf(el) {
  let out = "";
  const walk = (node) => {
    const list = node.children || [];
    for (let i = 0; i < list.length; i++) {
      const child = list[i];
      const name = (child.tagName || "").replace(/^.*:/, "");
      if (name === "rPh" || name === "phoneticPr") continue;
      if (name === "t") out += child.textContent || "";
      else walk(child);
    }
  };
  walk(el);
  return out;
}

module.exports = { readWorkbook, resolvePath, parseRels, parseTheme, parseSharedStrings };
