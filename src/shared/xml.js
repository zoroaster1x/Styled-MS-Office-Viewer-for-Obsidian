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

// XML helpers. Every part of an OOXML or ODF package is XML, so the traversal
// code here is shared by all the parsers. Namespace prefixes are dropped:
// a:blip and blip are the same element to us.

function tagName(el) {
  const t = el.tagName || el.nodeName || "";
  const i = t.indexOf(":");
  return i === -1 ? t : t.slice(i + 1);
}

function stripNs(name) {
  const s = String(name || "");
  const i = s.indexOf(":");
  return i === -1 ? s : s.slice(i + 1);
}

function parseXml(text) {
  return new DOMParser().parseFromString(text, "application/xml");
}

// Children with a given local name. Omit the tag to get every element child.
// Accepts a Document or an Element, so callers can pass either.
function elementOf(node) {
  if (!node) return null;
  if (node.nodeType === 9 && node.documentElement) return node.documentElement;
  return node;
}

function childrenOf(el, tag) {
  const out = [];
  const node = elementOf(el);
  if (!node) return out;
  const list = node.children || [];
  for (let i = 0; i < list.length; i++) {
    if (!tag || tagName(list[i]) === tag) out.push(list[i]);
  }
  return out;
}

function firstOf(el, tag) {
  const node = elementOf(el);
  if (!node) return null;
  const list = node.children || [];
  for (let i = 0; i < list.length; i++) {
    if (tagName(list[i]) === tag) return list[i];
  }
  return null;
}

function lastOf(el, tag) {
  const node = elementOf(el);
  if (!node) return null;
  const list = node.children || [];
  for (let i = list.length - 1; i >= 0; i--) {
    if (tagName(list[i]) === tag) return list[i];
  }
  return null;
}

// First descendant with a local name, depth first.
function findFirst(el, tag) {
  const node = elementOf(el);
  if (!node) return null;
  const list = node.children || [];
  for (let i = 0; i < list.length; i++) {
    const child = list[i];
    if (tagName(child) === tag) return child;
    const found = findFirst(child, tag);
    if (found) return found;
  }
  return null;
}

function findAll(el, tag, out) {
  const acc = out || [];
  const node = elementOf(el);
  if (!node) return acc;
  const list = node.children || [];
  for (let i = 0; i < list.length; i++) {
    const child = list[i];
    if (tagName(child) === tag) acc.push(child);
    findAll(child, tag, acc);
  }
  return acc;
}

function attr(el, name) {
  if (!el || !el.getAttribute) return null;
  // WordprocessingML and PresentationML qualify most attributes with the
  // document namespace, so a plain lookup has to fall back to the local name.
  let v = el.getAttribute(name);
  if (v !== null) return v;
  if (el.attributes) {
    for (let i = 0; i < el.attributes.length; i++) {
      const a = el.attributes[i];
      if (a.name === name) return a.value;
      const local = a.name.indexOf(":") === -1 ? a.name : a.name.slice(a.name.indexOf(":") + 1);
      if (local === name) return a.value;
    }
  }
  return null;
}

// Attributes in XML use a prefix; callers pass the local name.
function attrNs(el, local) {
  if (!el || !el.attributes) return null;
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    if (stripNs(a.name) === local) return a.value;
  }
  return null;
}

function attrInt(el, name, dflt) {
  const v = attr(el, name);
  if (v === null || v === "") return dflt === undefined ? 0 : dflt;
  const n = parseInt(v, 10);
  return isNaN(n) ? (dflt === undefined ? 0 : dflt) : n;
}

function attrNum(el, name, dflt) {
  const v = attr(el, name);
  if (v === null || v === "") return dflt === undefined ? 0 : dflt;
  const n = parseFloat(v);
  return isNaN(n) ? (dflt === undefined ? 0 : dflt) : n;
}

function attrBool(el, name, dflt) {
  const v = attr(el, name);
  if (v === null) return dflt === undefined ? false : dflt;
  return v === "1" || v === "true" || v === "on";
}

// OOXML writes booleans as <w:b/> or <w:b w:val="0"/>.
function flag(el, dflt) {
  if (!el) return dflt === undefined ? false : dflt;
  const v = el.getAttribute("val");
  if (v === null) return true;
  return v === "1" || v === "true" || v === "on";
}

function textContent(el) {
  return el ? el.textContent || "" : "";
}

// Text of the first child with a local name.
function childText(el, tag) {
  const child = firstOf(el, tag);
  return child ? child.textContent || "" : "";
}

module.exports = {
  tagName,
  stripNs,
  elementOf,
  parseXml,
  childrenOf,
  firstOf,
  lastOf,
  findFirst,
  findAll,
  attr,
  attrNs,
  attrInt,
  attrNum,
  attrBool,
  flag,
  textContent,
  childText,
};
