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

// Small DOM helpers used by every renderer. They lean on the Obsidian element
// extensions when present and fall back to plain DOM calls, so the same code
// runs inside Obsidian and under the linkedom test harness.

function isObsidian(el) {
  return el && typeof el.createDiv === "function";
}

function createDiv(host, className) {
  if (isObsidian(host)) return host.createDiv(className);
  const el = host.ownerDocument.createElement("div");
  if (className) el.className = className;
  host.appendChild(el);
  return el;
}

function createSpan(host, className) {
  if (isObsidian(host)) return host.createSpan(className);
  const el = host.ownerDocument.createElement("span");
  if (className) el.className = className;
  host.appendChild(el);
  return el;
}

// Creates an element, applies a class and optional text or attribute bag.
function createEl(host, tag, options) {
  const opts = options || {};
  const el = host.ownerDocument.createElement(tag);
  if (opts.cls) el.className = opts.cls;
  if (opts.text != null) el.textContent = String(opts.text);
  if (opts.attr) {
    for (const key of Object.keys(opts.attr)) el.setAttribute(key, opts.attr[key]);
  }
  if (opts.title) el.title = opts.title;
  host.appendChild(el);
  return el;
}

function setText(el, text) {
  if (!el) return el;
  el.textContent = text == null ? "" : String(text);
  return el;
}

function empty(el) {
  if (!el) return el;
  while (el.firstChild) el.removeChild(el.firstChild);
  return el;
}

function addClass(el, name) {
  if (!el || !name) return el;
  if (isObsidian(el) && typeof el.addClass === "function") return el.addClass(name);
  for (const part of String(name).split(/\s+/)) {
    if (part && !el.classList.contains(part)) el.classList.add(part);
  }
  return el;
}

function removeClass(el, name) {
  if (!el || !name) return el;
  if (isObsidian(el) && typeof el.removeClass === "function") return el.removeClass(name);
  for (const part of String(name).split(/\s+/)) {
    if (part) el.classList.remove(part);
  }
  return el;
}

function toggleClass(el, name, on) {
  if (!el || !name) return el;
  if (isObsidian(el) && typeof el.toggleClass === "function") return el.toggleClass(name, on);
  if (on === undefined) el.classList.toggle(name);
  else if (on) el.classList.add(name);
  else el.classList.remove(name);
  return el;
}

function hasClass(el, name) {
  if (!el || !name) return false;
  if (isObsidian(el) && typeof el.hasClass === "function") return el.hasClass(name);
  return el.classList.contains(name);
}

function setStyles(el, styles) {
  if (!el || !styles) return el;
  for (const key of Object.keys(styles)) {
    const value = styles[key];
    if (value == null) continue;
    el.style[key] = typeof value === "number" ? String(value) : value;
  }
  return el;
}

module.exports = {
  createDiv,
  createSpan,
  createEl,
  setText,
  empty,
  addClass,
  removeClass,
  toggleClass,
  hasClass,
  setStyles,
};
