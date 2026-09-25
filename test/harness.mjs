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

// Test harness: runs the parsers and renderers outside Obsidian.
//
// linkedom gives us a DOM but not Obsidian's element extensions, and no canvas.
// Both are shimmed here so the production code stays clean: it calls the same
// methods in the app and in the tests.

import { parseHTML } from "linkedom";

const STYLE_PROPS = new Map();

// linkedom's CSSStyleDeclaration is close enough once setting a property goes
// through setProperty; some builds only reflect that path.
function installStyleSupport() {
  const proto = globalThis.CSSStyleDeclaration && globalThis.CSSStyleDeclaration.prototype;
  if (!proto || proto.__ovPatched) return;
  proto.__ovPatched = true;
  if (!Object.getOwnPropertyDescriptor(proto, "cssText")) {
    Object.defineProperty(proto, "cssText", {
      get() {
        return this.toString ? this.toString() : "";
      },
      set(value) {
        this._cssText = value;
      },
    });
  }
}

function patchElement(ElementProto) {
  if (!ElementProto || ElementProto.__ovPatched) return;
  ElementProto.__ovPatched = true;

  ElementProto.createDiv = function (cls) {
    return createEl(this, "div", { cls });
  };
  ElementProto.createSpan = function (cls) {
    return createEl(this, "span", { cls });
  };
  ElementProto.createEl = function (tag, opts) {
    return createEl(this, tag, opts);
  };
  ElementProto.setText = function (text) {
    this.textContent = text == null ? "" : String(text);
    return this;
  };
  ElementProto.getText = function () {
    return this.textContent || "";
  };
  ElementProto.empty = function () {
    while (this.firstChild) this.removeChild(this.firstChild);
    return this;
  };
  ElementProto.addClass = function (name) {
    for (const part of String(name).split(/\s+/)) if (part) this.classList.add(part);
    return this;
  };
  ElementProto.removeClass = function (name) {
    for (const part of String(name).split(/\s+/)) if (part) this.classList.remove(part);
    return this;
  };
  ElementProto.toggleClass = function (name, on) {
    if (on === undefined) this.classList.toggle(name);
    else if (on) this.classList.add(name);
    else this.classList.remove(name);
    return this;
  };
  ElementProto.hasClass = function (name) {
    return this.classList.contains(name);
  };
  ElementProto.setCssStyles = function (styles) {
    for (const key of Object.keys(styles || {})) {
      const value = styles[key];
      if (value == null) continue;
      this.style[key] = typeof value === "number" ? String(value) : value;
    }
    return this;
  };
  ElementProto.setCssProps = function (styles) {
    return this.setCssStyles(styles);
  };
  ElementProto.scrollIntoView = ElementProto.scrollIntoView || function () {};
}

function createEl(host, tag, options) {
  const opts = typeof options === "string" ? { cls: options } : (options || {});
  const doc = host.ownerDocument || host.ownerDocument === null ? host.ownerDocument : globalThis.document;
  const el = (doc || globalThis.document).createElement(tag);
  if (opts.cls) el.className = opts.cls;
  if (opts.text != null) el.textContent = String(opts.text);
  if (opts.attr) for (const key of Object.keys(opts.attr)) el.setAttribute(key, opts.attr[key]);
  if (opts.title) el.title = opts.title;
  host.appendChild(el);
  return el;
}

// ---------- canvas stub ----------

const SANS_AVERAGE = {
  a: 0.5, b: 0.5, c: 0.45, d: 0.5, e: 0.5, f: 0.31, g: 0.5, h: 0.5, i: 0.22,
  j: 0.22, k: 0.48, l: 0.22, m: 0.78, n: 0.5, o: 0.5, p: 0.5, q: 0.5, r: 0.33,
  s: 0.42, t: 0.31, u: 0.5, v: 0.48, w: 0.72, x: 0.46, y: 0.48, z: 0.42,
  " ": 0.26, ".": 0.26, ",": 0.26, ":": 0.26, ";": 0.26, "!": 0.26, "?": 0.47,
  "-": 0.33, "(": 0.33, ")": 0.33, "'": 0.19, '"': 0.4,
  A: 0.67, B: 0.59, C: 0.63, D: 0.66, E: 0.55, F: 0.53, G: 0.66, H: 0.68,
  I: 0.28, J: 0.5, K: 0.63, L: 0.53, M: 0.81, N: 0.68, O: 0.7, P: 0.58,
  Q: 0.7, R: 0.6, S: 0.58, T: 0.56, U: 0.66, V: 0.64, W: 0.93, X: 0.61,
  Y: 0.6, Z: 0.55,
};

function measureText(text, font) {
  const sizeMatch = /(\d+(?:\.\d+)?)px/.exec(font || "");
  const size = sizeMatch ? parseFloat(sizeMatch[1]) : 12;
  let width = 0;
  for (const ch of String(text)) {
    width += SANS_AVERAGE[ch] != null ? SANS_AVERAGE[ch] : 0.55;
  }
  return width * size;
}

function createCanvasStub() {
  const ctx = {
    font: "12px sans-serif",
    fillStyle: "#000000",
    strokeStyle: "#000000",
    lineWidth: 1,
    textBaseline: "alphabetic",
    textAlign: "left",
    measureText(text) {
      return { width: measureText(text, ctx.font) };
    },
    fillText() {},
    strokeText() {},
    fillRect() {},
    clearRect() {},
    strokeRect() {},
    beginPath() {},
    closePath() {},
    moveTo() {},
    lineTo() {},
    quadraticCurveTo() {},
    bezierCurveTo() {},
    arc() {},
    ellipse() {},
    rect() {},
    fill() {},
    stroke() {},
    save() {},
    restore() {},
    translate() {},
    scale() {},
    rotate() {},
    transform() {},
    setTransform() {},
    setLineDash() {},
    clip() {},
    drawImage() {},
    putImageData() {},
    createImageData(w, h) {
      return { width: w, height: h, data: new Uint8ClampedArray(w * h * 4) };
    },
    getImageData(x, y, w, h) {
      return { width: w, height: h, data: new Uint8ClampedArray(w * h * 4) };
    },
  };
  return {
    width: 300,
    height: 150,
    style: {},
    getContext(kind) {
      return kind === "2d" ? ctx : null;
    },
    toDataURL() {
      return "data:image/png;base64,";
    },
    addEventListener() {},
    removeEventListener() {},
  };
}

function installCanvas() {
  const doc = globalThis.document;
  if (!doc || doc.__ovCanvasPatched) return;
  doc.__ovCanvasPatched = true;
  // Canvas is only ever used off-screen for measuring and encoding, so a plain
  // stub object is enough and avoids linkedom's read-only element properties.
  doc.createElement = new Proxy(doc.createElement.bind(doc), {
    apply(target, thisArg, args) {
      if (String(args[0]).toLowerCase() === "canvas") return createCanvasStub();
      return Reflect.apply(target, thisArg, args);
    },
  });
}

// ---------- environment ----------

export function setupDom(html) {
  const window = parseHTML(html || "<!doctype html><html><body></body></html>");
  globalThis.window = window;
  globalThis.document = window.document;
  globalThis.DOMParser = window.DOMParser;
  globalThis.Node = window.Node;
  globalThis.Element = window.Element;
  globalThis.HTMLElement = window.HTMLElement;
  globalThis.Image = window.Image || class Image {
    constructor() {
      this.onload = null;
      setTimeout(() => this.onload && this.onload(), 0);
    }
    set src(value) {
      this._src = value;
    }
    get src() {
      return this._src;
    }
  };
  globalThis.btoa = globalThis.btoa || ((s) => Buffer.from(s, "binary").toString("base64"));
  // The JPEG XR wasm codec builds an ImageData for every decoded picture;
  // linkedom has no such class, and the browser does.
  if (!globalThis.ImageData) {
    globalThis.ImageData = class ImageData {
      constructor(data, width, height) {
        this.data = data;
        this.width = width;
        this.height = height;
      }
    };
  }
  // Embedded presentation fonts register through the FontFace API; linkedom has
  // no font set, so a loadable stub stands in outside a browser.
  if (!globalThis.FontFace) {
    globalThis.FontFace = class FontFace {
      constructor(family, source, descriptors) {
        this.family = family;
        this.source = source;
        this.weight = (descriptors && descriptors.weight) || "normal";
        this.style = (descriptors && descriptors.style) || "normal";
      }
      load() {
        return Promise.resolve(this);
      }
    };
  }
  if (window.document && !window.document.fonts) {
    window.document.fonts = { add() {}, delete() {} };
  }
  // Events dispatched into a linkedom tree must be linkedom's own classes, or
  // its EventTarget rejects the native ones.
  if (window.Event) globalThis.Event = window.Event;
  if (window.CustomEvent) globalThis.CustomEvent = window.CustomEvent;
  if (window.MouseEvent) globalThis.MouseEvent = window.MouseEvent;
  // Blob URLs are how media is served in Obsidian; emulate them so tests cover
  // the same path the app takes rather than a fallback.
  if (typeof globalThis.URL === "function" && !globalThis.URL.createObjectURL) {
    let counter = 0;
    globalThis.URL.createObjectURL = () => "blob:test/" + (++counter);
    globalThis.URL.revokeObjectURL = () => {};
  }
  installStyleSupport();
  patchElement(window.Element && window.Element.prototype);
  patchElement(window.HTMLElement && window.HTMLElement.prototype);
  installCanvas();
  return window;
}

// linkedom has no MouseEvent. Build one from its Event class and copy the
// fields the plugin reads, which is enough for pointer handling in tests.
export function pointerEvent(type, options) {
  const opts = options || {};
  const event = new globalThis.Event(type, { bubbles: opts.bubbles !== false, cancelable: true });
  for (const key of ["button", "buttons", "clientX", "clientY", "shiftKey", "ctrlKey", "metaKey", "altKey", "key"]) {
    if (opts[key] !== undefined) event[key] = opts[key];
    else if (key === "button") event[key] = 0;
  }
  if (opts.key && !event.key) event.key = opts.key;
  return event;
}

export function createContainer() {
  const host = globalThis.document.createElement("div");
  globalThis.document.body.appendChild(host);
  return host;
}
