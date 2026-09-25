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
 * this program. If not, see <http://www.gnu.org/licenses/>.
 */

"use strict";

// OMML, the Word equation markup, reduced to a small tree the renderer can
// build a fraction, a script, a radical or a matrix from.
//
// Word stores an equation as m:oMath, whose children are runs (m:r) and
// structures: fractions (m:f with m:num and m:den), superscripts and
// subscripts (m:sSup, m:sSub, m:sSubSup), radicals (m:rad), delimiters (m:d),
// n-ary operators (m:nary), functions, accents, bars, boxes, grouped
// characters, matrices (m:m) and equation arrays (m:eqArr). Anything not
// understood becomes its own text, so an equation never loses words.

function localName(node) {
  const name = node && (node.localName || node.tagName || node.nodeName) || "";
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

function childrenOf(node, name) {
  const out = [];
  for (const child of (node && node.children) || []) {
    if (name === undefined || localName(child) === name) out.push(child);
  }
  return out;
}

function firstOf(node, name) {
  return childrenOf(node, name)[0] || null;
}

function attrOf(node, name) {
  if (!node || !node.getAttribute) return null;
  return node.getAttribute(name) || node.getAttribute("m:" + name) || node.getAttribute("w:" + name);
}

// m:r carries m:t text and m:rPr styling. m:nor means the run is not italic,
// which is how Word marks operators and units inside an equation.
function parseMathRun(node) {
  const rPr = firstOf(node, "rPr");
  const normal = Boolean(rPr && firstOf(rPr, "nor"));
  let text = "";
  for (const child of childrenOf(node, "t")) text += child.textContent || "";
  if (!text) return null;
  return { kind: "text", text, normal };
}

function parseMathRow(node) {
  const children = [];
  for (const child of (node && node.children) || []) {
    const parsed = parseMathNode(child);
    if (parsed) children.push(parsed);
  }
  if (children.length === 1) return children[0];
  return { kind: "row", children };
}

function parseMathNode(node) {
  if (!node) return null;
  const name = localName(node);
  switch (name) {
    case "oMath":
    case "oMathPara":
    case "e":
    case "num":
    case "den":
    case "sub":
    case "sup":
    case "deg":
    case "fName":
    case "lim":
    case "box":
    case "groupChr":
    case "phant":
      return parseMathRow(node);
    case "r":
      return parseMathRun(node);
    case "f": {
      return {
        kind: "fraction",
        numerator: parseMathRow(firstOf(node, "num")),
        denominator: parseMathRow(firstOf(node, "den")),
      };
    }
    case "sSup":
      return { kind: "sup", base: parseMathRow(firstOf(node, "e")), sup: parseMathRow(firstOf(node, "sup")) };
    case "sSub":
      return { kind: "sub", base: parseMathRow(firstOf(node, "e")), sub: parseMathRow(firstOf(node, "sub")) };
    case "sSubSup":
      return {
        kind: "subsup",
        base: parseMathRow(firstOf(node, "e")),
        sub: parseMathRow(firstOf(node, "sub")),
        sup: parseMathRow(firstOf(node, "sup")),
      };
    case "rad": {
      const pr = firstOf(node, "radPr");
      const hideDegree = Boolean(pr && firstOf(pr, "degHide"));
      return {
        kind: "radical",
        degree: hideDegree ? null : parseMathRow(firstOf(node, "deg")),
        base: parseMathRow(firstOf(node, "e")),
        pr,
      };
    }
    case "d": {
      const pr = firstOf(node, "dPr");
      const begin = (pr && attrOf(firstOf(pr, "begChr"), "val")) || "(";
      const end = (pr && attrOf(firstOf(pr, "endChr"), "val")) || ")";
      const separator = (pr && attrOf(firstOf(pr, "sepChr"), "val")) || ",";
      const items = childrenOf(node, "e").map(parseMathRow).filter(Boolean);
      return { kind: "delimiter", begin, end, separator, items };
    }
    case "nary": {
      const pr = firstOf(node, "naryPr");
      const chr = (pr && attrOf(firstOf(pr, "chr"), "val")) || "\u2211";
      const subHide = Boolean(pr && firstOf(pr, "subHide"));
      const supHide = Boolean(pr && firstOf(pr, "supHide"));
      return {
        kind: "nary",
        operator: chr,
        sub: subHide ? null : parseMathRow(firstOf(node, "sub")),
        sup: supHide ? null : parseMathRow(firstOf(node, "sup")),
        base: parseMathRow(firstOf(node, "e")),
      };
    }
    case "func":
      return { kind: "function", name: parseMathRow(firstOf(node, "fName")), base: parseMathRow(firstOf(node, "e")) };
    case "limLow":
      return { kind: "limLow", base: parseMathRow(firstOf(node, "e")), limit: parseMathRow(firstOf(node, "lim")) };
    case "limUpp":
      return { kind: "limUpp", base: parseMathRow(firstOf(node, "e")), limit: parseMathRow(firstOf(node, "lim")) };
    case "acc": {
      const pr = firstOf(node, "accPr");
      const chr = (pr && attrOf(firstOf(pr, "chr"), "val")) || "\u0302";
      return { kind: "accent", accent: chr, base: parseMathRow(firstOf(node, "e")) };
    }
    case "bar": {
      const pr = firstOf(node, "barPr");
      const position = (pr && attrOf(firstOf(pr, "pos"), "val")) || "top";
      return { kind: "bar", position, base: parseMathRow(firstOf(node, "e")) };
    }
    case "m": {
      const rows = [];
      for (const mr of childrenOf(node, "mr")) {
        const cells = childrenOf(mr, "e").map(parseMathRow).filter(Boolean);
        if (cells.length) rows.push(cells);
      }
      return { kind: "matrix", rows };
    }
    case "eqArr": {
      const rows = childrenOf(node, "e").map(parseMathRow).filter(Boolean);
      return { kind: "eqArr", rows };
    }
    case "t":
      // A stray m:t outside a run still carries words.
      return node.textContent ? { kind: "text", text: node.textContent, normal: false } : null;
    default:
      return null;
  }
}

function parseMath(node) {
  const parsed = parseMathNode(node);
  if (!parsed) return null;
  if (parsed.kind === "row") return parsed;
  return { kind: "row", children: [parsed] };
}

// Plain text for search, word counts and the text fidelity check, so an
// equation is never invisible to anything but the eye.
function mathText(node) {
  if (!node) return "";
  switch (node.kind) {
    case "text":
      return node.text;
    case "row":
      return (node.children || []).map(mathText).join("");
    case "fraction":
      return mathText(node.numerator) + "/" + mathText(node.denominator);
    case "sup":
      return mathText(node.base) + "^" + mathText(node.sup);
    case "sub":
      return mathText(node.base) + "_" + mathText(node.sub);
    case "subsup":
      return mathText(node.base) + "_" + mathText(node.sub) + "^" + mathText(node.sup);
    case "radical":
      return (node.degree ? mathText(node.degree) + "\u221a" : "\u221a") + "(" + mathText(node.base) + ")";
    case "delimiter":
      return node.begin + (node.items || []).map(mathText).join(node.separator) + node.end;
    case "nary":
      return node.operator + (node.sub ? "_" + mathText(node.sub) : "") + (node.sup ? "^" + mathText(node.sup) : "") + mathText(node.base);
    case "function":
      return mathText(node.name) + mathText(node.base);
    case "limLow":
      return mathText(node.base) + "_{" + mathText(node.limit) + "}";
    case "limUpp":
      return mathText(node.base) + "^{" + mathText(node.limit) + "}";
    case "accent":
      return mathText(node.base) + node.accent;
    case "bar":
      return mathText(node.base);
    case "matrix":
      return (node.rows || []).map((row) => row.map(mathText).join(", ")).join("; ");
    case "eqArr":
      return (node.rows || []).map(mathText).join("; ");
    default:
      return "";
  }
}

module.exports = { parseMath, mathText, parseMathNode };
