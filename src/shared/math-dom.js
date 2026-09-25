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

// Draws an equation tree from src/docx/math.js into DOM spans. Word and
// PowerPoint store the same OMML, so the document and the slide renderers share
// this instead of each keeping its own copy. The class names stay ov-docx-*
// because the stylesheet that shapes fractions, radicals and the rest is
// shared too.

function renderMathDom(node, parent, paraProps) {
  if (!node) return;
  switch (node.kind) {
    case "row":
      for (const child of node.children || []) renderMathDom(child, parent, paraProps);
      return;
    case "text": {
      const span = parent.createSpan("ov-docx-math");
      span.setText(node.text);
      span.addClass(node.normal ? "is-normal" : "is-variable");
      return;
    }
    case "fraction": {
      const wrap = parent.createSpan("ov-docx-math ov-docx-frac");
      const numerator = wrap.createSpan("ov-docx-frac-num");
      renderMathDom(node.numerator, numerator, paraProps);
      const denominator = wrap.createSpan("ov-docx-frac-den");
      renderMathDom(node.denominator, denominator, paraProps);
      return;
    }
    case "sup":
    case "sub":
    case "subsup": {
      renderMathDom(node.base, parent, paraProps);
      if (node.kind === "subsup") {
        const stack = parent.createSpan("ov-docx-math-stack");
        const above = stack.createSpan("ov-docx-math-script");
        renderMathDom(node.sup, above, paraProps);
        const below = stack.createSpan("ov-docx-math-script");
        renderMathDom(node.sub, below, paraProps);
      } else {
        const el = parent.createEl(node.kind === "sup" ? "sup" : "sub", { cls: "ov-docx-math-script" });
        renderMathDom(node.kind === "sup" ? node.sup : node.sub, el, paraProps);
      }
      return;
    }
    case "radical": {
      const wrap = parent.createSpan("ov-docx-math ov-docx-radical");
      if (node.degree) {
        const degree = wrap.createSpan("ov-docx-radical-deg");
        renderMathDom(node.degree, degree, paraProps);
      }
      wrap.createSpan("ov-docx-radical-sign").setText("\u221a");
      const body = wrap.createSpan("ov-docx-radical-body");
      renderMathDom(node.base, body, paraProps);
      return;
    }
    case "delimiter": {
      const wrap = parent.createSpan("ov-docx-math ov-docx-delim");
      wrap.createSpan("ov-docx-delim-char").setText(node.begin);
      (node.items || []).forEach((item, index) => {
        if (index) wrap.createSpan("ov-docx-delim-sep").setText(node.separator);
        renderMathDom(item, wrap, paraProps);
      });
      wrap.createSpan("ov-docx-delim-char").setText(node.end);
      return;
    }
    case "nary": {
      const wrap = parent.createSpan("ov-docx-math ov-docx-nary");
      wrap.createSpan("ov-docx-nary-op").setText(node.operator);
      if (node.sup || node.sub) {
        const limits = wrap.createSpan("ov-docx-nary-limits");
        if (node.sup) {
          const above = limits.createSpan("ov-docx-math-script");
          renderMathDom(node.sup, above, paraProps);
        }
        if (node.sub) {
          const below = limits.createSpan("ov-docx-math-script");
          renderMathDom(node.sub, below, paraProps);
        }
      }
      renderMathDom(node.base, wrap, paraProps);
      return;
    }
    case "function":
      renderMathDom(node.name, parent, paraProps);
      renderMathDom(node.base, parent, paraProps);
      return;
    case "limLow":
    case "limUpp": {
      const wrap = parent.createSpan("ov-docx-math ov-docx-limit");
      renderMathDom(node.base, wrap, paraProps);
      const limit = wrap.createSpan(node.kind === "limLow" ? "ov-docx-limit-below" : "ov-docx-limit-above");
      renderMathDom(node.limit, limit, paraProps);
      return;
    }
    case "accent": {
      const wrap = parent.createSpan("ov-docx-math ov-docx-accent");
      wrap.createSpan("ov-docx-accent-mark").setText(node.accent);
      renderMathDom(node.base, wrap, paraProps);
      return;
    }
    case "bar": {
      const wrap = parent.createSpan("ov-docx-math " + (node.position === "bot" ? "ov-docx-bar-bot" : "ov-docx-bar-top"));
      renderMathDom(node.base, wrap, paraProps);
      return;
    }
    case "matrix": {
      const table = parent.createEl("table", { cls: "ov-docx-matrix" });
      for (const row of node.rows || []) {
        const tr = table.createEl("tr");
        for (const cell of row) {
          const td = tr.createEl("td");
          renderMathDom(cell, td, paraProps);
        }
      }
      return;
    }
    case "eqArr": {
      const wrap = parent.createSpan("ov-docx-math ov-docx-eqarr");
      for (const row of node.rows || []) {
        const line = wrap.createSpan("ov-docx-eqarr-row");
        renderMathDom(row, line, paraProps);
      }
      return;
    }
    default:
      return;
  }
}

module.exports = { renderMathDom };
