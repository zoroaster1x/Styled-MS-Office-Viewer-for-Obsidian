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

// Number format engine: the subset of Excel number formats that shows up in
// normal workbooks. Handles General, decimals, thousands separators, percent,
// scientific, text, colours, conditions and date/time codes.

const BUILTIN_FORMATS = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  37: "#,##0;(#,##0)",
  38: "#,##0;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
};

const MONTHS_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const MONTHS_FULL = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];
const DAYS_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const DAYS_FULL = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

const COLOR_NAMES = {
  black: "#000000", blue: "#0000FF", cyan: "#00FFFF", green: "#008000",
  magenta: "#FF00FF", red: "#FF0000", white: "#FFFFFF", yellow: "#FFFF00",
};

function getFormatCode(numFmtId, customFmts) {
  if (customFmts && customFmts.has(numFmtId)) return customFmts.get(numFmtId);
  if (Object.prototype.hasOwnProperty.call(BUILTIN_FORMATS, numFmtId)) return BUILTIN_FORMATS[numFmtId];
  return "General";
}

// Splits a format code into its up to four sections on unescaped semicolons.
function splitSections(code) {
  const sections = [];
  let cur = "";
  let inQuotes = false;
  let inBrackets = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (inQuotes) {
      cur += ch;
      if (ch === '"') inQuotes = false;
      continue;
    }
    if (inBrackets) {
      cur += ch;
      if (ch === "]") inBrackets = false;
      continue;
    }
    if (ch === '"') { inQuotes = true; cur += ch; continue; }
    if (ch === "[") { inBrackets = true; cur += ch; continue; }
    if (ch === "\\") {
      cur += ch;
      if (i + 1 < code.length) { cur += code[i + 1]; i++; }
      continue;
    }
    if (ch === "_") {
      cur += " ";
      if (i + 1 < code.length) i++;
      continue;
    }
    if (ch === "*") {
      if (i + 1 < code.length) { cur += code[i + 1]; i++; }
      continue;
    }
    if (ch === ";") { sections.push(cur); cur = ""; continue; }
    cur += ch;
  }
  sections.push(cur);
  return sections;
}

// Scans a section into meta (colour, condition) and body.
function parseSection(section) {
  const meta = { color: null, condition: null };
  let body = "";
  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    if (ch === '"') {
      const end = findClosingQuote(section, i);
      body += section.slice(i, end + 1);
      i = end + 1;
      continue;
    }
    if (ch === "[") {
      const end = section.indexOf("]", i);
      const inner = section.slice(i + 1, end === -1 ? section.length : end);
      const lower = inner.trim().toLowerCase();
      if (COLOR_NAMES[lower]) {
        meta.color = COLOR_NAMES[lower];
      } else if (/^(h+|m+|s+)$/.test(lower)) {
        body += "[" + lower + "]";
      } else {
        const condMatch = /^(<=|>=|<>|=|<|>)(-?[\d.]+)$/.exec(inner.trim());
        if (condMatch) {
          meta.condition = { op: condMatch[1], value: Number(condMatch[2]) };
        }
        // Everything else (locale codes, currency names) is dropped.
      }
      i = end === -1 ? section.length : end + 1;
      continue;
    }
    body += ch;
    i++;
  }
  return { meta, body };
}

function findClosingQuote(text, start) {
  for (let i = start + 1; i < text.length; i++) {
    if (text[i] === '"') {
      if (text[i + 1] === '"') { i++; continue; }
      return i;
    }
  }
  return text.length - 1;
}

function conditionPasses(condition, value) {
  const v = condition.value;
  switch (condition.op) {
    case "<": return value < v;
    case "<=": return value <= v;
    case ">": return value > v;
    case ">=": return value >= v;
    case "=": return value === v;
    case "<>": return value !== v;
    default: return false;
  }
}

// Removes quoted text and escapes so date detection ignores literals.
function unquoted(body) {
  let out = "";
  let inQuotes = false;
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"') { inQuotes = !inQuotes; continue; }
    if (inQuotes) continue;
    if (ch === "\\" || ch === "_") { i++; continue; }
    if (ch === "*") { i++; continue; }
    out += ch;
  }
  return out;
}

function isDateFormatBody(body) {
  const plain = unquoted(body).toLowerCase();
  if (!plain) return false;
  if (/\[(h+|m+|s+)\]/.test(plain)) return true;
  if (/am\/pm|a\/p/.test(plain)) return true;
  return /[ymdhs]/.test(plain);
}

// Serial value to date parts. Excel's 1900 calendar contains a day that never
// existed: serial 60 is 1900-02-29. Every serial from 61 up is therefore one
// day ahead of the real calendar, so the epoch is 1899-12-31 and 61 and above
// lose that day again. The 1904 system has no such gap.
function dateParts(serial, date1904) {
  const days = Math.floor(serial);
  let frac = serial - days;
  if (frac < 0) frac = 0;
  const seconds = Math.round(frac * 86400);
  const time = {
    h: Math.floor(seconds / 3600) % 24,
    mi: Math.floor(seconds / 60) % 60,
    s: seconds % 60,
    ms: 0,
  };
  if (!date1904 && days === 60) {
    return Object.assign({ y: 1900, mo: 2, day: 29, weekday: 3 }, time);
  }
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
  const shifted = !date1904 && days >= 61 ? days - 1 : days;
  const d = new Date(epoch + shifted * 86400000);
  return {
    y: d.getUTCFullYear(),
    mo: d.getUTCMonth() + 1,
    day: d.getUTCDate(),
    weekday: d.getUTCDay(),
    h: time.h,
    mi: time.mi,
    s: time.s,
    ms: time.ms,
  };
}

// Fraction formats: an optional whole part, a numerator placeholder and a
// denominator that is either a literal ("?/8") or a width of question marks
// ("??/??" for halves, quarters, eighths and so on). Returns null for anything
// that is not a fraction format so the caller keeps the number path.
function formatFractionBody(body, value) {
  const match = /^([#0?,]+)?\s*([0#?]+)\s*\/\s*([0#?0-9]+)$/.exec(String(body).trim());
  if (!match) return null;
  const wholePlace = match[1] || null;
  const numPlace = match[2] || "?";
  const denPlace = match[3] || "?";
  const literal = /^[0-9]+$/.test(denPlace) ? Number(denPlace) : 0;
  const maxDenominator = literal || Math.pow(10, denPlace.length) - 1;
  if (!maxDenominator) return null;
  let whole = Math.floor(Math.abs(value));
  let frac = Math.abs(value) - whole;
  let numerator = 0;
  let denominator = 1;
  if (frac > 1e-9) {
    let bestError = Infinity;
    for (let d = 1; d <= maxDenominator; d++) {
      const n = Math.round(frac * d);
      if (n === 0) continue;
      const error = Math.abs(frac - n / d);
      if (error < bestError - 1e-12) {
        bestError = error;
        numerator = n;
        denominator = d;
      }
    }
    if (numerator === denominator) {
      whole += 1;
      numerator = 0;
      denominator = 1;
    }
  }
  const pad = (text, place, left) => {
    if (place.indexOf("?") !== -1) return text.padStart(place.length, " ");
    if (place.indexOf("0") !== -1) return text.padStart(place.length, "0");
    void left;
    return text;
  };
  const numeratorText = pad(String(numerator), numPlace, false);
  const denominatorText = literal ? String(literal) : String(denominator);
  if (wholePlace) {
    const wholeText = whole === 0 && wholePlace.indexOf("#") !== -1 ? "" : pad(String(whole), wholePlace, false);
    const joined = (wholeText ? wholeText + " " : "") + numeratorText + "/" + denominatorText;
    return joined;
  }
  return numeratorText + "/" + denominatorText;
}

function pad(n, len) {
  return String(Math.abs(n)).padStart(len, "0");
}

// First date letter in a piece of format, used to tell m for month from m for
// minutes. In "mm:ss" the s wins, in "d/m" the d wins.
function firstDateLetter(text) {
  const plain = unquoted(text).toLowerCase();
  for (let i = 0; i < plain.length; i++) {
    const ch = plain[i];
    if (ch === "y" || ch === "m" || ch === "d" || ch === "h" || ch === "s") return ch;
  }
  return "";
}

function formatDateBody(body, serial, date1904) {
  const parts = dateParts(serial, date1904);
  const lower = body.toLowerCase();
  const hasAmPm = /am\/pm|a\/p/.test(lower);
  let out = "";
  let i = 0;
  let lastToken = "";

  while (i < body.length) {
    const ch = body[i];
    const low = ch.toLowerCase();

    if (ch === '"') {
      const end = findClosingQuote(body, i);
      out += body.slice(i + 1, end);
      i = end + 1;
      continue;
    }
    if (ch === "\\") { out += body[i + 1] || ""; i += 2; continue; }
    if (ch === "[") {
      const end = body.indexOf("]", i);
      const inner = body.slice(i + 1, end === -1 ? body.length : end);
      if (/^h+$/.test(inner)) { out += String(Math.floor(serial * 24)); lastToken = "h"; }
      else if (/^m+$/.test(inner)) { out += String(Math.floor(serial * 1440)); lastToken = "m"; }
      else if (/^s+$/.test(inner)) { out += String(Math.floor(serial * 86400)); lastToken = "s"; }
      i = end === -1 ? body.length : end + 1;
      continue;
    }
    if (low === "a" && /^(am\/pm|a\/p)/.test(body.slice(i).toLowerCase())) {
      const token = /^(am\/pm|a\/p)/.exec(body.slice(i).toLowerCase())[1];
      const pm = parts.h >= 12;
      out += token === "am/pm" ? (pm ? "PM" : "AM") : (pm ? "P" : "A");
      i += token.length;
      continue;
    }
    if (low === "y" || low === "m" || low === "d" || low === "h" || low === "s") {
      let j = i;
      while (j < body.length && body[j].toLowerCase() === low) j++;
      const len = j - i;
      if (low === "y") {
        out += len >= 4 ? pad(parts.y, 4) : len === 2 ? pad(parts.y % 100, 2) : String(parts.y);
      } else if (low === "m") {
        const isMinute = lastToken === "h" || firstDateLetter(body.slice(j)) === "s";
        if (isMinute) {
          out += len >= 2 ? pad(parts.mi, 2) : String(parts.mi);
        } else if (len >= 5) {
          out += MONTHS_FULL[parts.mo - 1][0];
        } else if (len === 4) {
          out += MONTHS_FULL[parts.mo - 1];
        } else if (len === 3) {
          out += MONTHS_SHORT[parts.mo - 1];
        } else {
          out += len === 2 ? pad(parts.mo, 2) : String(parts.mo);
        }
      } else if (low === "d") {
        if (len >= 4) out += DAYS_FULL[parts.weekday];
        else if (len === 3) out += DAYS_SHORT[parts.weekday];
        else out += len === 2 ? pad(parts.day, 2) : String(parts.day);
      } else if (low === "h") {
        let h = parts.h;
        if (hasAmPm) {
          h = h % 12;
          if (h === 0) h = 12;
        }
        out += len >= 2 ? pad(h, 2) : String(h);
      } else if (low === "s") {
        out += len >= 2 ? pad(parts.s, 2) : String(parts.s);
        const frac = /^\.(0+)/.exec(body.slice(j));
        if (frac) {
          const digits = frac[1].length;
          const value = String(Math.round(parts.ms / Math.pow(10, 3 - digits)))
            .padStart(digits, "0").slice(0, digits);
          out += "." + value;
          j += frac[0].length;
        }
      }
      lastToken = low;
      i = j;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}

function formatGeneral(value) {
  if (!isFinite(value)) return String(value);
  if (Number.isInteger(value) && Math.abs(value) < 1e15) return String(value);
  const abs = Math.abs(value);
  if (abs !== 0 && (abs >= 1e11 || abs < 1e-4)) {
    return value.toExponential(5)
      .replace(/\.?0+e/, "e")
      .replace("e+", "E+")
      .replace("e-", "E-");
  }
  return String(Number(value.toPrecision(11)));
}

function groupThousands(intStr) {
  return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function formatNumberBody(body, value) {
  // Fractions are not implemented; fall back to General.
  if (/\d?\s*\/\s*[#0?]/.test(body)) return formatGeneral(value);

  const percentCount = (body.match(/%/g) || []).length;
  const v = value * Math.pow(100, percentCount);

  const sci = /[0#?]+(?:\.[0#?]+)?[eE][+-][0#?]+/.exec(body);
  if (sci) {
    const decMatch = /\.([0#?]+)/.exec(sci[0]);
    const dec = decMatch ? decMatch[1].length : 0;
    const expDigits = /[eE][+-]([0#?]+)/.exec(sci[0])[1].length;
    const exp = v === 0 ? 0 : Math.floor(Math.log10(Math.abs(v)));
    const mant = v === 0 ? 0 : v / Math.pow(10, exp);
    const expStr = String(Math.abs(exp)).padStart(expDigits, "0");
    return (v < 0 ? "-" : "") + mant.toFixed(dec) + "E" + (exp < 0 ? "-" : "+") + expStr;
  }

  const numMatch = /([#,0?]+)(?:\.([0#?]+))?/.exec(body);
  if (!numMatch) return stripFormatting(body).replace(/@/g, "");

  const intPlace = numMatch[1];
  const decPlace = numMatch[2] || "";
  const decimals = decPlace.length;
  const grouping = intPlace.indexOf(",") !== -1;
  const minIntDigits = (intPlace.match(/0/g) || []).length;

  let text = Math.abs(v).toFixed(decimals);
  let intPart = text.split(".")[0];
  let fracPart = decimals > 0 ? text.split(".")[1] : "";

  // A lone '#' shows nothing for zero.
  const isEmptyZero = minIntDigits === 0 && Number(text) === 0;

  if (!isEmptyZero) {
    if (minIntDigits > 0 && intPart.length < minIntDigits) intPart = intPart.padStart(minIntDigits, "0");
    if (grouping) intPart = groupThousands(intPart);
  } else {
    intPart = "";
    fracPart = decimals > 0 ? "".padEnd(decimals, " ") : "";
    text = "";
  }

  let numStr = intPart;
  if (decimals > 0 && fracPart !== undefined) {
    const f = fracPart === "" ? "".padEnd(decimals, "0") : fracPart;
    numStr += "." + f;
  }

  const before = stripFormatting(body.slice(0, numMatch.index));
  const after = stripFormatting(body.slice(numMatch.index + numMatch[0].length));
  return before + numStr + after;
}

// Keeps text and punctuation, removes placeholder characters.
function stripFormatting(piece) {
  let out = "";
  let inQuotes = false;
  for (let i = 0; i < piece.length; i++) {
    const ch = piece[i];
    if (inQuotes) {
      if (ch === '"') { inQuotes = false; continue; }
      out += ch;
      continue;
    }
    if (ch === '"') { inQuotes = true; continue; }
    if (ch === "\\") { out += piece[i + 1] || ""; i++; continue; }
    if (ch === "_") { out += " "; i++; continue; }
    if (ch === "*") { i++; continue; }
    if (ch === "0" || ch === "#" || ch === "?" || ch === "," || ch === ".") continue;
    out += ch;
  }
  return out;
}

function pickSection(parsed, value) {
  if (parsed.length === 1) {
    if (value < 0) return { section: parsed[0], useAbs: true, minus: true };
    return { section: parsed[0], useAbs: false };
  }
  if (value > 0) {
    const first = parsed[0];
    if (first.meta.condition && !conditionPasses(first.meta.condition, value)) {
      return { section: parsed[1] || first, useAbs: false };
    }
    return { section: first, useAbs: false };
  }
  if (value < 0) {
    const second = parsed[1];
    if (!second) return { section: parsed[0], useAbs: true, minus: true };
    if (second.meta.condition && !conditionPasses(second.meta.condition, value)) {
      return { section: parsed[0], useAbs: true, minus: !/^-|\(/.test(second.body) };
    }
    return { section: second, useAbs: !/[-()]/.test(stripFormatting(second.body)) };
  }
  return { section: parsed[2] || parsed[0], useAbs: false };
}

// Returns { text, color } for a cell value.
function formatValue(value, fmtCode, opts) {
  const options = opts || {};
  if (value == null) return { text: "" };
  if (typeof value === "boolean") return { text: value ? "TRUE" : "FALSE" };
  if (typeof value === "string") {
    if (!fmtCode || fmtCode === "General") return { text: value };
    const sections = splitSections(fmtCode);
    if (sections.length >= 4) {
      const fourth = parseSection(sections[3]);
      return { text: fourth.body.replace(/@/g, value), color: fourth.meta.color };
    }
    return { text: value };
  }

  const code = fmtCode || "General";
  if (code === "General") return { text: formatGeneral(value) };

  const parsed = splitSections(code).map(parseSection);
  const numValue = Number(value);
  const picked = pickSection(parsed, numValue);
  const section = picked.section;
  if (!section || section.body === "") return { text: "" };

  const input = picked.useAbs ? Math.abs(numValue) : numValue;
  let text;
  if (section.body === "General") {
    text = formatGeneral(input);
  } else if (isDateFormatBody(section.body)) {
    text = formatDateBody(section.body, input, Boolean(options.date1904));
  } else {
    const fraction = formatFractionBody(section.body, input);
    text = fraction != null ? fraction : formatNumberBody(section.body, input);
  }
  if (picked.minus && !/^\(.*\)$/.test(text)) text = "-" + text;
  return { text, color: section.meta.color };
}

module.exports = {
  BUILTIN_FORMATS,
  getFormatCode,
  splitSections,
  parseSection,
  isDateFormatBody,
  dateParts,
  formatGeneral,
  formatValue,
};
