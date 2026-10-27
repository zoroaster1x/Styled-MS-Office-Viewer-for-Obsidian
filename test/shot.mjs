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

// Screenshots a standalone HTML file with the headless Chrome shell over the
// DevTools protocol. No GUI, no extra packages: the shell speaks CDP over a
// WebSocket and this script speaks just enough of it to load a page, wait for
// its fonts and images, measure the boxes of the document's own page elements
// and capture them as PNG.
//
//   bun test/shot.mjs page.html out.png [--selector SEL] [--meta boxes.json]
//                      [--width W] [--height H] [--wait MS] [--scale S]
//
// --selector clips the capture to the union of the matched elements' boxes and
// records each box in the --meta JSON, which is how the document renderer's
// stacked pages are later cut into one image per page. The default selector is
// the union of the viewer's page elements.

import { spawn } from "node:child_process";
import { existsSync, readdirSync, writeFileSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

const args = process.argv.slice(2);
const positional = args.filter((a) => !a.startsWith("--"));
function valueOf(flag, fallback) {
  const i = args.indexOf(flag);
  return i === -1 ? fallback : args[i + 1];
}
const has = (flag) => args.includes(flag);

const htmlPath = positional[0];
const outPath = positional[1];
const selector = valueOf("--selector", ".ov-pptx-page, .ov-docx-page");
const metaPath = valueOf("--meta", null);
const width = Number(valueOf("--width", 1280));
const height = Number(valueOf("--height", 800));
const waitMs = Number(valueOf("--wait", 900));
const scale = Number(valueOf("--scale", 1));

if (!htmlPath || !outPath) {
  console.error("usage: bun test/shot.mjs page.html out.png [--selector SEL] [--meta boxes.json]");
  process.exit(2);
}

// The shell binary is found the way the rest of the test setup finds tools:
// an explicit environment variable first, then the well known puppeteer cache
// layout, then whatever is on PATH.
function findChrome() {
  if (process.env.OV_CHROME && existsSync(process.env.OV_CHROME)) return process.env.OV_CHROME;
  const home = process.env.HOME || "";
  if (home) {
    const cache = join(home, ".cache", "puppeteer", "chrome-headless-shell");
    try {
      for (const version of readdirSync(cache).sort().reverse()) {
        const candidate = join(cache, version, "chrome-headless-shell-linux64", "chrome-headless-shell");
        if (existsSync(candidate)) return candidate;
      }
    } catch {}
  }
  for (const name of ["chromium", "google-chrome", "chrome"]) {
    const dirs = String(process.env.PATH || "").split(":");
    for (const dir of dirs) {
      const candidate = join(dir, name);
      if (existsSync(candidate)) return candidate;
    }
  }
  return null;
}

const chrome = findChrome();
if (!chrome) {
  console.error("shot: no chrome found. Set OV_CHROME to a chrome or chromium binary.");
  process.exit(1);
}

const profile = mkdtempSync(join(tmpdir(), "ov-shot-"));
const child = spawn(chrome, [
  "--headless",
  "--no-sandbox",
  "--disable-gpu",
  "--hide-scrollbars",
  "--force-color-profile=srgb",
  "--allow-file-access-from-files",
  "--no-first-run",
  "--disable-background-timer-throttling",
  "--remote-debugging-port=0",
  `--user-data-dir=${profile}`,
  "about:blank",
], { stdio: ["ignore", "ignore", "pipe"] });

function shutdown(code, message) {
  try { child.kill("SIGKILL"); } catch {}
  console.error(message || "");
  process.exit(code);
}

// The shell prints the DevTools endpoint on stderr once it is listening.
const wsUrl = await new Promise((done, fail) => {
  let buffer = "";
  const timer = setTimeout(() => fail(new Error("chrome did not open a DevTools socket")), 20000);
  child.stderr.on("data", (chunk) => {
    buffer += String(chunk);
    const match = /DevTools listening on (ws:\/\/\S+)/.exec(buffer);
    if (match) {
      clearTimeout(timer);
      done(match[1]);
    }
  });
  child.on("exit", (code) => {
    clearTimeout(timer);
    fail(new Error("chrome exited with code " + code + ": " + buffer.slice(0, 400)));
  });
});

const ws = new WebSocket(wsUrl);
await new Promise((done, fail) => {
  ws.addEventListener("open", done, { once: true });
  ws.addEventListener("error", () => fail(new Error("websocket failed")), { once: true });
});

let nextId = 1;
const pending = new Map();
const events = [];
function send(method, params, sessionId) {
  const id = nextId++;
  const message = { id, method, params: params || {} };
  if (sessionId) message.sessionId = sessionId;
  ws.send(JSON.stringify(message));
  return new Promise((done, fail) => {
    pending.set(id, { done, fail, method });
    setTimeout(() => {
      if (pending.has(id)) {
        pending.delete(id);
        fail(new Error("CDP timeout: " + method));
      }
    }, 60000);
  });
}

ws.addEventListener("message", (event) => {
  const message = JSON.parse(String(event.data));
  if (message.id && pending.has(message.id)) {
    const entry = pending.get(message.id);
    pending.delete(message.id);
    if (message.error) entry.fail(new Error(entry.method + ": " + message.error.message));
    else entry.done(message.result);
    return;
  }
  if (message.method) events.push(message);
});
function once(method, sessionId) {
  return new Promise((done) => {
    const check = () => {
      const index = events.findIndex((m) => m.method === method && (!sessionId || m.sessionId === sessionId));
      if (index !== -1) {
        events.splice(index, 1);
        done(events[index] || null);
        return true;
      }
      return false;
    };
    if (check()) return;
    const timer = setInterval(() => {
      if (check()) clearInterval(timer);
    }, 25);
    setTimeout(() => clearInterval(timer), 60000);
  });
}

const { targetId } = await send("Target.createTarget", { url: "about:blank" });
const { sessionId } = await send("Target.attachToTarget", { targetId, flatten: true });
await send("Page.enable", {}, sessionId);
await send("Emulation.setDeviceMetricsOverride", {
  width, height, deviceScaleFactor: scale, mobile: false,
}, sessionId);

await send("Page.navigate", { url: "file://" + resolve(htmlPath) }, sessionId);
await once("Page.loadEventFired", sessionId);

// Wait for the document to settle: fonts first, then a fixed beat for image
// decode and layout. The page carries no scripts of its own.
const settle = `new Promise((done) => {
  const finish = () => setTimeout(done, ${waitMs});
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(finish, finish);
  else finish();
})`;
await send("Runtime.evaluate", {
  expression: settle, awaitPromise: true, returnByValue: true,
}, sessionId);

const probe = await send("Runtime.evaluate", {
  expression: `(() => {
    const nodes = Array.from(document.querySelectorAll(${JSON.stringify(selector)}));
    const boxes = nodes.map((el) => {
      const r = el.getBoundingClientRect();
      return { x: r.x, y: r.y + window.scrollY, w: r.width, h: r.height };
    });
    return JSON.stringify({
      docWidth: document.documentElement.scrollWidth,
      docHeight: document.documentElement.scrollHeight,
      boxes,
    });
  })()`,
  returnByValue: true,
}, sessionId);
const meta = JSON.parse(probe.result.value);

let clip;
if (meta.boxes.length) {
  const x0 = Math.min(...meta.boxes.map((b) => b.x));
  const y0 = Math.min(...meta.boxes.map((b) => b.y));
  const x1 = Math.max(...meta.boxes.map((b) => b.x + b.w));
  const y1 = Math.max(...meta.boxes.map((b) => b.y + b.h));
  clip = { x: x0, y: y0, width: Math.ceil(x1 - x0), height: Math.ceil(y1 - y0), scale };
} else {
  clip = { x: 0, y: 0, width: meta.docWidth, height: meta.docHeight, scale };
}

const shot = await send("Page.captureScreenshot", {
  format: "png",
  captureBeyondViewport: true,
  clip,
}, sessionId);
writeFileSync(outPath, Buffer.from(shot.data, "base64"));
if (metaPath) {
  writeFileSync(metaPath, JSON.stringify({
    clip,
    docWidth: meta.docWidth,
    docHeight: meta.docHeight,
    boxes: meta.boxes,
  }, null, 1));
}
console.log("shot", outPath, clip.width + "x" + clip.height, "boxes", meta.boxes.length);
try { child.kill("SIGKILL"); } catch {}
ws.close();
process.exit(0);
