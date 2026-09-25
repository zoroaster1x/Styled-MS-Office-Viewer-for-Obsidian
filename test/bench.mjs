import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { setupDom } from "./harness.mjs";
setupDom();
const require = createRequire(import.meta.url);
const { Package } = require("../src/shared/package.js");
const { mediaUrl } = require("../src/media/media.js");
const file = process.argv[2];
const bytes = new Uint8Array(readFileSync(file));
const size = bytes.length;
let t = performance.now();
const pkg = Package.open(bytes);
const unzipMs = performance.now() - t;
const textParts = pkg.list().filter((n) => n.endsWith(".xml"));
const mediaParts = pkg.list().filter((n) => /\.(png|jpe?g|gif|tiff?|emf|wmf|svg|bmp)$/i.test(n));
let xmlTotal = 0;
for (const name of textParts) {
  t = performance.now();
  pkg.xml(name);
  xmlTotal += performance.now() - t;
}
let mediaTotal = 0;
let mediaBytes = 0;
for (const name of mediaParts.slice(0, 10)) {
  t = performance.now();
  const url = mediaUrl(pkg, name, null);
  mediaTotal += performance.now() - t;
  if (url) mediaBytes += url.length;
}
console.log("file".padEnd(14), (size / 1024 / 1024).toFixed(1) + " MB");
console.log("unzip".padEnd(14), unzipMs.toFixed(0) + " ms", pkg.list().length + " parts");
console.log("xml all".padEnd(14), xmlTotal.toFixed(0) + " ms", textParts.length + " parts");
console.log("media 10".padEnd(14), mediaTotal.toFixed(0) + " ms", Math.round(mediaBytes / 1024) + " KB as data urls");
