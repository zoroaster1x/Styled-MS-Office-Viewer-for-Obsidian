#!/usr/bin/env python3
"""Visual 1:1 comparison of the viewer's render against the reference suite.

Both sides are rastered to same scale PNGs: the viewer's HTML output is
photographed by the headless Chrome shell (test/shot.mjs), the reference is a
headless office suite conversion read back with PyMuPDF. The preferred
reference is the OnlyOffice converter (the same engine that draws the Desktop
Editors, found through its flatpak install); LibreOffice is the fallback and
can be forced with --engine. Each page is compared as a whole and in halves
and quadrants, and every region is written out as a strip of three images:
ours, the reference, and a heatmap of the difference. The point is to see, at
a glance and 1:1, which region of which page the viewer draws differently.

  python3 test/visual-diff.py FILE... [--out DIR] [--slides N|1,2,3] [--pages N]
                                      [--sheet N] [--folder DIR --sample N]
                                      [--engine auto|onlyoffice|libreoffice]
                                      [--keep] [--quiet]

With --folder the file list is sampled per extension (N per extension,
deterministically, spread across the folder) instead of taken from FILE...

The report ranks every compared page by its bad pixel fraction: the share of
pixels whose channel distance from the reference exceeds 24 of 255. Layout is
not identical between two renderers, so a nonzero number is expected; the
ranking exists to find the pages whose difference is structural (missing
pictures, wrong colours, missing blocks) rather than a line or two of drift.
"""

import argparse
import glob as globmod
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from argparse import Namespace
from pathlib import Path

import numpy as np
import pymupdf
from PIL import Image

REPO = Path(__file__).resolve().parent.parent
BAD_THRESHOLD = 24  # channel distance out of 255 that counts as "different"
SOFT_THRESHOLD = 10  # drawn faintly on the heatmap so near misses stay visible
DEFAULT_SELECTOR = {
    "presentation": ".ov-pptx-page",
    "document": ".ov-docx-page",
    "spreadsheet": ".xlsx-sheet",
}
EXT_KIND = {
    "pptx": "presentation", "pptm": "presentation", "ppsx": "presentation",
    "ppsm": "presentation", "potx": "presentation", "potm": "presentation",
    "odp": "presentation",
    "docx": "document", "docm": "document", "dotx": "document",
    "dotm": "document", "odt": "document", "rtf": "document", "doc": "document",
    "xlsx": "spreadsheet", "xlsm": "spreadsheet", "xltx": "spreadsheet",
    "xltm": "spreadsheet", "ods": "spreadsheet", "csv": "spreadsheet",
    "tsv": "spreadsheet",
}


def run(cmd, timeout=300, env=None):
    merged = dict(os.environ)
    if env:
        merged.update(env)
    return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, check=False, env=merged)


def kind_of(path: Path) -> str:
    return EXT_KIND.get(path.suffix.lower().lstrip("."), "document")


def render_ours(path: Path, workdir: Path, kind: str, page: int, sheet: int) -> dict | None:
    """One page (or the whole document) through the viewer's own renderer."""
    html = workdir / "ours.html"
    meta = workdir / "boxes.json"
    png = workdir / "ours-raw.png"
    cmd = ["bun", str(REPO / "test" / "render-html.mjs"), str(path), "--out", str(html)]
    if kind == "presentation":
        cmd += ["--slide", str(page), "--bare"]
    elif kind == "document":
        cmd += ["--bare"]
    else:
        cmd += ["--sheet", str(sheet)]
    result = run(cmd, timeout=180)
    if result.returncode != 0:
        print("  viewer render failed:", (result.stderr or result.stdout).strip()[-300:])
        return None
    if not html.exists():
        return None
    selector = DEFAULT_SELECTOR[kind]
    shot_cmd = ["bun", str(REPO / "test" / "shot.mjs"), str(html), str(png),
                "--selector", selector, "--meta", str(meta)]
    if kind == "presentation":
        # The slide box is known from the page element itself.
        shot_cmd += ["--width", "1600", "--height", "1000"]
    result = run(shot_cmd, timeout=120)
    if result.returncode != 0 or not png.exists():
        print("  screenshot failed:", (result.stderr or result.stdout).strip()[-300:])
        return None
    boxes = []
    clip = {}
    if meta.exists():
        raw = json.loads(meta.read_text())
        boxes = raw["boxes"]
        clip = raw.get("clip", {})
    return {"png": png, "boxes": boxes, "clip": clip}


def cut_document_pages(rendered: dict, out_dir: Path) -> list[Path]:
    """Split the tall document screenshot into one image per page box."""
    if not rendered["boxes"]:
        return [rendered["png"]]
    full = Image.open(rendered["png"]).convert("RGB")
    # The screenshot was already clipped to the union of the page boxes, so a
    # box's coordinates are relative to the capture only after the clip origin
    # comes off.
    origin_x = rendered.get("clip", {}).get("x", 0)
    origin_y = rendered.get("clip", {}).get("y", 0)
    pages = []
    for index, box in enumerate(rendered["boxes"]):
        x0 = max(0, round(box["x"] - origin_x))
        y0 = max(0, round(box["y"] - origin_y))
        x1 = min(full.width, round(box["x"] + box["w"] - origin_x))
        y1 = min(full.height, round(box["y"] + box["h"] - origin_y))
        if x1 <= x0 or y1 <= y0:
            continue
        page = full.crop((x0, y0, x1, y1))
        target = out_dir / f"ours-p{index + 1}.png"
        page.save(target)
        pages.append(target)
    return pages or [rendered["png"]]


PDF_FORMAT_TO = 513  # the converter's cross platform PDF format id


def find_onlyoffice():
    """Locate the OnlyOffice command line converter and its font cache.

    Both come from the installed desktop editors: the converter binary inside
    the flatpak's own files, the generated font cache in the app's data
    directory. Environment variables override the search.
    """
    x2t = os.environ.get("OV_X2T")
    if not x2t:
        hits = globmod.glob(
            "/var/lib/flatpak/app/org.onlyoffice.desktopeditors/"
            "*/*/*/files/bin/opt/onlyoffice/desktopeditors/converter/x2t"
        )
        hits += globmod.glob(str(
            Path.home() / ".local/share/flatpak/app/org.onlyoffice.desktopeditors/"
            "*/*/*/files/bin/opt/onlyoffice/desktopeditors/converter/x2t"
        ))
        x2t = hits[0] if hits else None
    if not x2t or not Path(x2t).exists():
        return None
    all_fonts = os.environ.get("OV_X2T_FONTS")
    if not all_fonts:
        hits = globmod.glob(
            str(Path.home() / ".var/app/org.onlyoffice.desktopeditors/data/"
                "onlyoffice/desktopeditors/data/fonts/AllFonts.js")
        )
        all_fonts = hits[0] if hits else None
    if not all_fonts or not Path(all_fonts).exists():
        return None
    return Path(x2t), Path(all_fonts)


def to_pdf_onlyoffice(source: Path, workdir: Path) -> Path | None:
    """Convert with the OnlyOffice converter, one process per file, no GUI."""
    found = find_onlyoffice()
    if not found:
        print("  onlyoffice converter not found (set OV_X2T and OV_X2T_FONTS)")
        return None
    x2t, all_fonts = found
    out = workdir / "ref"
    tmp = workdir / "x2t-tmp"
    out.mkdir(exist_ok=True)
    tmp.mkdir(exist_ok=True)
    dest = out / (source.stem + ".pdf")
    params = workdir / "x2t-params.xml"
    params.write_text(
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
        f"  <m_sFileFrom>{source}</m_sFileFrom>\n"
        f"  <m_sFileTo>{dest}</m_sFileTo>\n"
        f"  <m_nFormatTo>{PDF_FORMAT_TO}</m_nFormatTo>\n"
        "  <m_sFontDir>/usr/share/fonts</m_sFontDir>\n"
        f"  <m_sAllFontsPath>{all_fonts}</m_sAllFontsPath>\n"
        f"  <m_sTempDir>{tmp}</m_sTempDir>\n"
        "</TaskQueueDataConvert>\n"
    )
    env = {
        "TMPDIR": str(tmp),
        "TEMP": str(tmp),
        "LD_LIBRARY_PATH": str(x2t.parent),
    }
    try:
        result = run([str(x2t), str(params)], timeout=600, env=env)
    except subprocess.TimeoutExpired:
        print("  onlyoffice conversion timed out")
        return None
    if result.returncode != 0 or not dest.exists():
        print("  onlyoffice conversion failed:",
              (result.stderr or result.stdout).strip()[-300:])
        return None
    return dest


def to_pdf_libreoffice(source: Path, workdir: Path) -> Path | None:
    profile = workdir / "lo-profile"
    out = workdir / "ref"
    out.mkdir(exist_ok=True)
    cmd = [
        "soffice", "--headless", "--nodefault", "--nofirststartwizard",
        "--norestore", "--nologo",
        f"-env:UserInstallation=file://{profile}",
        "--convert-to", "pdf", "--outdir", str(out), str(source),
    ]
    try:
        result = run(cmd, timeout=300)
    except subprocess.TimeoutExpired:
        print("  reference conversion timed out")
        return None
    pdf = out / (source.stem + ".pdf")
    if not pdf.exists():
        print("  reference conversion failed:",
              (result.stderr or result.stdout).strip()[-300:])
        return None
    return pdf


def to_pdf(source: Path, workdir: Path, engine: str):
    """Headless conversion with the chosen engine. Returns (pdf, engine)."""
    if engine in ("auto", "onlyoffice"):
        pdf = to_pdf_onlyoffice(source, workdir)
        if pdf:
            return pdf, "onlyoffice"
        if engine == "onlyoffice":
            return None, "onlyoffice"
    if engine in ("auto", "libreoffice"):
        pdf = to_pdf_libreoffice(source, workdir)
        if pdf:
            return pdf, "libreoffice"
    return None, engine


def raster_ref(pdf: Path, page: int, zoom: float, out_png: Path) -> Image.Image | None:
    doc = pymupdf.open(pdf)
    if page - 1 >= doc.page_count:
        doc.close()
        return None
    pdf_page = doc[page - 1]
    pix = pdf_page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), colorspace=pymupdf.csRGB, alpha=False)
    pix.save(out_png)
    doc.close()
    return Image.open(out_png).convert("RGB")


def pad_to(a: Image.Image, b: Image.Image):
    width = max(a.width, b.width)
    height = max(a.height, b.height)
    def pad(img):
        if img.size == (width, height):
            return img
        canvas = Image.new("RGB", (width, height), (255, 255, 255))
        canvas.paste(img, (0, 0))
        return canvas
    return pad(a), pad(b)


def region_boxes(width: int, height: int):
    return {
        "whole": (0, 0, width, height),
        "top-half": (0, 0, width, height // 2),
        "bottom-half": (0, height // 2, width, height),
        "left-half": (0, 0, width // 2, height),
        "right-half": (width // 2, 0, width, height),
        "quadrant-TL": (0, 0, width // 2, height // 2),
        "quadrant-TR": (width // 2, 0, width, height // 2),
        "quadrant-BL": (0, height // 2, width // 2, height),
        "quadrant-BR": (width // 2, height // 2, width, height),
    }


def compare_pair(ours: Image.Image, ref: Image.Image, out_prefix: Path, name: str) -> dict:
    ours, ref = pad_to(ours, ref)
    a = np.asarray(ours, dtype=np.int16)
    b = np.asarray(ref, dtype=np.int16)
    dist = np.abs(a - b).max(axis=2)
    bad = dist > BAD_THRESHOLD
    soft = dist > SOFT_THRESHOLD

    heat = np.full((*bad.shape, 3), 255, dtype=np.uint8)
    heat[soft] = (255, 210, 120)
    heat[bad] = (230, 40, 40)
    heat_img = Image.fromarray(heat)

    ours.save(out_prefix.with_name(out_prefix.name + ".ours.png"))
    ref.save(out_prefix.with_name(out_prefix.name + ".ref.png"))
    heat_img.save(out_prefix.with_name(out_prefix.name + ".heat.png"))

    height, width = bad.shape
    stats = {"page": name, "size": [width, height]}
    for label, (x0, y0, x1, y1) in region_boxes(width, height).items():
        region_bad = bad[y0:y1, x0:x1]
        region_dist = dist[y0:y1, x0:x1]
        entry = {
            "badFrac": round(float(region_bad.mean()), 4),
            "meanDist": round(float(region_dist.mean()), 2),
        }
        stats[label] = entry
        # The comparison strip: ours, the reference, the difference. Halves and
        # quadrants cut at the same lines for every page, so two strips can be
        # eyeballed 1:1 against each other.
        strip = Image.new("RGB", (3 * (x1 - x0) + 8, y1 - y0), (60, 60, 60))
        strip.paste(ours.crop((x0, y0, x1, y1)), (0, 0))
        strip.paste(ref.crop((x0, y0, x1, y1)), ((x1 - x0) + 4, 0))
        strip.paste(heat_img.crop((x0, y0, x1, y1)), (2 * (x1 - x0) + 8, 0))
        strip.save(out_prefix.with_name(f"{out_prefix.name}.{label}.png"))

    return stats


def slide_indices(count: int, want) -> list[int]:
    if want:
        return sorted({s for s in want if 1 <= s <= count})
    if count <= 3:
        return list(range(1, count + 1))
    # First, middle and last: three probes that usually catch a deck's range.
    return sorted({1, count // 2 + 1, count})


def ref_page_count(pdf: Path) -> int:
    doc = pymupdf.open(pdf)
    count = doc.page_count
    doc.close()
    return count


def compare_file(path: Path, out_dir: Path, options) -> list[dict]:
    kind = kind_of(path)
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", path.stem)[:60]
    file_dir = out_dir / safe
    file_dir.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix="ovvd-", dir=str(out_dir)))

    results = []
    try:
        pdf, engine = to_pdf(path, work, options.engine)
        if not pdf:
            return [{"file": str(path.name), "status": "no-pdf", "engine": options.engine}]
        if not options.quiet:
            print(f"  engine: {engine}")
        ref_count = ref_page_count(pdf)

        if kind == "presentation":
            probe = run(["bun", str(REPO / "test" / "inventory.mjs"), str(path)], timeout=120)
            slide_count = ref_count
            if probe.returncode == 0:
                try:
                    slide_count = min(ref_count, json.loads(probe.stdout).get("pages") and len(json.loads(probe.stdout)["pages"]) or ref_count)
                except Exception:
                    pass
            pages = slide_indices(slide_count, options.slides)
        elif kind == "document":
            pages = list(range(1, min(options.pages, ref_count) + 1))
        else:
            pages = [1]

        # The document's pages come out of one screenshot; presentations and
        # sheets are photographed one page at a time.
        pre_rendered = None
        if kind == "document":
            pre_rendered = render_ours(path, work, kind, 1, 1)
            if not pre_rendered:
                return [{"file": str(path.name), "status": "viewer-failed"}]
            ours_pages = cut_document_pages(pre_rendered, work)
        for page in pages:
            page_dir = file_dir / f"p{page}"
            page_dir.mkdir(parents=True, exist_ok=True)
            if kind == "document":
                if page - 1 >= len(ours_pages):
                    continue
                ours_img = Image.open(ours_pages[page - 1]).convert("RGB")
            else:
                rendered = render_ours(path, work, kind, page, options.sheet)
                if not rendered:
                    return [{"file": str(path.name), "status": "viewer-failed"}]
                ours_img = Image.open(rendered["png"]).convert("RGB")

            ref_png = work / f"ref-p{page}.png"
            # Same pixel scale on both sides: the reference PDF is rastered at
            # the zoom that maps its point based page onto the viewer's pixel
            # box. Spreadsheets have no shared page geometry, so they fall back
            # to 96 dpi and expect a size note.
            doc = pymupdf.open(pdf)
            pdf_width = doc[page - 1].rect.width
            doc.close()
            zoom = ours_img.width / pdf_width if pdf_width else 4 / 3
            ref_img = raster_ref(pdf, page, zoom, ref_png)
            if ref_img is None:
                results.append({"file": str(path.name), "page": page, "status": "no-ref-page"})
                continue
            stats = compare_pair(ours_img, ref_img, page_dir / "cmp", f"{path.name} page {page}")
            stats["file"] = str(path.name)
            stats["page"] = page
            stats["status"] = "ok"
            stats["engine"] = engine
            if ours_img.size != ref_img.size:
                stats["note"] = f"size {ours_img.size} vs {ref_img.size}"
            results.append(stats)
            if not options.quiet:
                print(f"  page {page}: bad {stats['whole']['badFrac'] * 100:.1f}%"
                      f" mean {stats['whole']['meanDist']}"
                      + (f" ({stats['note']})" if stats.get("note") else ""))
    finally:
        if not options.keep:
            shutil.rmtree(work, ignore_errors=True)
    return results


def sample_folder(folder: Path, per_ext: int) -> list[Path]:
    files = sorted(p for p in folder.rglob("*") if p.is_file() and p.suffix.lower().lstrip(".") in EXT_KIND)
    by_ext: dict[str, list[Path]] = {}
    for path in files:
        by_ext.setdefault(path.suffix.lower().lstrip("."), []).append(path)
    picked: list[Path] = []
    for ext, group in sorted(by_ext.items()):
        if per_ext <= 0 or len(group) <= per_ext:
            picked.extend(group)
            continue
        # Spread the sample over the list rather than taking the head, so one
        # week's burst of similar files does not eat the whole budget.
        step = len(group) / per_ext
        picked.extend(sorted({group[min(len(group) - 1, round(i * step))] for i in range(per_ext)}))
    return picked


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("files", nargs="*")
    parser.add_argument("--out", default="/tmp/ov-visual-diff")
    parser.add_argument("--slides", default="")
    parser.add_argument("--pages", type=int, default=2)
    parser.add_argument("--sheet", type=int, default=1)
    parser.add_argument("--folder", default="")
    parser.add_argument("--sample", type=int, default=0)
    parser.add_argument("--keep", action="store_true")
    parser.add_argument("--quiet", action="store_true")
    parser.add_argument(
        "--engine", choices=["auto", "onlyoffice", "libreoffice"], default="auto",
        help="reference converter: auto prefers the OnlyOffice converter and "
             "falls back to LibreOffice",
    )
    options = parser.parse_args()

    slides = None
    if options.slides:
        if "," in options.slides:
            slides = [int(s) for s in options.slides.split(",")]
        else:
            slides = list(range(1, int(options.slides) + 1))

    files = [Path(f) for f in options.files]
    if options.folder:
        files.extend(sample_folder(Path(options.folder).expanduser().resolve(), options.sample))
    if not files:
        parser.print_usage()
        sys.exit(2)

    out_dir = Path(options.out).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    all_results = []
    for path in files:
        path = path.expanduser().resolve()
        if not path.exists():
            print("missing:", path)
            continue
        if not options.quiet:
            print("==", path.name)
        all_results.extend(compare_file(path, out_dir, Namespace(**{**vars(options), "slides": slides})))

    ok = [r for r in all_results if r.get("status") == "ok"]
    ok.sort(key=lambda r: -r["whole"]["badFrac"])
    engines = sorted({r.get("engine", "?") for r in ok})
    print("\n=== ranking (worst first)")
    for r in ok[:40]:
        print(f"  {r['whole']['badFrac'] * 100:6.1f}%  p{r['page']:<3} {r['file']}")
    other = [r for r in all_results if r.get("status") != "ok"]
    for r in other:
        print(f"  ------  p{r.get('page', '?'):<3} {r['file']}: {r['status']}")
    if ok:
        fracs = sorted(r["whole"]["badFrac"] for r in ok)
        median = fracs[len(fracs) // 2]
        mean = sum(fracs) / len(fracs)
        print(f"\nengine {', '.join(engines)}: {len(ok)} pages, "
              f"mean {mean * 100:.1f}%, median {median * 100:.1f}%, "
              f"worst {fracs[-1] * 100:.1f}%")
    (out_dir / "summary.json").write_text(json.dumps(all_results, indent=1))
    print(f"artefacts in {out_dir}")


if __name__ == "__main__":
    main()
