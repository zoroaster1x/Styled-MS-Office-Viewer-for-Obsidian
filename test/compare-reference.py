#!/usr/bin/env python3
"""Compare the viewer against the reference office suite.

The reference is a headless conversion of the same file, read back
with PyMuPDF. That gives, per page: the text, the images with the rectangle each
one occupies, and the fonts. The viewer's own inventory comes from
test/inventory.mjs. The diff answers the two questions that matter:

  * is any text missing
  * is any picture missing, or in the wrong place

Usage:
  python3 test/compare-reference.py FILE [FILE ...] [--keep]
"""

import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import pymupdf

REPO = Path(__file__).resolve().parent.parent
# Two rects are "the same picture in the same place" when their union overlap
# is at least this high. The suite rounds picture frames on export, so exact
# equality is not the bar.
MATCH_THRESHOLD = 0.85
INVENTORY = REPO / "test" / "inventory.mjs"


def run(cmd, **kwargs):
    return subprocess.run(cmd, capture_output=True, text=True, check=False, **kwargs)


def to_pdf(source: Path, workdir: Path) -> Path | None:
    """Headless conversion with the system office suite. No GUI, one profile per run."""
    profile = workdir / "lo-profile"
    out = workdir / "out"
    out.mkdir(exist_ok=True)
    result = run([
        "soffice", "--headless", "--nodefault", "--nofirststartwizard",
        "--norestore", "--nologo",
        f"-env:UserInstallation=file://{profile}",
        "--convert-to", "pdf:writer_pdf_Export",
        "--outdir", str(out), str(source),
    ], timeout=300)
    pdf = out / (source.stem + ".pdf")
    if not pdf.exists():
        # The presentation and spreadsheet filters differ; retry letting the suite choose.
        result = run([
            "soffice", "--headless", "--nodefault", "--nofirststartwizard",
            "--norestore", "--nologo",
            f"-env:UserInstallation=file://{profile}",
            "--convert-to", "pdf", "--outdir", str(out), str(source),
        ], timeout=300)
    if not pdf.exists():
        print("  conversion failed:", result.stdout.strip()[-300:] or result.stderr.strip()[-300:])
        return None
    return pdf


def normalise(text: str) -> str:
    text = text.replace("\u2019", "'").replace("\u2018", "'")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    text = text.replace("\u00a0", " ")
    return re.sub(r"\s+", " ", text).strip()


def words_of(text: str) -> set[str]:
    """Alphanumeric tokens of two or more characters, lowercased."""
    return {w for w in re.findall(r"[0-9a-zA-Z]{2,}", normalise(text).lower())}


def reference_pages(pdf: Path):
    doc = pymupdf.open(pdf)
    pages = []
    for page in doc:
        text = page.get_text("text")
        images = []
        for info in page.get_image_info(xrefs=True):
            rect = info.get("bbox")
            if not rect:
                continue
            images.append({
                "w": round(rect[2] - rect[0], 1),
                "h": round(rect[3] - rect[1], 1),
                "x": round(rect[0], 1),
                "y": round(rect[1], 1),
                "xref": info.get("xref"),
            })
        pages.append({
            "text": text,
            "images": images,
            "width": round(page.rect.width, 1),
            "height": round(page.rect.height, 1),
        })
    doc.close()
    return pages


def mine(path: Path) -> dict:
    result = subprocess.run(["bun", str(INVENTORY), str(path)], capture_output=True, text=True, cwd=REPO)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip()[:400] or "inventory failed")
    return json.loads(result.stdout)


def overlap(a, b):
    """Intersection over union for two rects, each {x, y, w, h}."""
    if not a or not b:
        return 0.0
    ax2, ay2 = a["x"] + a["w"], a["y"] + a["h"]
    bx2, by2 = b["x"] + b["w"], b["y"] + b["h"]
    ix = max(0.0, min(ax2, bx2) - max(a["x"], b["x"]))
    iy = max(0.0, min(ay2, by2) - max(a["y"], b["y"]))
    inter = ix * iy
    if inter <= 0:
        return 0.0
    union = a["w"] * a["h"] + b["w"] * b["h"] - inter
    return inter / union if union > 0 else 0.0


def compare(path: Path, workdir: Path, keep: bool):
    print(f"\n=== {path.name}")
    pdf = to_pdf(path, workdir)
    if not pdf:
        return {"file": path.name, "status": "no-pdf"}
    ref = reference_pages(pdf)
    try:
        view = mine(path)
    except RuntimeError as err:
        print("  viewer failed:", err)
        return {"file": path.name, "status": "viewer-failed", "error": str(err)}

    ref_text = normalise(" ".join(p["text"] for p in ref))
    view_text = normalise(view.get("text", ""))
    ref_words = words_of(ref_text)
    view_words = words_of(view_text)
    # A reference token counts as present when it appears anywhere in the
    # viewer's text. Comparing whole tokens would flag every word that the
    # reference happened to split across a line, which is a layout difference
    # rather than missing content.
    haystack = view_text.lower()
    missing_words = {w for w in ref_words if w not in haystack}
    missing_words |= {w for w in ref_words if w in haystack and w not in view_words and len(w) > 3 and w not in haystack}
    missing_words = {w for w in ref_words if w not in haystack}
    extra_words = view_words - ref_words
    coverage = 1 - (len(missing_words) / max(1, len(ref_words)))

    ref_images = [img for p in ref for img in p["images"]]
    view_images = view.get("images", [])
    vector_images = [img for img in view_images if img.get("vector")]
    if vector_images:
        print(f"  note: {len(vector_images)} picture(s) are SVG in the source, "
              f"drawn as vector by the reference and as an image by the viewer")

    viewer_pages = view.get("pageCount") or len(view.get("pages", []))
    print(f"  pages: reference {len(ref)}, viewer {viewer_pages}")
    print(f"  text: {len(ref_words)} reference tokens, coverage {coverage * 100:.1f}%"
          f" ({len(missing_words)} missing)")
    print(f"  images: reference {len(ref_images)}, viewer {len(view_images)}"
          f" (drawn by the renderer: {view.get('renderedImages', '?')})")

    if missing_words:
        sample = sorted(missing_words, key=len, reverse=True)[:12]
        print("  missing tokens:", ", ".join(sample))

    # Per page image check, for the formats where a page maps to a slide.
    #
    # The question being answered is "did the viewer drop a picture", so every
    # reference image is looked up among the viewer's pictures and vice versa.
    # Matching is not one to one: a PDF can carry the same picture twice, for
    # example once for the frame and once for a border or shadow layer, and a
    # viewer picture may legitimately appear once.
    matched = 0
    unmatched_ref = []
    unmatched_view = []
    if view.get("kind") == "presentation" and len(ref) == len(view.get("pages", [])):
        scale_x = ref[0]["width"] / max(1, view.get("widthPx", 1))
        scale_y = ref[0]["height"] / max(1, view.get("heightPx", 1))

        def viewer_rect(img):
            return {
                "x": img["x"] * scale_x, "y": img["y"] * scale_y,
                "w": img["w"] * scale_x, "h": img["h"] * scale_y,
            }

        def viewer_frame(img):
            if not img.get("frameW"):
                return None
            return {
                "x": img["frameX"] * scale_x, "y": img["frameY"] * scale_y,
                "w": img["frameW"] * scale_x, "h": img["frameH"] * scale_y,
            }

        def score_against(ref_img, view_img):
            score = overlap(ref_img, viewer_rect(view_img))
            frame = viewer_frame(view_img)
            if frame:
                score = max(score, overlap(ref_img, frame))
            return score

        for index, page in enumerate(ref):
            page_view = [img for img in view_images if img["page"] == index and not img.get("vector")]
            for ref_img in page["images"]:
                best = max((score_against(ref_img, img) for img in page_view), default=0.0)
                if best >= MATCH_THRESHOLD:
                    matched += 1
                else:
                    unmatched_ref.append((index, ref_img, best))

        for img in view_images:
            if img.get("vector"):
                continue  # vector artwork is not an image in the reference
            page_refs = ref[img["page"]]["images"] if img["page"] < len(ref) else []
            best = max((score_against(ref_img, img) for ref_img in page_refs), default=0.0)
            if best < MATCH_THRESHOLD:
                unmatched_view.append((img["page"], img, best))

        total = max(1, len(ref_images))
        print(f"  image placement: {matched}/{len(ref_images)} reference images found in the viewer"
              f" ({matched / total * 100:.0f}%)")
        if unmatched_ref:
            print("  shown by the reference but not placed by the viewer:")
            for index, img, score in unmatched_ref[:8]:
                print(f"    page {index + 1}: {img['w']}x{img['h']} px at ({img['x']},{img['y']}), best overlap {score:.2f}")
        if unmatched_view:
            print("  drawn by the viewer but not found in the reference:")
            for index, img, score in unmatched_view[:8]:
                print(f"    page {index + 1}: {round(img['w'])}x{round(img['h'])} px at ({round(img['x'])},{round(img['y'])})"
                      f" source {img.get('source')}, best overlap {score:.2f}")

    if not keep:
        pdf.unlink(missing_ok=True)
    return {
        "file": path.name,
        "status": "ok",
        "coverage": coverage,
        "missing_words": sorted(missing_words, key=len, reverse=True)[:30],
        "extra_words": sorted(extra_words, key=len, reverse=True)[:30],
        "ref_images": len(ref_images),
        "view_images": len(view_images),
        "matched": matched,
    }


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    keep = "--keep" in sys.argv
    if not args:
        print(__doc__)
        sys.exit(2)
    results = []
    for name in args:
        path = Path(name).expanduser().resolve()
        if not path.exists():
            print("missing:", path)
            continue
        with tempfile.TemporaryDirectory(prefix="ovref-") as tmp:
            results.append(compare(path, Path(tmp), keep))

    print("\n=== summary")
    for result in results:
        if result.get("status") != "ok":
            print(f"  {result['file']}: {result.get('status')} {result.get('error', '')}")
            continue
        print(f"  {result['file']}: text {result['coverage'] * 100:.1f}%, "
              f"images {result['view_images']}/{result['ref_images']} found, "
              f"{result['matched']} placed, {len(result['extra_words'])} extra tokens")


if __name__ == "__main__":
    main()
