#!/usr/bin/env python3
"""
check_deck_render.py — scan the RENDERED deck for truncated text.

The PPTX XML can be perfectly clean while the delivered PDF shows "…": the renderer
clips any text that does not fit its box, and a one-line-high box silently truncates a
two-line string even with word_wrap on. Checking the XML therefore proves nothing —
this checks the artefact people actually see.

Usage: python3 tools/check_deck_render.py [deck.pdf]
"""
import glob, os, re, subprocess, sys, collections, xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)


def find_pdf():
    hits = sorted(glob.glob(os.path.join(REPO, "courseware", "*-v*.pdf")))
    hits = [h for h in hits if not os.path.basename(h).startswith(("LP-", "LG-"))]
    return hits[-1] if hits else None


def main():
    pdf = sys.argv[1] if len(sys.argv) > 1 else find_pdf()
    if not pdf or not os.path.exists(pdf):
        print("deck PDF not found — render the .pptx first")
        return 1
    text = subprocess.run(["pdftotext", pdf, "-"], capture_output=True, text=True).stdout
    pages = text.split("\f")
    hits = collections.defaultdict(list)
    for i, pg in enumerate(pages, 1):
        for line in pg.split("\n"):
            line = line.strip()
            if "…" in line:
                hits[i].append(line)

    total = sum(len(v) for v in hits.values())
    print(f"{os.path.basename(pdf)}: {len(pages)} page(s)")
    ok = True
    if total:
        ok = False
        print(f"FAIL — {total} truncated string(s) on {len(hits)} page(s):")
        for p in sorted(hits):
            print(f"  page {p}:")
            for l in hits[p][:6]:
                print(f"    {l[:90]}")
    else:
        print("  no truncated text")

    # --- text-line collisions ---------------------------------------------------
    # An ellipsis grep cannot see a label whose extra wrapped line strikes through the
    # caption beneath it. Compare the rendered line boxes and fail on any overlap.
    coll = collect_overlaps(pdf)
    if coll:
        ok = False
        print(f"FAIL — {len(coll)} overlapping text line(s):")
        for pg, a, b, amt in coll[:12]:
            print(f"    page {pg}: {a[:38]!r} over {b[:38]!r} by {amt:.1f}pt")
    else:
        print("  no overlapping text lines")

    print("PASS — rendered deck is clean." if ok else "FAIL — fix the above and re-render.")
    return 0 if ok else 1


def collect_overlaps(pdf, tol=0.6):
    """Overlapping text lines in the rendered PDF, via pdftotext -bbox-layout."""
    xml = subprocess.run(["pdftotext", "-bbox-layout", pdf, "-"],
                         capture_output=True, text=True).stdout
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return []
    ns = {"x": "http://www.w3.org/1999/xhtml"}
    out = []
    for pi, page in enumerate(root.iter("{http://www.w3.org/1999/xhtml}page"), 1):
        lines = []
        for ln in page.iter("{http://www.w3.org/1999/xhtml}line"):
            try:
                box = (float(ln.get("xMin")), float(ln.get("yMin")),
                       float(ln.get("xMax")), float(ln.get("yMax")))
            except (TypeError, ValueError):
                continue
            text = "".join(w.text or "" for w in
                           ln.iter("{http://www.w3.org/1999/xhtml}word")).strip()
            if text:
                lines.append((box, text))
        for i in range(len(lines)):
            for j in range(i + 1, len(lines)):
                (ax0, ay0, ax1, ay1), at = lines[i]
                (bx0, by0, bx1, by1), bt = lines[j]
                ox = min(ax1, bx1) - max(ax0, bx0)
                oy = min(ay1, by1) - max(ay0, by0)
                if ox > tol and oy > tol:
                    out.append((pi, at, bt, oy))
    return out


if __name__ == "__main__":
    sys.exit(main())
