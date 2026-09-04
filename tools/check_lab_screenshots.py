#!/usr/bin/env python3
"""
check_lab_screenshots.py — guard the Learner Guide's lab screenshots.

md5 uniqueness is not enough: two captures of the same page taken seconds apart differ
by a few pixels and pass an md5 check while being the SAME picture to a reader. This
uses a perceptual hash so a visual duplicate fails the build.

Run before publishing:  python3 tools/check_lab_screenshots.py
"""
import glob, itertools, os, sys
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
SHOTS = os.path.join(REPO, ".claude", "skills", "courseware-build", "assets", "labs")
EXPECTED_LABS = 14
THRESHOLD = 4          # hamming distance below which two images read as the same


def phash(path):
    try:
        import imagehash
        return imagehash.phash(Image.open(path)), "phash"
    except ImportError:
        im = Image.open(path).convert("L").resize((16, 16))
        px = list(im.getdata())
        avg = sum(px) / len(px)
        return tuple(1 if p > avg else 0 for p in px), "avghash"


def distance(a, b):
    if isinstance(a, tuple):
        return sum(x != y for x, y in zip(a, b))
    return a - b


# Text that means the capture caught an error page or a stale failure banner rather
# than the lab's subject. A uniqueness check cannot see these — OCR-free heuristics
# would be fragile, so the capture step must assert on page text BEFORE saving; this
# list documents what to look for when reviewing.
BAD_PAGE_MARKERS = [
    "there's been a disconnect",
    "We can't find the page",
    "failed to import",
]


def main():
    files = sorted(glob.glob(os.path.join(SHOTS, "*.png")))
    problems = []

    # one screenshot per lab, 1..14
    have = set()
    for f in files:
        base = os.path.basename(f)
        try:
            have.add(int(base.split("-")[1]))
        except (IndexError, ValueError):
            problems.append(f"unparseable screenshot name: {base}")
    missing = [n for n in range(1, EXPECTED_LABS + 1) if n not in have]
    if missing:
        problems.append(f"no screenshot for lab(s): {missing}")

    # perceptual uniqueness
    hashes = {}
    mode = None
    for f in files:
        h, mode = phash(f)
        hashes[os.path.basename(f)] = h
    for a, b in itertools.combinations(sorted(hashes), 2):
        d = distance(hashes[a], hashes[b])
        if d <= THRESHOLD:
            problems.append(f"visually identical (distance {d}): {a} and {b}")

    print(f"lab screenshots: {len(files)} file(s), hash mode {mode}")
    if problems:
        print("FAIL")
        for p in problems:
            print("  -", p)
        return 1
    print(f"PASS — {len(files)} screenshots, one per lab, all visually distinct.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
