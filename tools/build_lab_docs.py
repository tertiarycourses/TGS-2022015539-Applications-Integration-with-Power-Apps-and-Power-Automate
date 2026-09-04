#!/usr/bin/env python3
"""
build_lab_docs.py — write one lab-NN.md per lab plus the labs/README.md index,
all generated from the single source (course_data.py + data_domainN.py) so the
labs can never drift from the slides, the LG or the assessment.
"""
import os, sys, glob

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
LABS = os.path.join(REPO, "labs")
BUILD = os.path.join(REPO, ".claude", "skills", "courseware-build", "build")
sys.path.insert(0, BUILD)

import course_data as C
import data_domain1 as d1, data_domain2 as d2, data_domain3 as d3, data_domain4 as d4

ALL = d1.DOMAIN1 + d2.DOMAIN2 + d3.DOMAIN3 + d4.DOMAIN4
TOPIC = {t["num"]: t for t in C.TOPICS}


def lab_folder(n):
    return os.path.join(LABS, f"lab-{n:02d}")


def packages_for(n):
    """the .zip files that live in this lab's folder"""
    f = lab_folder(n)
    if not os.path.isdir(f):
        return []
    return sorted(x for x in os.listdir(f) if x.endswith(".zip"))


def data_for(n):
    d = os.path.join(lab_folder(n), "data")
    if not os.path.isdir(d):
        return []
    return sorted(os.listdir(d))


def write_lab(a):
    n = a["num"]
    t = TOPIC[a["topic"]]
    folder = lab_folder(n)
    os.makedirs(folder, exist_ok=True)

    L = []
    L.append(f"# Lab {n} — {a['title']}")
    L.append("")
    L.append(f"**Course:** {C.TITLE} ({C.COURSE_CODE})  ")
    L.append(f"**Topic {t['num']}:** {t['title']}  ")
    L.append(f"**Maps to:** {a['objective']}")
    L.append("")
    L.append("---")
    L.append("")
    L.append("## What you are building")
    L.append("")
    L.append(a["desc"])
    L.append("")
    L.append(f"**Deliverable —** {a['build']}")
    L.append("")
    L.append(f"**Tools —** {a['services']}")
    L.append("")

    if a.get("environment"):
        L.append("> **Environment.** Every lab in this course runs in the Power Platform")
        L.append(f"> environment **{C.LAB_ENVIRONMENT}**.")
        L.append("> Check the environment picker in the top-right of the maker portal BEFORE you")
        L.append("> build anything — work created in the wrong environment cannot simply be moved.")
        L.append("")

    if a.get("flow_name"):
        L.append(f"> **Name the flow exactly** `{a['flow_name']}`.")
        L.append("> The trailing (DO NOT DELETE) marks it as courseware in a shared training tenant.")
        L.append("")
    if a.get("app_name"):
        L.append(f"> **Name the app exactly** `{a['app_name']}`.")
        L.append("")

    files = data_for(n)
    if files:
        L.append("## Data")
        L.append("")
        for f in files:
            L.append(f"- `data/{f}` — upload this to your OneDrive for Business before you start.")
        L.append("")
        L.append("Each workbook already contains a real Excel **Table**. Power Apps and Power")
        L.append("Automate can only bind to a named table, never to a loose range — if you build")
        L.append("your own workbook later, remember to Insert > Table first.")
        L.append("")

    pkgs = packages_for(n)
    if pkgs:
        L.append("## Prebuilt packages")
        L.append("")
        L.append("If you want to inspect or restore the finished flow instead of building it:")
        L.append("")
        for p in pkgs:
            if p.startswith("Solution-"):
                L.append(f"- `{p}` — **Dataverse solution**. Import via Solutions > Import solution.")
            else:
                L.append(f"- `{p}` — **legacy package**. Import via My flows > Import > Import Package (Legacy).")
        L.append("")
        L.append("Imported flows arrive **turned off** until you supply your own connection —")
        L.append("open the flow, re-authenticate each connector, then turn it on.")
        L.append("")

    L.append("## Steps")
    L.append("")
    for i, (text, cmd) in enumerate(a["steps"], 1):
        L.append(f"{i}. {text}")
        if cmd:
            if cmd.startswith("http"):
                L.append(f"   - <{cmd}>")
            else:
                L.append("")
                L.append("   ```")
                L.append(f"   {cmd}")
                L.append("   ```")
        L.append("")

    L.append("## Test it")
    L.append("")
    L.append(a["test"])
    L.append("")

    prev_n, next_n = n - 1, n + 1
    prev = next((x for x in ALL if x["num"] == prev_n), None)
    nxt = next((x for x in ALL if x["num"] == next_n), None)
    if prev or nxt:
        L.append("---")
        L.append("")
        if prev:
            L.append(f"◀ Previous: [Lab {prev['num']} — {prev['title']}](../lab-{prev_n:02d}/lab-{prev_n:02d}.md)  ")
        if nxt:
            L.append(f"▶ Next: [Lab {nxt['num']} — {nxt['title']}](../lab-{next_n:02d}/lab-{next_n:02d}.md)")
        L.append("")

    path = os.path.join(folder, f"lab-{n:02d}.md")
    with open(path, "w") as f:
        f.write("\n".join(L))
    return path


def write_index():
    L = []
    L.append(f"# Labs — {C.TITLE}")
    L.append("")
    L.append(f"**Course code:** {C.COURSE_CODE}  ")
    L.append(f"**TSC:** {C.TSC_TITLE} ({C.TSC_CODE}), {C.TSC_LEVEL}  ")
    L.append(f"**Environment:** `{C.LAB_ENVIRONMENT}`")
    L.append("")
    L.append("The labs are **progressive** — each one builds on the artefact the previous one")
    L.append("produced. Labs 1–2 decide *what* to integrate, Labs 3–6 build the automation,")
    L.append("Labs 7–10 build the apps, and Labs 11–14 join the two halves together and then")
    L.append("break, diagnose and repair the result.")
    L.append("")
    L.append("| Lab | Title | Topic | Maps to |")
    L.append("|---|---|---|---|")
    for a in ALL:
        codes = ", ".join(sorted(set(
            c for c in [w.strip(" ;,.") for w in a["objective"].replace(";", " ").split()]
            if len(c) == 2 and c[0] in "KA" and c[1].isdigit())))
        L.append(f"| {a['num']} | [{a['title']}](lab-{a['num']:02d}/lab-{a['num']:02d}.md) "
                 f"| {a['topic']} | {codes} |")
    L.append("")
    L.append("## Coverage")
    L.append("")
    L.append("Every TSC knowledge and ability statement is exercised by at least one lab.")
    L.append("")
    L.append("| Code | Statement |")
    L.append("|---|---|")
    for k, d in C.TSC_KNOWLEDGE:
        L.append(f"| {k} | {d} |")
    for k, d in C.TSC_ABILITIES:
        L.append(f"| {k} | {d} |")
    L.append("")
    L.append("## Packages")
    L.append("")
    L.append("The tenant used for this course has *Create in Dataverse solutions* enabled,")
    L.append("which disables legacy package import. Each lab folder therefore ships both")
    L.append("formats, and `Solution-All-Labs.zip` installs every flow at once.")
    L.append("")
    path = os.path.join(LABS, "README.md")
    with open(path, "w") as f:
        f.write("\n".join(L))
    return path


def main():
    made = [write_lab(a) for a in ALL]
    made.append(write_index())
    print(f"wrote {len(made)} file(s)")
    for m in made:
        print("  ", os.path.relpath(m, REPO))


if __name__ == "__main__":
    main()
