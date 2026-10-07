#!/usr/bin/env python3
"""
Crypto Analysis one-page summary.

Reads the Crypto Analysis workbook (sheet "Grades") and writes a one-page landscape Word document:
a two-column table of every coin (asset, group, score, coverage, tier) plus highlights that are
computed from the workbook. It retrieves NO new data and changes NO file.

Usage:
    python3 crypto_analysis_summary.py                      # newest workbook found in /mnt/project or /mnt/user-data/uploads
    python3 crypto_analysis_summary.py WORKBOOK.xlsx        # a specific workbook
    python3 crypto_analysis_summary.py WORKBOOK.xlsx --out "/mnt/user-data/outputs/Crypto Analysis Summary.docx" --date 2026-10-06
"""
import sys, re, os, glob, math, argparse, datetime

try:
    import openpyxl
except ImportError:
    os.system(f"{sys.executable} -m pip install openpyxl --break-system-packages -q")
    import openpyxl
try:
    from docx import Document
except ImportError:
    os.system(f"{sys.executable} -m pip install python-docx --break-system-packages -q")
    from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

TIER_NAMES = {1: "Exceptional", 2: "Strong", 3: "Average/Decent", 4: "Weak"}
TIER_FILL = {"Exceptional": "C6E0B4", "Strong": "DDEBF7", "Average/Decent": "FFF2CC",
             "Weak": "F8CBAD", "Not rated": "E7E6E6"}
GROUP_LABEL = {"Money coin": "Money coin", "Chain / protocol": "Chain", "Memecoin": "Memecoin",
               "Stablecoin / RWA": "Stable/RWA"}
LINE_NAMES = {1: "dilution", 2: "new supply", 3: "insider share", 4: "holder value capture",
              5: "exit liquidity", 6: "usage trend", 7: "sentiment", 8: "incidents",
              9: "regulatory", 10: "admin control", 11: "purpose", 12: "backers",
              13: "adoption", 14: "size", 15: "market-share trend"}


def version_key(p):
    m = re.search(r"V(\d+(?:\.\d+)*)", os.path.basename(p))
    return (tuple(int(x) for x in m.group(1).split(".")) if m else (0,), os.path.getmtime(p))


def find_workbook(arg):
    if arg:
        return arg
    cands = []
    for folder in ("/mnt/project", "/mnt/user-data/uploads", "."):
        cands += [p for p in glob.glob(os.path.join(folder, "Crypto Analysis V*.xlsx"))]
    if not cands:
        sys.exit("No 'Crypto Analysis V*.xlsx' found. Pass the path as the first argument.")
    cands.sort(key=version_key)
    if len(cands) > 1:
        print("WARNING: several workbooks found; using the highest version:")
        for c in cands:
            print("   ", c)
    return cands[-1]


def fmt_date(iso):
    try:
        d = datetime.date.fromisoformat(iso)
        return f"{d.day} {d.strftime('%b %y')}"
    except Exception:
        return iso


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def set_cell(cell, text, size, bold=False, fill=None, italic=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.name = "Arial"
    r.bold = bold
    r.italic = italic
    if fill:
        shade(cell, fill)


def set_borders_and_margins(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "BFBFBF")
        borders.append(el)
    tblPr.append(borders)
    mar = OxmlElement("w:tblCellMar")
    for edge, w in (("top", 10), ("bottom", 10), ("left", 50), ("right", 50)):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), str(w))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def read_workbook(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Grades" not in wb.sheetnames:
        sys.exit('Sheet "Grades" not found in this workbook.')
    ws = wb["Grades"]
    head = [c.value for c in ws[4]]
    col = {h: i for i, h in enumerate(head) if h}
    lines = [col[f"L{n}"] for n in range(1, 16)]
    coins = []
    for r in ws.iter_rows(min_row=5, values_only=True):
        t, score = r[col["Ticker"]], r[col["Score"]]
        if not t or not isinstance(score, (int, float)):
            continue
        last = r[col["Last evaluated"]]
        if hasattr(last, "isoformat"):
            last = last.date().isoformat() if hasattr(last, "date") else last.isoformat()
        coins.append(dict(t=str(t), group=r[col["Group"]], score=float(score), cov=float(r[col["Coverage"]]),
                          tier=r[col["Tier"]], bands=[r[i] for i in lines], stier=r[col["Score tier #"]],
                          ftier=r[col["Final tier #"]], evald=str(last) if last else "n/a"))
    if not coins:
        sys.exit("No calculated scores found on Grades. Open the file in Excel (or recalculate it) and save, then try again.")
    dag = set()
    if "Usage method" in wb.sheetnames:
        for r in wb["Usage method"].iter_rows(min_row=4, values_only=True):
            if r and r[0] and len(r) > 8 and isinstance(r[8], str) and r[8].strip().lower().startswith("yes"):
                dag.add(str(r[0]))
    return coins, dag


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook", nargs="?")
    ap.add_argument("--out")
    ap.add_argument("--date")
    a = ap.parse_args()
    path = find_workbook(a.workbook)
    coins, dag = read_workbook(path)

    # as-of date = newest Last evaluated date (or --date)
    date = a.date or max((c["evald"] for c in coins if c["evald"] != "n/a"), default=datetime.date.today().isoformat())

    # tier labels
    rows = []
    for c in coins:
        name = c["tier"]
        if not name or str(name).lower().startswith("not rated") or c["cov"] < 0.70:
            label, order = "Not rated", 4
        else:
            label = name
            order = {"Exceptional": 0, "Strong": 1, "Average/Decent": 2, "Weak": 3}.get(name, 5)
            capped = (c["group"] == "Memecoin") or (c["stier"] is not None and c["ftier"] is not None and c["stier"] != c["ftier"])
            if name == "Average/Decent" and capped:
                label = "Avg/Decent (capped)"
        c["label"], c["order"] = label, order
        rows.append(c)
    rows.sort(key=lambda c: (c["order"], -c["score"]))
    n = len(rows)

    # ------------------------------------------------------------ document
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Inches(11), Inches(8.5)
    sec.left_margin = sec.right_margin = Inches(0.45)
    sec.top_margin, sec.bottom_margin = Inches(0.35), Inches(0.3)
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(8)

    fs = 7 if n <= 44 else 6
    body = 7.5 if n <= 44 else 6.5

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("Crypto Analysis: One-Page Summary")
    r.bold = True
    r.font.size = Pt(15)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(f"{n} assets that trade spot on Coinbase, Kraken or Binance.US  |  Data as of {date}  |  "
                  f"Score, rating (score ÷ 10) and tier only; no buy/sell calls")
    r.italic = True
    r.font.size = Pt(7.5)

    half = math.ceil(n / 2)
    widths = [0.65, 0.85, 0.5, 0.45, 1.4, 0.75]
    spacer = 0.2
    NC = len(widths)
    tbl = doc.add_table(rows=half + 1, cols=2 * NC + 1)
    tbl.autofit = False
    set_borders_and_margins(tbl)
    hdr = ["Asset", "Group", "Score", "Cov.", "Tier", "Last eval."]
    allw = widths + [spacer] + widths
    for ri in range(half + 1):
        for ci in range(2 * NC + 1):
            tbl.cell(ri, ci).width = Inches(allw[ci])
    grid = tbl._tbl.tblGrid
    for gc, w in zip(grid.findall(qn('w:gridCol')), allw):
        gc.set(qn('w:w'), str(int(w * 1440)))
    for side in (0, 1):
        off = side * (NC + 1)
        for i, h in enumerate(hdr):
            set_cell(tbl.cell(0, off + i), h, fs, bold=True, fill="D9D9D9")
    set_cell(tbl.cell(0, NC), "", fs)
    for ri in range(half):
        set_cell(tbl.cell(ri + 1, NC), "", fs)
        for side in (0, 1):
            idx = ri + side * half
            off = side * (NC + 1)
            if idx >= n:
                for i in range(NC):
                    set_cell(tbl.cell(ri + 1, off + i), "", fs)
                continue
            c = rows[idx]
            vals = [c["t"] + (" †" if c["t"] in dag else ""), GROUP_LABEL.get(c["group"], str(c["group"])),
                    f"{c['score']:.1f}", f"{round(c['cov'] * 100)}%", c["label"], fmt_date(c["evald"])]
            for i, v in enumerate(vals):
                set_cell(tbl.cell(ri + 1, off + i), v, fs, bold=(i in (0, 2)),
                         fill=(TIER_FILL.get(c["label"].replace("Avg/Decent (capped)", "Average/Decent")) if i == 4 else None))

    # legend
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("Tiers: Exceptional 90+, Strong 75–89.9, Average/Decent 60–74.9, Weak below 60. Cov. = share of the "
                  "applicable points backed by a retrieved fact; below 70% no tier is printed (Not rated). "
                  "“Capped” = held at Average/Decent by the memecoin cap or by a red flag. † = the usage trend is "
                  "within 5 points of a grade cut-off. Chain = chain/protocol. Last eval. = the coin's Last evaluated date on the Grades sheet.")
    r.italic = True
    r.font.size = Pt(6.5)

    def bullet(head, text):
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.space_before = Pt(0)
        r1 = p.add_run(head)
        r1.bold = True
        r1.font.size = Pt(body)
        r2 = p.add_run(text)
        r2.font.size = Pt(body)

    # highlights (all computed from the workbook)
    by = {k: [] for k in ("Exceptional", "Strong", "Average/Decent", "Avg/Decent (capped)", "Weak", "Not rated")}
    for c in rows:
        by[c["label"]].append(c)

    def lst(cs, with_score=True):
        return ", ".join(f"{c['t']} {c['score']:.1f}" if with_score else c["t"] for c in cs)

    parts = []
    if by["Exceptional"]: parts.append(f"Exceptional ({len(by['Exceptional'])}): {lst(by['Exceptional'])}.")
    if by["Strong"]: parts.append(f"Strong ({len(by['Strong'])}): {lst(by['Strong'])}.")
    ad = by["Average/Decent"]
    if ad: parts.append(f"Average/Decent ({len(ad)}): {lst(ad)}.")
    if by["Avg/Decent (capped)"]: parts.append(f"Held at Average/Decent by a cap ({len(by['Avg/Decent (capped)'])}): {lst(by['Avg/Decent (capped)'])}.")
    if by["Weak"]: parts.append(f"Weak ({len(by['Weak'])}): {lst(by['Weak'])}.")
    if by["Not rated"]: parts.append(f"Not rated, coverage under 70% ({len(by['Not rated'])}): " + ", ".join(f"{c['t']} (provisional {c['score']:.1f}, {round(c['cov']*100)}%)" for c in by["Not rated"]) + ".")
    bullet("Where each asset sits: ", " ".join(parts))

    near = []
    for c in rows:
        for line in (90, 75, 60):
            if abs(c["score"] - line) <= 1.0 and c["label"] != "Not rated":
                near.append(f"{c['t']} {c['score']:.1f} (line {line})")
                break
    thin = [c for c in rows if c["label"] != "Not rated" and c["cov"] < 0.80]
    txt = ""
    if near: txt += "Within 1 point of a tier line, so one fact can change the tier: " + ", ".join(near) + ". "
    if thin: txt += "Rated on thinner evidence (coverage under 80%): " + ", ".join(f"{c['t']} {round(c['cov']*100)}%" for c in thin) + "."
    if txt: bullet("Fragile results: ", txt.strip())

    counts = []
    for i in range(15):
        b = sum(1 for c in rows if c["bands"][i] in (None, ""))
        if b: counts.append((b, i + 1))
    counts.sort(reverse=True)
    if counts:
        bullet("Most common gaps (lines left blank, never scored): ",
               "; ".join(f"line {ln} {LINE_NAMES[ln]} blank for {b} of {n} assets" for b, ln in counts[:4]) + ".")

    bullet("Read with care: ",
           "A blank line is never scored, so a coin with many blanks can score higher than it would with full data: read "
           "coverage next to the score. Market-share trend barely separates coins in a rising market and includes newly "
           "unlocked supply. Exit liquidity, sentiment and prices are one-day snapshots. Hack records are matched by name "
           "and chain, so smaller incidents can be missed. Many adoption calls rest on crypto media or the projects’ own "
           "announcements.")

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(f"This is an analytical framework, not financial advice. Ratings do not predict price.   "
                  f"Source: {os.path.basename(path)} (sheet Grades).")
    r.italic = True
    r.font.size = Pt(6.5)

    out = a.out or f"/mnt/user-data/outputs/Crypto Analysis Summary {date}.docx"
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    doc.save(out)
    print(f"Workbook used : {path}")
    print(f"Assets        : {n}   Data as of: {date}")
    print(f"Saved         : {out}")


if __name__ == "__main__":
    main()
