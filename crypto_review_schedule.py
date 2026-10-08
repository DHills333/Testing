#!/usr/bin/env python3
"""
Crypto Analysis review schedule.

Keeps a reassessment slot for every coin on the Grades sheet: coins you hold every 13 weeks (quarterly), every other
coin every 26 weeks (twice a year). Slots are spread so that about the same number of coins comes due each week.
It writes the Schedule sheet of the workbook, prints this week's list, and writes a calendar file (.ics) with one
all-day event per Monday listing that week's coins.

Rules
- A coin's slot stays fixed. When its "Last evaluated" date on Grades moves forward to within half an interval of its
  slot (or later), the slot moves on by one interval (13 or 26 weeks). A coin not yet reassessed stays on the list as
  overdue.
- New coin on Grades: placed in the least-busy week within two weeks either side of (last evaluated + interval).
- Coin you start holding: if its next slot is more than 13 weeks away, it moves to the least-busy week in the next
  13 weeks. Coin you stop holding: keeps its slot, then every 26 weeks.
- First run (no Schedule sheet yet): every coin is placed in the least-busy week in its first interval, so the first
  round is staggered; held coins are placed first.

Usage:
    python3 crypto_review_schedule.py WORKBOOK.xlsx --tracker TRACKER.xlsm --out NEW_WORKBOOK.xlsx [--ics FILE.ics]
                                      [--today YYYY-MM-DD] [--weeks 52]
Without --tracker, the Held column already on the Schedule sheet is used.
It retrieves no data. It writes only the files named in --out and --ics.
"""
import sys, os, re, argparse, datetime, shutil, subprocess, tempfile

try:
    import openpyxl
except ImportError:
    os.system(f"{sys.executable} -m pip install openpyxl --break-system-packages -q")
    import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

D = datetime.date
WK = datetime.timedelta(weeks=1)
HELD_WEEKS, OTHER_WEEKS = 13, 26


def monday(d):
    return d - datetime.timedelta(days=d.weekday())


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, D):
        return v
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", str(v))
    return D(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def read_grades(wb):
    g = wb["Grades"]
    head = {c.value: i for i, c in enumerate(g[4]) if c.value}
    coins = []
    for r in g.iter_rows(min_row=5, values_only=True):
        t = r[head["Ticker"]]
        if t:
            coins.append(dict(t=str(t), last=to_date(r[head["Last evaluated"]]), tier=r[head["Tier"]], score=r[head["Score"]]))
    return coins


def read_held(tracker):
    """Tickers held (quantity > 0), mapped through the tracker's Rated as column. Uses typed inputs only."""
    wb = openpyxl.load_workbook(tracker, data_only=True)
    pr, ra, pf = wb["Prices"], wb["Ratings"], wb["Portfolio"]
    rated_as = {}
    for i in range(34):
        coin = pr.cell(8 + i, 1).value
        if coin:
            b = ra.cell(6 + i, 2).value
            rated_as[str(coin)] = None if b == "none" else (str(b) if b else str(coin))
    held = set()
    for r in range(10, 70):
        coin, q = pf.cell(r, 2).value, pf.cell(r, 3).value
        if coin and isinstance(q, (int, float)) and q > 0:
            k = rated_as.get(str(coin), str(coin))
            if k:
                held.add(k)
    return held


def read_schedule(wb):
    if "Schedule" not in wb.sheetnames:
        return None
    s = wb["Schedule"]
    out = {}
    for r in range(7, s.max_row + 1):
        t = s.cell(r, 1).value
        if t:
            out[str(t)] = dict(held=str(s.cell(r, 2).value or "").lower() == "yes", next=to_date(s.cell(r, 5).value),
                               seen=to_date(s.cell(r, 8).value))
    return out


def hits(next_d, every, week):
    return week >= next_d and ((week - next_d).days // 7) % every == 0


def load(plan, week, skip=None):
    return sum(1 for t, p in plan.items() if t != skip and p["next"] and hits(p["next"], p["every"], week))


def least_busy(plan, first, last, every, skip=None):
    """Week in [first, last] whose recurring load (over one cycle of this coin) is lowest; ties -> earliest."""
    best, best_key = None, None
    w = first
    while w <= last:
        key = (sum(load(plan, w + k * every * WK, skip) for k in range(4)), w)
        if best_key is None or key < best_key:
            best, best_key = w, key
        w += WK
    return best


def recalc(path):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    with tempfile.TemporaryDirectory() as td:
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", td, path],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=240)
        conv = os.path.join(td, os.path.basename(path))
        if os.path.exists(conv):
            shutil.copyfile(conv, path)
            return True
    return False


def ics_text(lines):
    out = []
    for line in lines:
        b = line.encode("utf-8")
        while len(b) > 74:
            cut = 74
            while (b[cut] & 0xC0) == 0x80:
                cut -= 1
            out.append(b[:cut].decode("utf-8")); b = b" " + b[cut:]
        out.append(b.decode("utf-8"))
    return "\r\n".join(out) + "\r\n"


def esc(s):
    return str(s).replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--tracker")
    ap.add_argument("--out", required=True)
    ap.add_argument("--ics")
    ap.add_argument("--today")
    ap.add_argument("--weeks", type=int, default=52)
    a = ap.parse_args()
    today = to_date(a.today) or D.today()
    this_mon = monday(today)

    wb = openpyxl.load_workbook(a.workbook)
    wv = openpyxl.load_workbook(a.workbook, data_only=True)
    coins = read_grades(wv)
    known = {c["t"] for c in coins}
    old = read_schedule(wb)
    first_run = old is None
    old = old or {}
    held = read_held(a.tracker) if a.tracker else {t for t, p in old.items() if p["held"]}
    unrated = sorted(held - known)

    plan, changes = {}, []
    # keep existing slots first, so new placements see the real load
    for c in coins:
        t = c["t"]; h = t in held; every = HELD_WEEKS if h else OTHER_WEEKS
        p = old.get(t)
        if p and p["next"]:
            nxt, seen = p["next"], p["seen"]
            if c["last"] and (seen is None or c["last"] > seen):
                moved = False
                while c["last"] >= nxt - (every // 2) * WK:
                    nxt += every * WK; moved = True
                if moved:
                    changes.append(f"{t}: reassessed {c['last']}, next review moved to week of {nxt}")
                seen = c["last"]
            plan[t] = dict(every=every, next=nxt, seen=seen, held=h, was_held=p["held"])
    order = [c for c in coins if c["t"] in held] + [c for c in coins if c["t"] not in held]
    for c in order:
        t = c["t"]; h = t in held; every = HELD_WEEKS if h else OTHER_WEEKS
        if t in plan:
            p = plan[t]
            if h and not p["was_held"] and p["next"] > this_mon + HELD_WEEKS * WK:
                p["next"] = least_busy(plan, this_mon + WK, this_mon + HELD_WEEKS * WK, every, skip=t)
                changes.append(f"{t}: now held, next review moved to week of {p['next']} (every 13 weeks)")
            elif not h and p["was_held"]:
                changes.append(f"{t}: no longer held, now every 26 weeks (next week of {p['next']})")
            continue
        if first_run or not c["last"]:
            lo, hi = this_mon + WK, this_mon + every * WK
        else:
            due = monday(c["last"]) + every * WK
            lo, hi = max(this_mon + WK, due - 2 * WK), max(this_mon + WK, due + 2 * WK)
        plan[t] = dict(every=every, next=None, seen=c["last"], held=h, was_held=h)
        plan[t]["next"] = least_busy(plan, lo, hi, every, skip=t)
        if not first_run:
            changes.append(f"{t}: added to the schedule, first review week of {plan[t]['next']}")
    # coins removed from Grades simply drop off

    # ---------------- write the Schedule sheet
    old_cal = {}
    if "Schedule" in wb.sheetnames:
        os_ = wb["Schedule"]
        for r in range(7, os_.max_row + 1):
            wv_ = to_date(os_.cell(r, 10).value)
            if wv_:
                old_cal[wv_] = (os_.cell(r, 11).value or "", os_.cell(r, 12).value or "")
        idx = wb.sheetnames.index("Schedule"); del wb["Schedule"]
    else:
        idx = wb.sheetnames.index("Grades") + 1
    s = wb.create_sheet("Schedule", idx)
    F = "Arial"
    bold, norm, note = Font(name=F, size=10, bold=True), Font(name=F, size=10), Font(name=F, size=9, italic=True, color="595959")
    hdr_fill, in_fill = PatternFill("solid", fgColor="D9D9D9"), PatternFill("solid", fgColor="FFF2CC")
    thin = Side(style="thin", color="BFBFBF"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
    s["A1"] = "SCHEDULE: when each coin is due for reassessment"; s["A1"].font = Font(name=F, size=14, bold=True)
    s["A2"] = ("Held coins every 13 weeks, every other coin every 26 weeks; slots are spread so about the same number comes "
               "due each week. Written by crypto_review_schedule.py: run it after any coin is added or reassessed, or after "
               "buying or selling a coin, then re-import the calendar file.")
    s["A2"].font = note
    s["A3"] = f"Schedule written {today.isoformat()} (week of {this_mon.isoformat()})."; s["A3"].font = note
    heads = ["Ticker", "Held", "Every (weeks)", "Last evaluated", "Next review (week of)", "Status", "Tier", "Last evaluated when scheduled"]
    for i, h in enumerate(heads, 1):
        c = s.cell(6, i, h); c.font = bold; c.fill = hdr_fill; c.border = box; c.alignment = Alignment(wrap_text=True, vertical="center")
    row = 7
    for c in coins:
        p = plan[c["t"]]
        vals = [c["t"], "yes" if p["held"] else "no", None, None, p["next"], None, None, p["seen"]]
        for i, v in enumerate(vals, 1):
            cell = s.cell(row, i, v); cell.font = norm; cell.border = box
        s.cell(row, 1).font = bold
        s.cell(row, 2).fill = in_fill; s.cell(row, 2).font = Font(name=F, size=10, color="0000FF")
        s.cell(row, 3, f'=IF(B{row}="yes",{HELD_WEEKS},{OTHER_WEEKS})')
        s.cell(row, 4, f'=IFERROR(INDEX(Grades!$X:$X,MATCH(A{row},Grades!$A:$A,0)),"")')
        s.cell(row, 6, f'=IF(E{row}="","",IF(E{row}<TODAY()-WEEKDAY(TODAY(),3),"Overdue",IF(E{row}=TODAY()-WEEKDAY(TODAY(),3),"This week",IF(E{row}=TODAY()-WEEKDAY(TODAY(),3)+7,"Next week",""))))')
        s.cell(row, 7, f'=IFERROR(INDEX(Grades!$G:$G,MATCH(A{row},Grades!$A:$A,0)),"")')
        for col in (5, 8):
            s.cell(row, col).number_format = "yyyy-mm-dd"
        row += 1
    last_row = row - 1
    dv = DataValidation(type="list", formula1='"yes,no"', allow_blank=False); s.add_data_validation(dv); dv.add(f"B7:B{last_row}")
    s.conditional_formatting.add(f"F7:F{last_row}", FormulaRule(formula=['$F7="Overdue"'], font=Font(color="C00000", bold=True)))
    s.conditional_formatting.add(f"F7:F{last_row}", FormulaRule(formula=['$F7="This week"'], fill=PatternFill("solid", fgColor="FFF2CC")))

    # calendar table
    weeks = []
    w = this_mon
    for _ in range(a.weeks + 1):
        due = [t for t in (c["t"] for c in coins) if hits(plan[t]["next"], plan[t]["every"], w)]
        over = [t for t in (c["t"] for c in coins) if w == this_mon and plan[t]["next"] < this_mon]
        weeks.append((w, [t for t in due if plan[t]["held"]], [t for t in due if not plan[t]["held"]], over))
        w += WK
    s.cell(5, 10, f"CALENDAR: the next {a.weeks} weeks").font = bold
    for i, h in enumerate(["Week of (Monday)", "Coins I hold (quarterly)", "Other coins (twice a year)", "Count"], 10):
        c = s.cell(6, i, h); c.font = bold; c.fill = hdr_fill; c.border = box; c.alignment = Alignment(wrap_text=True, vertical="center")
    for k, (w, hd, ot, ov) in enumerate(weeks):
        r = 7 + k
        vals = [w, ", ".join(hd + [f"{t} (overdue)" for t in ov if plan[t]["held"]]),
                ", ".join(ot + [f"{t} (overdue)" for t in ov if not plan[t]["held"]]), len(hd) + len(ot) + len(ov)]
        for i, v in enumerate(vals, 10):
            cell = s.cell(r, i, v); cell.font = norm; cell.border = box
        s.cell(r, 10).number_format = "yyyy-mm-dd"
    for col, wdt in zip("ABCDEFGHIJKLM", [8, 6, 8, 12, 13, 11, 16, 13, 3, 13, 30, 36, 7]):
        s.column_dimensions[col].width = wdt
    s.column_dimensions["H"].hidden = True
    s.freeze_panes = "A7"
    wb.save(a.out)
    recalculated = recalc(a.out)

    # ---------------- calendar file
    if a.ics:
        info = {c["t"]: c for c in coins}
        L = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Crypto Analysis//Review schedule//EN", "CALSCALE:GREGORIAN",
             "METHOD:PUBLISH", "X-WR-CALNAME:Crypto Reviews"]
        stamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        for w, hd, ot, ov in weeks:
            if w == this_mon or not (hd or ot):
                continue
            names = hd + ot
            desc = ["Reassess with the Crypto Analysis method (in a chat inside the Project: \"Reassess "
                    + ", ".join(names) + " using the Crypto Analysis method\")."]
            for t in names:
                desc.append(f"{t}: {'a coin I hold (every 13 weeks)' if t in hd else 'every 26 weeks'}")
            L += ["BEGIN:VEVENT", f"UID:crypto-review-{w.isoformat()}@crypto-analysis", f"DTSTAMP:{stamp}",
                  f"DTSTART;VALUE=DATE:{w.strftime('%Y%m%d')}", f"DTEND;VALUE=DATE:{(w + datetime.timedelta(days=1)).strftime('%Y%m%d')}",
                  f"SUMMARY:{esc('Crypto reviews: ' + ', '.join(names))}", f"DESCRIPTION:{esc(chr(10).join(desc))}",
                  "TRANSP:TRANSPARENT",
                  "BEGIN:VALARM", "ACTION:DISPLAY", f"DESCRIPTION:{esc('Crypto reviews this week: ' + ', '.join(names))}",
                  "TRIGGER:PT8H", "END:VALARM", "END:VEVENT"]
        L.append("END:VCALENDAR")
        with open(a.ics, "w", encoding="utf-8", newline="") as f:
            f.write(ics_text(L))

    # ---------------- report
    w0 = weeks[0]; w1 = weeks[1]
    def fmt(wk):
        hd, ot, ov = wk[1], wk[2], wk[3]
        parts = [", ".join(hd) + " (held)" if hd else "", ", ".join(ot) if ot else "",
                 ", ".join(ov) + " (overdue)" if ov else ""]
        return "; ".join(p for p in parts if p) or "nothing due"
    counts = [len(x[1]) + len(x[2]) for x in weeks[1:]]
    print(f"Workbook      : {a.workbook} -> {a.out}{' (recalculated)' if recalculated else ' (open in Excel to recalculate)'}")
    print(f"Held coins    : {', '.join(sorted(held & known)) or 'none'}")
    if unrated:
        print(f"Held but NOT on Grades (evaluate first): {', '.join(unrated)}")
    print(f"This week ({this_mon}): {fmt(w0)}")
    print(f"Next week ({w1[0]}): {fmt(w1)}")
    print(f"Coins per week over the next {a.weeks} weeks: min {min(counts)}, max {max(counts)}, average {sum(counts)/len(counts):.1f}")
    if first_run:
        print("First run: every coin placed in its first cycle so the reviews are staggered.")
    for ch in changes:
        print("Change        :", ch)
    new_cal = {w: (", ".join(hd), ", ".join(ot)) for w, hd, ot, ov in weeks[1:]}
    common = [w for w in new_cal if w in old_cal]
    if first_run or not old_cal:
        verdict = "new calendar: import it"
    elif any(new_cal[w] != (old_cal[w][0].replace(" (overdue)", ""), old_cal[w][1].replace(" (overdue)", "")) for w in common):
        verdict = "CHANGED since the last run: delete the Crypto Reviews calendar and import the new file"
    else:
        verdict = "unchanged since the last run: no need to re-import"
    print(f"Calendar      : {verdict}" + (f" ({a.ics})" if a.ics else ""))


if __name__ == "__main__":
    main()
