"""
Dry-run of the Mixer Compound Planner logic (same rules as MixerPlanner.bas).

Reads the cached values of Mixing_Production_Plan.xlsm (STOCK, Raw Data,
Recipe Details) and writes a preview workbook with:
  * Priority      - days-of-cover ranking, plan quantity, eligible mixers
  * MX01/MX02/MX03 - 1440-minute sequence with line-clearance blocks and shifts

Usage:
    python dry_run_planner.py <Mixing_Production_Plan.xlsm> <output.xlsx>

The xlsm is only read, never modified.
"""
import math
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# ---------------- Settings (same defaults as 'Planner Settings' sheet) -------
S = {
    "critical_days": 1.5,     # stock cover below this -> must plan today
    "watch_days": 2.5,        # cover below this -> plan if mixer has spare time
    "cont_days": 3.5,         # CONT compounds: produce this many days of requirement
    "include_watch": True,
    "day_minutes": 1440,
    "lc_minutes": 80,         # total line clearance incl. 4 cleanout batches
    "co_before": 3,           # cleanout batches before physical cleaning
    "co_after": 1,            # cleanout batch after physical cleaning
    "shift_minutes": 480,
    "cycle": {1: 6.5, 2: 6.5, 3: 6.5},   # minutes per mixer batch (ASSUMPTION - set real value)
    "default_co": "29000024",
    "max_compounds": 4,       # different 21 codes per mixer per day
    "last_running": {1: "", 2: "", 3: ""},
}
SHIFTS = ["A", "B", "C"]


# ---------------- tiny xlsx reader (values only, no openpyxl needed) ---------
def _col(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def read_workbook(path):
    z = zipfile.ZipFile(path)
    ss = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        ss = ["".join(t.text or "" for t in si.iter(NS + "t")) for si in root.findall(NS + "si")]
    wb = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rmap = {}
    for rel in re.findall(r"<Relationship [^>]*>", rels):
        rid = re.search(r'Id="([^"]+)"', rel).group(1)
        tgt = re.search(r'Target="([^"]+)"', rel).group(1)
        rmap[rid] = tgt
    sheets = {}
    for tag in re.findall(r"<sheet [^>]*>", wb):
        name = re.search(r'name="([^"]+)"', tag).group(1).replace("&amp;", "&")
        rid = re.search(r'r:id="([^"]+)"', tag).group(1)
        sheets[name] = "xl/" + rmap[rid]

    def read(name, maxrow=10 ** 7):
        cells = {}
        with z.open(sheets[name]) as fh:
            for _, el in ET.iterparse(fh):
                if el.tag == NS + "row":
                    if int(el.get("r")) > maxrow:
                        break
                    for c in el.findall(NS + "c"):
                        v = c.find(NS + "v")
                        if v is None:
                            continue
                        val = v.text
                        t = c.get("t")
                        if t == "s":
                            val = ss[int(val)]
                        elif t not in ("str", "e", "inlineStr"):
                            try:
                                val = float(val)
                            except ValueError:
                                pass
                        m = re.match(r"([A-Z]+)(\d+)", c.get("r"))
                        cells[(int(m.group(2)), _col(m.group(1)))] = val
                    el.clear()
        return cells

    return read


def code(v):
    """Normalise SAP codes: 22000699.0 -> '22000699'."""
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else str(v)
    return str(v).strip()


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# ---------------- master data ------------------------------------------------
def load_master(read):
    raw = read("Raw Data", 400)
    rawd = {}
    for r in range(2, 401):
        c22 = code(raw.get((r, 2)))
        if c22 and c22 not in rawd:
            rawd[c22] = {"c21": code(raw.get((r, 3))), "pallet": num(raw.get((r, 8))),
                         "desc": raw.get((r, 5)) or "", "colour": raw.get((r, 7)) or ""}
    rec = read("Recipe Details", 80)
    # (21 code col, weight col, fallback weight col) for MX01, MX02, MX03
    layout = {1: (2, 7, 8), 2: (12, 17, 17), 3: (21, 26, 26)}
    wts = {1: {}, 2: {}, 3: {}}
    for m, (cc, wc, wc2) in layout.items():
        for r in range(3, 81):
            c21 = code(rec.get((r, cc)))
            if not c21 or c21 in wts[m]:
                continue
            w = num(rec.get((r, wc))) or num(rec.get((r, wc2)))
            if w > 0:
                wts[m][c21] = w
    return rawd, wts


# ---------------- priority ---------------------------------------------------
def build_priority(read, rawd, wts):
    st = read("STOCK", 61)
    items = []
    for r in range(2, 62):
        c22 = code(st.get((r, 4)))
        if not c22 or not c22.startswith("22"):
            continue
        info = rawd.get(c22, {})
        stock = num(st.get((r, 7)))
        req = num(st.get((r, 8)))
        lim_raw = st.get((r, 11))
        cont = lim_raw is None or str(lim_raw).strip().upper() in ("CONT", "")
        remaining = None if cont else max(0, int(num(lim_raw)))
        cover = stock / req if req > 0 else 99
        w = {m: wts[m].get(info.get("c21", "")) for m in (1, 2, 3)}
        if not info.get("pallet") or not any(w.values()):
            status = "CHECK DATA"
        elif req <= 0:
            status = "NO DEMAND"
        elif not cont and remaining <= 0:
            status = "LIMIT DONE"
        elif cover < S["critical_days"]:
            status = "1-CRITICAL"
        elif cover < S["watch_days"]:
            status = "2-WATCH"
        else:
            status = "3-OK"
        plan = 0
        if status == "1-CRITICAL" or (status == "2-WATCH" and S["include_watch"]):
            plan = max(1, math.ceil(S["cont_days"] * req)) if cont else remaining
        items.append({"row": r, "c22": c22, "c21": info.get("c21", ""), "mat": st.get((r, 5)) or info.get("desc", ""),
                      "colour": st.get((r, 12)) or "", "stock": stock, "req": req, "cover": cover,
                      "limit": "CONT" if cont else remaining, "status": status, "plan_ex": plan,
                      "pallet": info.get("pallet", 0), "wt": w, "mixer": "", "mx_batches": 0, "remark": ""})
    need = [i for i in items if i["plan_ex"] > 0]
    for i in need:
        i["rank"] = sum(1 for j in need if (j["cover"], -j["req"]) < (i["cover"], -i["req"])) + 1
    return items


def mixer_batches(ex, pallet, wt):
    return math.ceil(ex * pallet / wt - 1e-9)


# ---------------- mixer allocation + sequencing ------------------------------
def allocate(items):
    # most urgent first; on equal cover the compound consumed fastest goes first
    need = sorted([i for i in items if i["plan_ex"] > 0], key=lambda i: (i["cover"], -i["req"], i["row"]))
    groups, seen = [], set()
    for i in need:                         # group 22 codes by 21 code (no LC between them)
        if i["c21"] in seen:
            continue
        seen.add(i["c21"])
        groups.append([j for j in need if j["c21"] == i["c21"]])

    load = {1: 0.0, 2: 0.0, 3: 0.0}
    ncomp = {1: 0, 2: 0, 3: 0}
    seq = {1: [], 2: [], 3: []}
    for g in groups:
        cands = [m for m in (1, 2, 3) if all(x["wt"][m] for x in g)]
        best, best_key = None, None
        for m in cands:
            if ncomp[m] >= S["max_compounds"]:
                continue
            lc = 0 if (not seq[m] and S["last_running"][m] == g[0]["c21"]) else S["lc_minutes"]
            need_min = lc + sum(mixer_batches(x["plan_ex"], x["pallet"], x["wt"][m]) for x in g) * S["cycle"][m]
            fits = load[m] + need_min <= S["day_minutes"]
            key = (0 if fits else 1, load[m] + need_min)
            if best_key is None or key < best_key:
                best, best_key = (m, lc), key
        if best is None:
            for x in g:
                x["remark"] = "Not planned - no eligible mixer / compound limit"
            continue
        m, lc = best
        avail = S["day_minutes"] - load[m] - lc
        if avail < S["cycle"][m]:
            for x in g:
                x["remark"] = "Not planned - mixer capacity full"
            continue
        if lc:
            seq[m].append(("LC", g[0]["c21"]))
        load[m] += lc
        ncomp[m] += 1
        for x in g:
            n = mixer_batches(x["plan_ex"], x["pallet"], x["wt"][m])
            avail = S["day_minutes"] - load[m]
            if n * S["cycle"][m] > avail:           # partial: fit what the day allows
                fit_mx = int(avail // S["cycle"][m])
                ex_fit = int(fit_mx * x["wt"][m] // x["pallet"])
                if ex_fit < 1:
                    x["remark"] = "Not planned - mixer capacity full"
                    continue
                x["remark"] = f"Partial: {ex_fit} of {x['plan_ex']} EX batches (capacity)"
                x["plan_ex"] = ex_fit
                n = mixer_batches(ex_fit, x["pallet"], x["wt"][m])
            x["mixer"] = f"MX0{m}"
            x["mx_batches"] = n
            load[m] += n * S["cycle"][m]
            seq[m].append(("PROD", x, n))
    return seq, load


def timeline(seq, m):
    cyc = S["cycle"][m]
    sl = S["shift_minutes"]
    rows, t, sr = [], 0.0, 1

    def shift(tt):
        return SHIFTS[min(int(tt // sl), 2)]

    for idx, s in enumerate(seq):
        if s[0] == "LC":
            nxt = seq[idx + 1][1] if idx + 1 < len(seq) and seq[idx + 1][0] == "PROD" else None
            co = S["default_co"]
            phys = max(0.0, S["lc_minutes"] - (S["co_before"] + S["co_after"]) * cyc)
            for _ in range(S["co_before"]):
                rows.append([shift(t), sr, co, "Cleanout", "", "", "", "", t]); sr += 1; t += cyc
            rows.append([shift(t), sr, "Line Clearance", f"Physical LC before {nxt['mat'] if nxt else ''}", "", "", "", "", t]); sr += 1; t += phys
            for _ in range(S["co_after"]):
                rows.append([shift(t), sr, co, "Cleanout", "", "", "", "", t]); sr += 1; t += cyc
            continue
        _, x, n = s
        rel, ex_prev, left = 1, 0, n
        while left > 0:
            sidx = int(t // sl)
            if sidx >= 2:
                take = left
            else:
                fit = int(((sidx + 1) * sl - t) / cyc + 1e-9)
                if fit <= 0:
                    t = (sidx + 1) * sl
                    continue
                take = min(fit, left)
            to = rel + take - 1
            ex_to = x["plan_ex"] if to == n else int(to * x["plan_ex"] // n)
            ex_from = ex_prev + 1
            rows.append([shift(t), sr, x["c22"], x["mat"], rel, to,
                         ex_from if ex_to >= ex_from else "", ex_to if ex_to >= ex_from else "", t])
            ex_prev = max(ex_prev, ex_to)
            sr += 1; t += take * cyc; rel = to + 1; left -= take
    return rows, t


# ---------------- output -----------------------------------------------------
def write_output(items, seq, path):
    wb = Workbook()
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    hdr_font = Font(name="Arial", bold=True, color="FFFFFF")
    base = Font(name="Arial", size=10)
    fills = {"1-CRITICAL": "F8CBAD", "2-WATCH": "FFE699", "3-OK": "C6EFCE", "LIMIT DONE": "D9D9D9",
             "CHECK DATA": "BDD7EE"}

    ws = wb.active
    ws.title = "Priority"
    cols = ["Rank", "22 Code", "Material", "21 Code", "Colour", "Stock (EX batches)", "Req / day",
            "Days cover", "Limit left", "Status", "Plan EX batches", "MX01", "MX02", "MX03",
            "Assigned mixer", "Mixer batches", "Remark"]
    ws.append(cols)
    order = sorted(items, key=lambda i: (i.get("rank", 999), i["cover"]))
    for i in order:
        ws.append([i.get("rank", ""), int(i["c22"]), i["mat"], i["c21"], i["colour"], i["stock"], i["req"],
                   round(i["cover"], 2) if i["cover"] < 99 else "", i["limit"], i["status"], i["plan_ex"] or "",
                   "Y" if i["wt"][1] else "", "Y" if i["wt"][2] else "", "Y" if i["wt"][3] else "",
                   i["mixer"], i["mx_batches"] or "", i["remark"]])
        f = fills.get(i["status"])
        if f:
            ws.cell(ws.max_row, 10).fill = PatternFill("solid", fgColor=f)
    for m in (1, 2, 3):
        w = wb.create_sheet(f"MX0{m}")
        w.append(["Shift", "Sr no", "22 Code / 29 Code", "Material", "21 code Iter From", "21 code Iter To",
                  "22 code Iter From", "22 code Iter To", "Start (min from shift A)"])
        rows, end = timeline(seq[m], m)
        for r in rows:
            r[-1] = round(r[-1], 1)
            w.append(r)
            if r[2] == "Line Clearance":
                for c in range(1, 10):
                    w.cell(w.max_row, c).fill = PatternFill("solid", fgColor="FFFF00")
        w.append([])
        w.append(["", "", "Planned minutes", round(end, 1), "of", S["day_minutes"],
                  "Cycle min/batch (assumed)", S["cycle"][m]])
    for w in wb.worksheets:
        for row in w.iter_rows():
            for c in row:
                c.font = base
                c.alignment = Alignment(vertical="center")
        for c in w[1]:
            c.fill, c.font = hdr_fill, hdr_font
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for col in w.columns:
            w.column_dimensions[col[0].column_letter].width = max(10, min(32, max(len(str(c.value or "")) for c in col) + 2))
        w.freeze_panes = "A2"
    wb.save(path)


def last_running(read, rawd, plan_serial):
    """21 code of the last compound on each MX sheet, if that plan is older than today."""
    for m in (1, 2, 3):
        if S["last_running"][m]:
            continue
        c = read(f"MX0{m}", 140)
        d = num(c.get((4, 1)))
        if not d or d >= plan_serial:
            continue
        for r in range(140, 3, -1):
            c22 = code(c.get((r, 4)))
            if c22.startswith("22") and len(c22) == 8:
                S["last_running"][m] = rawd.get(c22, {}).get("c21", "")
                break


def main(src, out):
    import datetime
    read = read_workbook(src)
    rawd, wts = load_master(read)
    last_running(read, rawd, (datetime.date.today() - datetime.date(1899, 12, 30)).days)
    print("last running 21 code per mixer:", S["last_running"])
    items = build_priority(read, rawd, wts)
    seq, load = allocate(items)
    write_output(items, seq, out)
    for i in sorted(items, key=lambda i: i["cover"]):
        if i["plan_ex"] or i["status"] in ("1-CRITICAL", "CHECK DATA"):
            print(f'{i["c22"]} {str(i["mat"]):22} c21={i["c21"]:9} cover={i["cover"]:5.2f} {i["status"]:11} '
                  f'plan={i["plan_ex"]:3} -> {i["mixer"]:5} mx={i["mx_batches"]:3} {i["remark"]}')
    for m in (1, 2, 3):
        print(f"MX0{m}: {load[m]:.0f} / {S['day_minutes']} min")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
