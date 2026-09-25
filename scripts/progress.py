"""Summarise build progress into data/progress.json for the progress page (Artifact db doc status/current).

Usage: python progress.py "<fase>" "<wat er nu gebeurt>"
No SolidWorks needed; safe to run any time.
"""
import csv, json, os, sys, time
import rules
from build_parts import excel_rows, load, borrow
from swlib import DATA_DIR, OUT_DIR

LABELS = {
    "din912": "Socket head cap screw DIN 912", "iso10642": "Countersunk socket ISO 10642",
    "iso7380_2": "Button head flange ISO 7380-2", "hexbolt": "Hex bolt", "nut": "Hex nut", "washer": "Washer",
    "setscrew": "Set screw", "carriagebolt": "Carriage bolt", "panhead": "Tapping screw", "splitpen": "Split pin",
    "lock_nut": "Lock nut", "cap_nut": "Cap nut", "flange_nut": "Flange nut", "crown_nut": "Castle nut",
    "knurl_screw": "Knurled screw", "belleville": "Disc spring", "rivet_dome": "Rivet dome head",
    "rivet_cs": "Rivet countersunk", "rivet_nut": "Rivnut", "insert": "Thread insert", "stud": "Clinch stud",
    "clinch_nut": "Clinch nut", "anchor": "Wall plug", "pem_so": "PEM standoff", "pem_bso": "PEM blind standoff",
    "pem_stud": "PEM stud",
}


def main(phase, message):
    rows = excel_rows()
    extra = set()
    f = os.path.join(DATA_DIR, "extra_rows.json")
    if os.path.exists(f):
        extra = {str(r[23]) for r in json.load(open(f, encoding="utf-8"))}
    built = {os.path.splitext(n)[0] for n in os.listdir(OUT_DIR) if n.upper().endswith(".SLDPRT")} if os.path.isdir(OUT_DIR) else set()

    by_kind, by_source, no_rule, geoms, borrowed = {}, {}, [], set(), []
    for code, row in rows.items():
        d = borrow(code, rows, load(code))
        if d.get("borrowed_from"):
            borrowed.append({"code": code, "reason": ("Niet bij Fabory - " if not load(code).get("title") else "Geen maattabel op Fabory - ") + "maten van " + d["borrowed_from"] + ", lengte uit Excel"})
        g = rules.geometry(d, row)
        kind = g[0] if g else ("download" if code in built else "geen regel")
        if g:
            geoms.add((g[0], tuple(sorted(g[1].items()))))
        elif code not in built:
            no_rule.append(code)
        k = by_kind.setdefault(kind, {"kind": kind, "label": LABELS.get(kind, kind), "total": 0, "built": 0})
        src = "Extra variant" if code in extra else (row[24] or "Onbekend")
        s = by_source.setdefault(src, {"source": src, "total": 0, "built": 0})
        for bucket in (k, s):
            bucket["total"] += 1
            bucket["built"] += code in built

    log, fails = {}, {}
    lf = os.path.join(DATA_DIR, "build_log.csv")
    if os.path.exists(lf):
        for r in csv.reader(open(lf, encoding="utf-8"), delimiter=";"):
            if len(r) < 6:
                continue
            if r[2] == "FOUT":
                fails[r[0]] = r[5]
            else:
                log[r[0]] = r
                fails.pop(r[0], None)
    buckets = [("< -30%", -1e9, -30), ("-30..-10%", -30, -10), ("-10..+10%", -10, 10), ("+10..+30%", 10, 30), ("> +30%", 30, 1e9)]
    hist = {b[0]: 0 for b in buckets}
    for r in log.values():
        try:
            dev = float(r[5].rstrip("%"))
        except ValueError:
            continue
        if r[1].startswith("rivet") or r[1] == "anchor":
            continue  # Fabory weight includes mandrel / solid plug model: not comparable
        hist[next(b[0] for b in buckets if b[1] <= dev < b[2])] += 1
    recent = [{"code": r[0], "kind": LABELS.get(r[1], r[1]), "material": r[2], "mass": r[3], "fabory": r[4], "dev": r[5]}
              for r in list(log.values())[-12:]][::-1]

    out = {
        "phase": phase, "message": message, "updated": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "totals": {"codes": len(rows), "built": len(built & set(rows)), "geometries": len(geoms),
                   "failed": len(fails), "noRule": len(no_rule), "extra": len(extra)},
        "byKind": sorted(by_kind.values(), key=lambda x: -x["total"]),
        "bySource": sorted(by_source.values(), key=lambda x: -x["total"]),
        "massDev": [{"bucket": b, "count": hist[b]} for b, *_ in buckets],
        "recent": recent,
        "failures": [{"code": c, "reason": r} for c, r in list(fails.items())[:30]] + borrowed[:30],
        "noRuleCodes": no_rule[:40],
    }
    json.dump(out, open(os.path.join(DATA_DIR, "progress.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    t = out["totals"]
    print(f"{phase}: {t['built']}/{t['codes']} gebouwd, {t['failed']} fout, {t['noRule']} zonder regel")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "Voorbereiding", sys.argv[2] if len(sys.argv) > 2 else "")
