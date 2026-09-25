"""Build SolidWorks parts for Excel rows (one model per unique geometry, saved per code + material).

Usage:
  python build_parts.py <code> ...          build these codes (overwrites)
  python build_parts.py --all               build every row without SLDPRT/<code>.SLDPRT yet
  python build_parts.py --norm "DIN 933"    build rows with this Norm (Excel col AC) not yet built
  python build_parts.py --check <code> ...  only show geometry/params, no SolidWorks

Result lines are appended to data/build_log.csv.
"""
import csv, json, os, sys, traceback
import openpyxl
import rules
from swlib import Builder, sw_lock, MATERIALS, DATA_DIR, OUT_DIR, XLSX


def excel_rows():
    """Excel rows + extra variant rows (data/extra_rows.json from extend_list.py), keyed by code."""
    ws = openpyxl.load_workbook(XLSX, read_only=True)["Sheet compact"]
    rows = {str(r[23]): r for r in ws.iter_rows(min_row=3, values_only=True) if r[0] and r[23]}
    f = os.path.join(DATA_DIR, "extra_rows.json")
    if os.path.exists(f):
        for r in json.load(open(f, encoding="utf-8")):
            rows.setdefault(str(r[23]), tuple(r))
    return rows


def load(code):
    f = os.path.join(DATA_DIR, f"{code}.json")
    return json.load(open(f, encoding="utf-8")) if os.path.exists(f) else {}


def borrow(code, rows, d):
    """No Fabory dimension table (code not sold or page without table): borrow head dimensions from a
    sibling with the same norm + size (same family first) and take the length from the Excel row.
    Marked with d['borrowed_from'] so reports can flag it."""
    row = rows[code]
    if row[24] not in ("Fabory", "Nord-Lock", "fischer"):
        return d  # PEM / Elesa use their own data or downloaded CAD
    if d.get("title") and (d.get("Afmetingen") or (row[24] == "fischer" and (d.get("Specificaties") or {}).get("Diameter"))):
        return d  # own data present (fischer plugs keep dims in Specificaties)
    fam = code.split(".")[0]
    cands = [c for c, r in rows.items() if c != code and r[28] == row[28] and r[25] == row[25] and r[21] == row[21]]
    cands.sort(key=lambda c: (c.split(".")[0] != fam, c.split(".")[:2] != code.split(".")[:2]))
    for c in cands:
        s = load(c)
        if s.get("Afmetingen") and s.get("title"):
            b = json.loads(json.dumps(s))
            if row[29] is not None:
                for k in ("L (mm)",):
                    if k in b["Afmetingen"]:
                        b["Afmetingen"][k] = str(row[29])
            b.update(borrowed_from=c, excel_code=code)
            b.setdefault("Specificaties", {}).pop("Gewicht per stuk (g)", None)
            return b
    return d


def plan(codes, rows):
    groups, missing = {}, []
    for c in codes:
        d = borrow(c, rows, load(c))
        g = rules.geometry(d, rows[c])
        if not g:
            missing.append(c)
            continue
        kind, p = g
        wt = (d.get("Specificaties") or {}).get("Gewicht per stuk (g)")
        try:
            wt = float(str(wt).replace(",", "."))
        except (TypeError, ValueError):
            wt = None
        groups.setdefault((kind, tuple(sorted(p.items()))), []).append((c, rows[c][22], wt))
    return groups, missing


def keep_awake():
    """Ask Windows not to sleep while this process runs (released automatically on exit)."""
    try:
        import ctypes
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001)  # ES_CONTINUOUS | ES_SYSTEM_REQUIRED
    except Exception:
        pass


def main(argv):
    keep_awake()
    os.makedirs(OUT_DIR, exist_ok=True)
    rows = excel_rows()
    check = "--check" in argv
    argv = [a for a in argv if a != "--check"]
    if argv and argv[0] == "--all":
        codes = [c for c in rows if not os.path.exists(os.path.join(OUT_DIR, f"{c}.SLDPRT"))]
    elif argv and argv[0] == "--norm":
        codes = [c for c, r in rows.items() if r[28] == argv[1]
                 and not os.path.exists(os.path.join(OUT_DIR, f"{c}.SLDPRT"))]
    else:
        codes = argv
    groups, missing = plan(codes, rows)
    print(f"{len(codes)} codes -> {len(groups)} geometrieen; geen regel: {len(missing)}")
    if missing:
        print("  geen regel voor:", ", ".join(missing[:30]), "..." if len(missing) > 30 else "")
    if check:
        for (kind, p), members in groups.items():
            print(f"  {kind:12} {dict(p)}  <- {[m[0] for m in members]}")
        return
    log = open(os.path.join(DATA_DIR, "build_log.csv"), "a", newline="", encoding="utf-8")
    wr = csv.writer(log, delimiter=";")
    from snapshot import snap
    b = Builder()
    n_built = 0
    for (kind, p), members in groups.items():
        with sw_lock():
            try:
                b.new()
                rules.builder(kind)(b, dict(p))
                for code, mat, wt in members:
                    swmat = MATERIALS[mat]
                    path, err, mass = b.save_as(code, swmat)
                    dev = f"{(mass - wt) / wt * 100:+.1f}%" if wt else "-"
                    print(f"  {code:16} {kind:12} {swmat[:30]:30} SW {mass:8.2f} g  Fabory {wt} g  {dev}  err={err}", flush=True)
                    wr.writerow([code, kind, swmat, f"{mass:.2f}", wt, dev, err])
                b.close()
                n_built += 1
                if n_built % 40 == 1:  # snapshot for the progress page (data/snapshots)
                    try:
                        snap(b.sw, members[-1][0])
                    except Exception as e:
                        print("  snapshot mislukt:", e)
            except Exception as e:
                print(f"  FOUT {kind} {[m[0] for m in members]}: {e}", flush=True)
                for code, *_ in members:
                    wr.writerow([code, kind, "FOUT", "", "", str(e)[:120], ""])
                traceback.print_exc()
                try:
                    b.close()
                except Exception:
                    pass
        log.flush()


if __name__ == "__main__":
    main(sys.argv[1:])
