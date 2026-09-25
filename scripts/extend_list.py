"""Create extra Excel rows for all Fabory variants (data/variants/<fam>.json, from variants.py).

- Bolts/screws etc.: every length of the M-sizes already in the Excel (same family).
- Blind rivets (families in RIVET_FAMS): every diameter and length.
Each extra row is a copy of a sibling row of the same family (same size if possible) with Code, Name,
Length (and Size) replaced. Result: data/extra_rows.json = [ [34 cell values], ... ] (same layout as
'Sheet compact'). The original Excel is not touched; fill_excel.py appends these rows to the copy.
"""
import json, os, re
import openpyxl
from swlib import XLSX, DATA_DIR

RIVET_FAMS = {"34110", "34180", "34184"}
RIVNUT_FAMS = {"69045", "69155", "69315"}  # all sizes; description ends with type code, e.g. "6 OCH 55"
RIVNUT = re.compile(r"\b(\d+)\s+(OC\w*H\s+\d+)\s*$")
VAR_DIR = os.path.join(DATA_DIR, "variants")
DIM = re.compile(r"(M?\d+(?:[,.]\d+)?)\s*X\s*(\d+(?:[,.]\d+)?)", re.I)


def fmt(v):
    v = float(str(v).replace(",", "."))
    return int(v) if v == int(v) else v


def main():
    ws = openpyxl.load_workbook(XLSX, read_only=True)["Sheet compact"]
    rows = [list(r) for r in ws.iter_rows(min_row=3, values_only=True) if r[0] and r[23]]
    codes = {str(r[23]) for r in rows}
    by_fam = {}
    for r in rows:
        c = str(r[23])
        if re.fullmatch(r"\d{5}\.\d{3}\.\d{3}", c):
            by_fam.setdefault(c[:5], []).append(r)
    extra, skipped = [], []
    for fam, sibs in sorted(by_fam.items()):
        f = os.path.join(VAR_DIR, f"{fam}.json")
        if not os.path.exists(f):
            continue
        sizes = {str(s[23]).split(".")[1] for s in sibs}
        for v in json.load(open(f, encoding="utf-8")):
            code = v["code"]
            size = code.split(".")[1]
            if code in codes or (fam not in RIVET_FAMS | RIVNUT_FAMS and size not in sizes):
                continue
            if fam in RIVNUT_FAMS:
                m = RIVNUT.search(v["text"])
                if not m:
                    skipped.append(code)
                    continue
                sib = next((s for s in sibs if str(s[23]).split(".")[1] == size), sibs[0])
                row = list(sib)
                row[23], row[25], row[29] = code, f"M{int(size) // 10}", None  # length filled from Fabory json
                row[5] = re.sub(r"\bM\d+\b", f"M{int(size) // 10}", str(sib[5]), count=1) + f" {m.group(2)}"
                extra.append(row)
                continue
            m = None
            for m in DIM.finditer(v["text"]):
                pass  # last "AxB" in the description = size x length
            if not m:
                skipped.append(code)
                continue
            dia, length = m.group(1), fmt(m.group(2))
            sib = next((s for s in sibs if str(s[23]).split(".")[1] == size), sibs[0])
            row = list(sib)
            row[23], row[29] = code, length
            name = str(sib[5])
            if fam in RIVET_FAMS:
                d = fmt(dia.lstrip("Mm"))
                row[25] = sib[25] if str(sib[23]).split(".")[1] == size else None
                name = DIM.sub(f"{d}x{length}", name, count=1)
            else:
                name = re.sub(r"(M\d+(?:[.,]\d+)?)x\d+(?:[.,]\d+)?", rf"\g<1>x{length}", name, count=1, flags=re.I)
            row[5] = name
            extra.append(row)
    json.dump(extra, open(os.path.join(DATA_DIR, "extra_rows.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0, default=str)
    print(f"{len(extra)} extra regels; overgeslagen (geen maat in omschrijving): {len(skipped)} {skipped[:10]}")
    for r in extra[:5] + extra[-5:]:
        print("  ", r[23], "|", r[5], "|", r[25], r[29])


if __name__ == "__main__":
    main()
