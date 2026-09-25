"""Explain codes without a rule: is the code in Fabory's variant list, and what are the neighbours?"""
import json, os
import rules
from build_parts import excel_rows, load
from swlib import DATA_DIR

rows = excel_rows()
for code, row in rows.items():
    d = load(code)
    if rules.geometry(d, row) or os.path.exists(os.path.join(DATA_DIR, "..", "SLDPRT", f"{code}.SLDPRT")):
        continue
    fam, size = code.split(".")[0], (code.split(".") + [""])[1]
    vf = os.path.join(DATA_DIR, "variants", f"{fam}.json")
    v = json.load(open(vf, encoding="utf-8")) if os.path.exists(vf) else []
    same = [x["code"] for x in v if x["code"].split(".")[1] == size]
    print(f"{code:16} {str(row[5])[:42]:42} json={'ja' if d else 'nee'} title={bool(d.get('title'))} "
          f"afm={bool(d.get('Afmetingen'))} in_variantlijst={code in {x['code'] for x in v}} zelfde_maat={same[:6]}")
