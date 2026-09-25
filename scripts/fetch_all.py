"""Fetch Fabory PDF + page data for every Fabory row in the Excel that has no data/<code>.json yet.

Extra mappings (non-Fabory codes sold by Fabory) go in data/fabory_map.json: {"1244": "37980.100.001"}.
"""
import json, os, time
import openpyxl
from fetch_fabory import fetch, DATA_DIR, ROOT
from swlib import XLSX



def main():
    ws = openpyxl.load_workbook(XLSX, read_only=True)["Sheet compact"]
    rows = [r for r in ws.iter_rows(min_row=3, values_only=True) if r[0] and r[23]]
    extra_file = os.path.join(DATA_DIR, "extra_rows.json")
    if os.path.exists(extra_file):
        rows += json.load(open(extra_file, encoding="utf-8"))  # variants from extend_list.py
    map_file = os.path.join(DATA_DIR, "fabory_map.json")
    extra = json.load(open(map_file)) if os.path.exists(map_file) else {}
    todo = [(str(r[23]), extra.get(str(r[23]))) for r in rows if r[24] == "Fabory" or str(r[23]) in extra]
    todo = [t for t in todo if not os.path.exists(os.path.join(DATA_DIR, f"{t[0]}.json"))]
    print(f"{len(todo)} te downloaden", flush=True)
    fails = []
    for i, (code, fab) in enumerate(todo, 1):
        try:
            d = fetch(code, fab)
            if not d["pdf_ok"] or not d.get("title"):
                fails.append(code)
        except Exception as e:
            fails.append(f"{code}: {e}")
        if i % 25 == 0:
            print(f"{i}/{len(todo)}", flush=True)
        time.sleep(0.3)
    print("klaar; mislukt:", fails)


if __name__ == "__main__":
    main()
