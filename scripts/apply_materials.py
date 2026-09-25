"""Assign Excel material to downloaded (not generated) parts, e.g. Elesa-Ganter.

Reads data/elesa.json {"<code>": "<Excel material>"}; opens SLDPRT/<code>.SLDPRT, sets material if missing
(or different), saves, logs mass to data/build_log.csv.
"""
import csv, json, os
from swlib import connect, sw_lock, MATERIALS, DB, DATA_DIR, OUT_DIR


def main():
    f = os.path.join(DATA_DIR, "elesa.json")
    if not os.path.exists(f):
        print("geen data/elesa.json")
        return
    todo = json.load(open(f, encoding="utf-8"))
    sw = connect()
    log = csv.writer(open(os.path.join(DATA_DIR, "build_log.csv"), "a", newline="", encoding="utf-8"), delimiter=";")
    for code, mat in todo.items():
        path = os.path.join(OUT_DIR, f"{code}.SLDPRT")
        if not os.path.exists(path):
            print("ontbreekt:", code)
            continue
        with sw_lock():
            doc = sw.OpenDoc(path, 1)
            if doc is None:
                print("kan niet openen:", code)
                continue
            cfg = doc.ConfigurationManager.ActiveConfiguration.Name
            swmat = MATERIALS.get(mat)
            if swmat:
                doc.SetMaterialPropertyName2(cfg, DB, swmat)
                doc.ForceRebuild3(False)
                doc.SaveAs3(path, 0, 1)  # silent, same path
            mass = doc.Extension.CreateMassProperty.Mass * 1000
            print(f"{code:24} {swmat} {mass:.2f} g")
            log.writerow([code, "download", swmat, f"{mass:.2f}", "", "", ""])
            sw.CloseDoc(doc.GetTitle)


if __name__ == "__main__":
    main()
