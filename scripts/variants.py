"""List all Fabory variants per article family present in the Excel -> data/variants/<family>.json.

A family = first 5 digits of a Fabory code (e.g. 53570 = DIN 912 A4). Output per variant:
{"code": "53570.100.016", "text": "<variant description>"}.
Prints per family: variants on Fabory, already in Excel, new for sizes already in Excel.
"""
import html, json, os, re, subprocess, time
import openpyxl
from swlib import XLSX, DATA_DIR

OUT = os.path.join(DATA_DIR, "variants")
CODE = re.compile(r"\b(\d{5}\.\d{3}\.\d{3})\b")


def get(url):
    r = subprocess.run(["curl.exe", "-s", "-L", "-A", "Mozilla/5.0", url], capture_output=True)
    return r.stdout.decode("utf-8", "replace")


def family_variants(fam):
    seen, page = {}, 0
    while True:
        h = get(f"https://www.fabory.com/nl/p/{fam}/variants?pageSize=100&page={page}")
        t = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", h)))
        new = 0
        for m in CODE.finditer(t):
            c = m.group(1)
            if c.startswith(fam) and c not in seen:
                seen[c] = t[m.end(): m.end() + 160].split(" Ga naar")[0].strip()
                new += 1
        if new == 0 or page > 30:
            break
        page += 1
        time.sleep(0.3)
    return [{"code": c, "text": s} for c, s in seen.items()]


def main():
    os.makedirs(OUT, exist_ok=True)
    ws = openpyxl.load_workbook(XLSX, read_only=True)["Sheet compact"]
    rows = [r for r in ws.iter_rows(min_row=3, values_only=True) if r[0] and r[23]]
    codes = {str(r[23]) for r in rows}
    fams = sorted({c[:5] for c in codes if CODE.fullmatch(c)})
    total_new = 0
    for fam in fams:
        f = os.path.join(OUT, f"{fam}.json")
        if os.path.exists(f):
            v = json.load(open(f, encoding="utf-8"))
        else:
            v = family_variants(fam)
            json.dump(v, open(f, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        sizes = {c.split(".")[1] for c in codes if c.startswith(fam + ".")}
        new = [x for x in v if x["code"] not in codes and x["code"].split(".")[1] in sizes]
        total_new += len(new)
        print(f"{fam}: fabory {len(v):4}  excel {sum(c.startswith(fam + '.') for c in codes):3}  nieuw(zelfde M) {len(new):4}  {v[0]['text'][:60] if v else ''}")
    print("TOTAAL nieuw:", total_new)


if __name__ == "__main__":
    main()
