"""Download Fabory datasheet PDF + parse product page tables into JSON.

Usage: python fetch_fabory.py <excel_code> [<fabory_code>] ...
  excel_code  = code as in Excel (used for filenames)
  fabory_code = optional Fabory article nr when it differs (e.g. Nord-Lock 1244 -> 37980.100.001)
"""
import html, json, os, re, subprocess, sys

ROOT = os.path.abspath(os.environ.get("FASTENER_PROJECT", os.getcwd()))  # project folder (config.json, data, SLDPRT, PDF)
PDF_DIR = os.path.join(ROOT, "PDF")
DATA_DIR = os.path.join(ROOT, "data")
UA = "Mozilla/5.0"


def curl(url, out):
    r = subprocess.run(["curl.exe", "-s", "-L", "-A", UA, "-o", out, "-w", "%{http_code}", url],
                       capture_output=True, text=True)
    return r.stdout.strip()


def clean(s):
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", html.unescape(s)).strip()


def parse_page(h):
    data = {}
    m = re.search(r"<h1[^>]*>(.*?)</h1>", h, re.S)
    data["title"] = clean(m.group(1)) if m else None
    m = re.search(r"Verpakt per\s*(\d+)", h)
    data["moq"] = int(m.group(1)) if m else None
    # Split page into sections by <h2>; collect th/td pairs per section
    for sec in re.split(r"<h2", h)[1:]:
        name = clean(sec.split("</h2>")[0].split(">", 1)[-1])
        name = name.split(" ")[0] if name else name
        pairs = re.findall(r"<tr[^>]*>\s*<th[^>]*>(.*?)</th>\s*<td[^>]*>(.*?)</td>", sec, re.S)
        if pairs and name:
            data.setdefault(name, {}).update({clean(k): clean(v) for k, v in pairs})
    return data


def fetch(excel_code, fabory_code=None):
    fab = (fabory_code or excel_code).replace(".", "")
    os.makedirs(PDF_DIR, exist_ok=True)
    os.makedirs(DATA_DIR, exist_ok=True)
    pdf = os.path.join(PDF_DIR, f"{excel_code}.pdf")
    st_pdf = curl(f"https://www.fabory.com/nl/p/{fab}/generatePdf", pdf)
    tmp = os.path.join(DATA_DIR, f"{excel_code}.html")
    st_html = curl(f"https://www.fabory.com/nl/p/{fab}", tmp)
    with open(tmp, encoding="utf-8", errors="replace") as f:
        data = parse_page(f.read())
    os.remove(tmp)
    ok_pdf = st_pdf == "200" and open(pdf, "rb").read(4) == b"%PDF"
    if not ok_pdf and os.path.exists(pdf):
        os.remove(pdf)
    data.update(excel_code=excel_code, fabory_code=fabory_code or excel_code,
                pdf_ok=ok_pdf, html_status=st_html)
    with open(os.path.join(DATA_DIR, f"{excel_code}.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    return data


if __name__ == "__main__":
    args = sys.argv[1:]
    for a in args:
        ex, _, fab = a.partition("=")
        d = fetch(ex, fab or None)
        print(ex, "pdf" if d["pdf_ok"] else "NO-PDF", d.get("title"))
        print("   ", d.get("Afmetingen"), "| gewicht:", (d.get("Specificaties") or {}).get("Gewicht per stuk (g)"))
