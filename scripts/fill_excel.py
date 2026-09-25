"""Vult de kopie van de Import-Excel aan met Fabory-data.

Maakt een KOPIE van de importlijst (config.json "excel")
(het origineel wordt nooit aangeraakt) genaamd
<excel>_aangevuld.xlsx en:

1. Voegt in tab 'Sheet compact' direct na kolom 'Length (mm)' vier kolommen in:
   'Preferred Supplier', 'Supplier Article Nr', 'MOQ', 'Clamping range (mm)'
   (met correcte opmaak, kolombreedtes en meeschuivende merged cells in rij 1).
   Clamping range wordt alleen gevuld voor popnagels/blindklinkmoeren (json
   Afmetingen-sleutel 'Klembereik (min. - max.)').
2. Vult alleen LEGE cellen aan in Size (M), Strength Class, Thread Pitch, Norm,
   Length (mm) met data uit data\\<code>.json (Fabory). Aangevulde cellen worden
   geel gemarkeerd.
3. Bestaande, van Fabory afwijkende waarden worden NIET gewijzigd maar oranje
   gemarkeerd + celopmerking met de Fabory-waarde. Voor Norm tellen verschillen
   in alleen spaties/streepjes/hoofdletters (en een Excel-norm die een prefix
   is van de Fabory-norm, bv. 'DIN 2093' vs 'DIN 2093-A/C') niet als afwijking.
4. Voegt extra rijen uit data\\extra_rows.json (gemaakt door extend_list.py,
   34 kolommen in de oude lay-out) onderaan 'Sheet compact' toe, past er
   dezelfde aanvulling op toe, en markeert de hele rij lichtblauw.
5. Twee kolommen direct na 'Clamping range (mm)':
   - 'Drill hole (mm)': alleen voor popnagels (34110/34180/34184),
     blindklinkmoeren (69045/69155/69315), fischer-pluggen, draadbussen
     (71752) en PEM-parts (data\\pem\\<code>.json).
   - 'Drive type': alleen 'Inbus'/'Torx'/'Zeskant'/'Kruiskop', afgeleid uit
     Specificaties 'Aandraaivoorziening'/'Aandrijving'/'Kopsoort'/titel (of
     voor Elesa+Ganter uit de Norm-kolom). Leeg voor ringen, popnagels,
     pluggen, inserts, PEM, splitpennen.
6. Voegt een tabblad 'Aanvulling log' toe met alle mutaties (incl. 1 regel per
   toegevoegde variant-rij, actie 'nieuwe variant').

Idempotent: opnieuw draaien maakt gewoon een nieuwe kopie en overschrijft die.
"""
import copy
import json
import os
import re
import shutil

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

ROOT = os.path.abspath(os.environ.get("FASTENER_PROJECT", os.getcwd()))  # project folder (config.json, data, SLDPRT, PDF)
from swlib import XLSX as SRC
DST = os.path.splitext(SRC)[0] + "_aangevuld.xlsx"
DATA_DIR = os.path.join(ROOT, "data")
FABORY_MAP_PATH = os.path.join(DATA_DIR, "fabory_map.json")
EXTRA_ROWS_PATH = os.path.join(DATA_DIR, "extra_rows.json")
PEM_DIR = os.path.join(DATA_DIR, "pem")

SHEET = "Sheet compact"
LOG_SHEET = "Fill log"

# 1-based column indices in 'Sheet compact' (rij 2 = kop), in de layout
# VOOR onze eigen kolom-invoegingen (= layout van extra_rows.json, 34 kolommen)
N_ORIG_COLS = 34
COL_MANUF_CODE = 24  # X
COL_MANUFACTURER = 25  # Y
COL_SIZE = 26  # Z
COL_STRENGTH = 27  # AA
COL_PITCH = 28  # AB
COL_NORM = 29  # AC
COL_LENGTH = 30  # AD
INSERT_AT = COL_LENGTH + 1  # 31 = AE, direct na Length (mm)
N_NEW = 6
COL_PREF_SUPPLIER = INSERT_AT
COL_SUPPLIER_ARTNR = INSERT_AT + 1
COL_MOQ = INSERT_AT + 2
COL_CLAMPING = INSERT_AT + 3
COL_DRILL_HOLE = INSERT_AT + 4
COL_DRIVE_TYPE = INSERT_AT + 5
N_TOTAL_COLS = N_ORIG_COLS + N_NEW

RIVET_NUT_FAMS = {"34110", "34180", "34184", "69045", "69155", "69315"}
INSERT_FAM = "71752"
DRIVE_TYPE_VALUES = {"Hex socket", "Torx", "Hex", "Phillips"}

YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
ORANGE = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
LIGHT_BLUE = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

NORM_PRIORITY = ["DIN", "ISO", "NEN", "GN"]


# --------------------------------------------------------------------------
# JSON helpers
# --------------------------------------------------------------------------

def load_fabory_map():
    if os.path.exists(FABORY_MAP_PATH):
        try:
            with open(FABORY_MAP_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def load_json_for(code, fabory_map):
    """Zoek data\\<code>.json; val terug op fabory_map.json voor codes die niet
    direct als Fabory-artikelnummer bestaan (bv. Nord-Lock / fischer)."""
    candidates = [code]
    mapped = fabory_map.get(code)
    if mapped and mapped not in candidates:
        candidates.append(mapped)
    for c in candidates:
        p = os.path.join(DATA_DIR, f"{c}.json")
        if os.path.exists(p):
            try:
                with open(p, encoding="utf-8") as f:
                    data = json.load(f)
                if data.get("pdf_ok") and data.get("title"):
                    return data
            except Exception:
                continue
    return None


def load_pem_json(code):
    """PEM-parts hebben een handmatig samengestelde json in data\\pem\\<code>.json
    (zelfde bestandsnaam als de Manufacturer Code, geen punten-notatie)."""
    p = os.path.join(PEM_DIR, f"{code}.json")
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None


# --------------------------------------------------------------------------
# Field extraction from Fabory json
# --------------------------------------------------------------------------

def clean_norm_value(v):
    if v is None:
        return None
    v = str(v).strip()
    v = v.lstrip("≈").strip()  # strip leading '≈'
    return v


def fabory_size(data):
    afm = data.get("Afmetingen") or {}
    spec = data.get("Specificaties") or {}
    d = afm.get("d-D")
    if d:
        d = d.strip()
        if re.match(r"^M\d+(\.\d+)?$", d):
            return d
    d2 = spec.get("Diameter")
    if d2 and re.match(r"^M\d+(\.\d+)?$", d2.strip()):
        return d2.strip()
    return None


def fabory_pitch(data):
    afm = data.get("Afmetingen") or {}
    p = afm.get("P")
    if p is None:
        return None
    p = str(p).strip()
    if not p:
        return None
    return p


def fabory_length(data):
    afm = data.get("Afmetingen") or {}
    l = afm.get("L (mm)")
    if l is None:
        return None
    try:
        f = float(str(l).replace(",", "."))
    except ValueError:
        return None
    return int(f) if f == int(f) else f


def fabory_strength(data):
    spec = data.get("Specificaties") or {}
    klasse = spec.get("Klasse") or spec.get("Sterkteklasse")
    tech = spec.get("Materiaalsoort Technisch")
    if klasse:
        klasse = re.sub(r"\s*\(.*?\)\s*", "", str(klasse)).strip()
        klasse = klasse.strip("|").strip()  # strip stray '|' markup artifacts, e.g. '|8|'
        if not klasse:
            klasse = None
    if tech and klasse:
        return f"{tech}-{klasse}"
    if klasse:
        return klasse
    # No 'Klasse'/'Sterkteklasse' at all (e.g. washers/rings): Materiaalsoort
    # Technisch alone (e.g. '254 SMO') is a MATERIAL grade, not a strength
    # class, so don't fall back to it here.
    return None


def fabory_norm(data, preferred_system=None):
    norms = data.get("Normeringen") or {}
    if not norms:
        return None
    order = []
    if preferred_system and preferred_system in norms:
        order.append(preferred_system)
    for sys_ in NORM_PRIORITY:
        if sys_ not in order and sys_ in norms:
            order.append(sys_)
    for sys_ in norms:
        if sys_ not in order:
            order.append(sys_)
    for sys_ in order:
        val = clean_norm_value(norms.get(sys_))
        if val:
            return f"{sys_} {val}"
    return None


def norm_system_of(norm_str):
    if not norm_str:
        return None
    m = re.match(r"^([A-Za-z]+)\s", str(norm_str).strip())
    return m.group(1) if m else None


def _fmt_num(x):
    f = float(str(x).replace(",", "."))
    return str(int(f)) if f == int(f) else str(f)


CLAMPING_KEY = "Klembereik (min. - max.)"


def fabory_clamping_range(data):
    """Klembereik voor popnagels/blindklinkmoeren, bv. '12.50 - 16.50' -> '12.5 - 16.5'.
    Alleen aanwezig voor deze productfamilies, dus dit is meteen ook de check
    of een regel een popnagel/blindklinkmoer is."""
    afm = data.get("Afmetingen") or {}
    v = afm.get(CLAMPING_KEY)
    if not v:
        return None
    m = re.match(r"^\s*([\d.,]+)\s*-\s*([\d.,]+)\s*$", str(v))
    if not m:
        return None
    try:
        return f"{_fmt_num(m.group(1))} - {_fmt_num(m.group(2))}"
    except ValueError:
        return None


def norm_key(s):
    """Normaliseer een normcode voor vergelijking: spaties/streepjes weg, hoofdletters."""
    return re.sub(r"[\s\-]+", "", str(s)).upper()


def norms_equal(existing, fabory_val):
    """Verschillen in alleen spaties/streepjes/hoofdletters tellen niet als
    afwijking. Ook niet als de Excel-norm een prefix is van de Fabory-norm
    (bv. 'DIN 2093' vs 'DIN 2093-A/C') of andersom."""
    e = norm_key(existing)
    f = norm_key(fabory_val)
    if not e or not f:
        return e == f
    return e == f or e.startswith(f) or f.startswith(e)


def _fmt_maybe_num(v):
    s = str(v).strip()
    try:
        return _fmt_num(s)
    except ValueError:
        return s


def compute_drill_hole(code, manufacturer, data, pem_data):
    """'Drill hole (mm)': alleen voor popnagels/blindklinkmoeren, fischer-
    pluggen, draadbussen (71752) en PEM-parts. Alle andere regels: None."""
    prefix = code.split(".")[0] if "." in code else code

    if prefix in RIVET_NUT_FAMS or prefix == INSERT_FAM:
        if not data:
            return None
        afm = data.get("Afmetingen") or {}
        v = afm.get("Boor ø")  # 'Boor ø'
        if v:
            return _fmt_maybe_num(v)
        if prefix == INSERT_FAM:
            # draadbus/insert: geen aparte 'Boor ø', d2 = buitendiameter huls
            # (= benodigd boorgat voor de zelftappende bus)
            v = afm.get("d2")
            if v:
                return _fmt_maybe_num(v)
        return None

    if manufacturer == "fischer":
        if not data:
            return None
        spec = data.get("Specificaties") or {}
        v = spec.get("Boorgatdiameter (mm)")
        if not v:
            # geen aparte boorgat-sleutel: boorgat = plugdiameter
            v = spec.get("Diameter")
        if not v:
            return None
        v = re.sub(r"\s*mm\s*$", "", str(v)).strip()
        return _fmt_maybe_num(v)

    if manufacturer == "PennEngineering (PEM)":
        if not pem_data:
            return None
        afm = pem_data.get("Afmetingen") or {}
        for k in ("Hole Size", "Mounting hole", "Sheet hole", "Panel hole",
                  "Hole dia", "Hole", "D (hole)", "Hole diameter"):
            v = afm.get(k) or pem_data.get(k)
            if v:
                return _fmt_maybe_num(v)
        return None

    return None


DRIVE_KEY_CANDIDATES = ("Aandraaivoorziening", "Aandrijving")


def fabory_drive_type(data):
    """Alleen 'Inbus'/'Torx'/'Zeskant'/'Kruiskop', zonder maat. Bron:
    Specificaties Aandraaivoorziening/Aandrijving/Kopsoort, met titel als
    laatste terugval. Geeft None voor ringen/popnagels/pluggen/inserts/PEM/
    splitpennen (die hebben geen van deze sleutels)."""
    spec = data.get("Specificaties") or {}
    parts = []
    for k in DRIVE_KEY_CANDIDATES:
        v = spec.get(k)
        if v:
            parts.append(str(v))
    kop = spec.get("Kopsoort")
    if kop:
        parts.append(str(kop))
    parts.append(data.get("title") or "")
    text = " ".join(parts).lower()

    if "binnenzeskant" in text:
        return "Hex socket"
    if "t-ster" in text or "torx" in text:
        return "Torx"
    if "phillips" in text or "pozidriv" in text or "kruisgleuf" in text or "kruiskop" in text:
        return "Phillips"
    if "zeskant" in text:
        return "Hex"
    return None


def elesa_drive_type(existing_norm):
    """Elesa+Ganter heeft geen Fabory-json; drive type volgt uit de
    (al aanwezige) Norm-kolom: ISO 7379 -> Inbus, DIN 466/464 (kartel) -> leeg,
    GN 732.1 -> Inbus."""
    if not existing_norm:
        return None
    n = norm_key(existing_norm)
    if n.startswith("ISO7379"):
        return "Hex socket"
    if n.startswith("DIN466") or n.startswith("DIN464"):
        return None
    if n.startswith("GN732.1") or n.startswith("GN7321"):
        return "Hex socket"
    return None


# --------------------------------------------------------------------------
# Excel structure repair after insert_cols
# --------------------------------------------------------------------------

def insert_supplier_columns(ws):
    # capture BEFORE insert
    orig_merges = [
        (mc.min_row, mc.min_col, mc.max_row, mc.max_col) for mc in list(ws.merged_cells.ranges)
    ]
    orig_widths = {}
    for letter, dim in ws.column_dimensions.items():
        if dim.width is not None:
            orig_widths[letter] = dim.width
    orig_col_idx_widths = {}
    from openpyxl.utils import column_index_from_string
    for letter, w in orig_widths.items():
        try:
            idx = column_index_from_string(letter)
        except ValueError:
            continue
        orig_col_idx_widths[idx] = w

    header_style_src = ws.cell(row=2, column=COL_LENGTH)
    row1_style_src = ws.cell(row=1, column=COL_LENGTH)

    ws.insert_cols(INSERT_AT, N_NEW)

    # --- fix merged ranges: shift ranges that started at/after INSERT_AT ---
    # insert_cols() does not keep the merged-cell bookkeeping consistent with
    # the shifted grid, so drop the (now stale) internal range set directly
    # instead of going through unmerge_cells() (which tries to touch cells
    # that insert_cols already moved/removed).
    ws.merged_cells.ranges = set()
    for (min_row, min_col, max_row, max_col) in orig_merges:
        if min_col >= INSERT_AT:
            min_col += N_NEW
            max_col += N_NEW
        ws.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)

    # --- fix column widths: shift widths for columns >= INSERT_AT by +3 ---
    for idx in sorted((i for i in orig_col_idx_widths if i >= INSERT_AT), reverse=True):
        new_letter = get_column_letter(idx + N_NEW)
        ws.column_dimensions[new_letter].width = orig_col_idx_widths[idx]

    # --- headers + styles for the new columns ---
    headers = ["Preferred Supplier", "Supplier Article Nr", "MOQ", "Clamping range (mm)",
               "Drill hole (mm)", "Drive type"]
    widths = [20, 22, 10, 20, 16, 14]
    for i, (h, w) in enumerate(zip(headers, widths)):
        col = INSERT_AT + i
        letter = get_column_letter(col)
        c2 = ws.cell(row=2, column=col, value=h)
        c2.font = copy.copy(header_style_src.font)
        c2.fill = copy.copy(header_style_src.fill)
        c2.border = copy.copy(header_style_src.border)
        c2.alignment = copy.copy(header_style_src.alignment)
        c1 = ws.cell(row=1, column=col)
        c1.font = copy.copy(row1_style_src.font)
        c1.fill = copy.copy(row1_style_src.fill)
        c1.border = copy.copy(row1_style_src.border)
        c1.alignment = copy.copy(row1_style_src.alignment)
        ws.column_dimensions[letter].width = w


# --------------------------------------------------------------------------
# Extra (variant) rows from data\extra_rows.json
# --------------------------------------------------------------------------

def copy_row_style(ws, src_row, dst_row, ncols):
    for c in range(1, ncols + 1):
        src = ws.cell(row=src_row, column=c)
        dst = ws.cell(row=dst_row, column=c)
        dst.font = copy.copy(src.font)
        dst.fill = copy.copy(src.fill)
        dst.border = copy.copy(src.border)
        dst.alignment = copy.copy(src.alignment)
        dst.number_format = src.number_format


def append_extra_rows(ws):
    """Voegt data\\extra_rows.json (34-koloms rijen, layout vóór onze eigen
    kolom-invoegingen) onderaan de sheet toe, met de celstijl van de rij
    erboven. Geeft de set rijnummers van de toegevoegde variant-rijen terug."""
    if not os.path.exists(EXTRA_ROWS_PATH):
        return set()
    try:
        with open(EXTRA_ROWS_PATH, encoding="utf-8") as f:
            extra = json.load(f)
    except Exception:
        return set()
    if not extra:
        return set()

    style_src_row = ws.max_row
    added_rows = set()
    for row_vals in extra:
        new_row = ws.max_row + 1
        copy_row_style(ws, style_src_row, new_row, N_TOTAL_COLS)
        for i, v in enumerate(row_vals[:N_ORIG_COLS], start=1):
            ws.cell(row=new_row, column=i, value=v)
        added_rows.add(new_row)
    return added_rows


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def norm_str(v):
    if v is None:
        return ""
    return str(v).strip()


def values_equal(existing, fabory_val, numeric=False, is_norm=False):
    if is_norm:
        return norms_equal(existing, fabory_val)
    if numeric:
        try:
            return float(str(existing).replace(",", ".")) == float(str(fabory_val).replace(",", "."))
        except (ValueError, TypeError):
            pass
    return norm_str(existing) == norm_str(fabory_val)


def main():
    shutil.copy2(SRC, DST)
    fabory_map = load_fabory_map()

    wb = openpyxl.load_workbook(DST)  # not read_only -> keep styles/merges/tabs
    ws = wb[SHEET]

    insert_supplier_columns(ws)

    extra_rows = append_extra_rows(ws)

    # remove any stale log sheet from a previous run (idempotent)
    if LOG_SHEET in wb.sheetnames:
        del wb[LOG_SHEET]
    log_ws = wb.create_sheet(LOG_SHEET)
    log_ws.append(["Code", "Column", "Old", "New/Fabory", "Action"])
    for c in log_ws[1]:
        f = copy.copy(c.font)
        f.bold = True
        c.font = f

    # --- pass 1: learn the Norm "system" (DIN/ISO/...) already used per code-prefix ---
    from collections import Counter
    prefix_norm_system = {}
    prefix_counter = {}
    max_row = ws.max_row
    for r in range(3, max_row + 1):
        code = ws.cell(row=r, column=COL_MANUF_CODE).value
        if not code:
            continue
        norm_val = ws.cell(row=r, column=COL_NORM).value
        if not norm_val:
            continue
        prefix = str(code).split(".")[0]
        sys_ = norm_system_of(norm_val)
        if sys_:
            prefix_counter.setdefault(prefix, Counter())[sys_] += 1
    for prefix, cnt in prefix_counter.items():
        prefix_norm_system[prefix] = cnt.most_common(1)[0][0]

    stats = Counter()
    log_rows = []

    for r in range(3, max_row + 1):
        code = ws.cell(row=r, column=COL_MANUF_CODE).value
        if not code:
            continue
        code = str(code)
        manufacturer = ws.cell(row=r, column=COL_MANUFACTURER).value
        data = load_json_for(code, fabory_map)
        pem_data = load_pem_json(code) if manufacturer == "PennEngineering (PEM)" else None

        # --- Drill hole (mm): popnagels/blindklinkmoeren/fischer-pluggen/
        # draadbussen/PEM only ---
        drill = compute_drill_hole(code, manufacturer, data, pem_data)
        if drill is not None:
            ws.cell(row=r, column=COL_DRILL_HOLE, value=drill)
            stats["drill_hole"] += 1

        # --- Drive type ---
        if manufacturer == "Elesa+Ganter":
            drive = elesa_drive_type(ws.cell(row=r, column=COL_NORM).value)
        elif data is not None:
            drive = fabory_drive_type(data)
        else:
            drive = None
        if drive:
            ws.cell(row=r, column=COL_DRIVE_TYPE, value=drive)
            stats["drive_type"] += 1

        # --- Preferred Supplier / Supplier Article Nr / MOQ ---
        pref_supplier = None
        if manufacturer == "Elesa+Ganter":
            pref_supplier = "Elesa-Ganter"
        elif data is not None:
            pref_supplier = "Fabory"

        if pref_supplier:
            ws.cell(row=r, column=COL_PREF_SUPPLIER, value=pref_supplier)
            stats["pref_supplier"] += 1

        if data is not None:
            art_nr = (data.get("Artikelnummers") or {}).get("Artikelnummer") or data.get("fabory_code")
            if art_nr:
                ws.cell(row=r, column=COL_SUPPLIER_ARTNR, value=art_nr)
                stats["supplier_artnr"] += 1
            moq = data.get("moq")
            if moq is not None:
                ws.cell(row=r, column=COL_MOQ, value=moq)
                stats["moq"] += 1

            clamp = fabory_clamping_range(data)
            if clamp is not None:
                ws.cell(row=r, column=COL_CLAMPING, value=clamp)
                stats["clamping"] += 1

            prefix = code.split(".")[0]
            preferred_system = prefix_norm_system.get(prefix)

            fields = [
                (COL_SIZE, "Size (M)", fabory_size(data), False, False),
                (COL_STRENGTH, "Strength Class", fabory_strength(data), False, False),
                (COL_PITCH, "Thread Pitch", fabory_pitch(data), True, False),
                (COL_NORM, "Norm", fabory_norm(data, preferred_system), False, True),
                (COL_LENGTH, "Length (mm)", fabory_length(data), True, False),
            ]

            for col, label, fval, numeric, is_norm in fields:
                if fval is None:
                    continue
                cell = ws.cell(row=r, column=col)
                existing = cell.value
                if existing in (None, ""):
                    cell.value = fval
                    cell.fill = YELLOW
                    stats["filled"] += 1
                    log_rows.append([code, label, "", fval, "filled"])
                else:
                    if not values_equal(existing, fval, numeric=numeric, is_norm=is_norm):
                        cell.fill = ORANGE
                        cell.comment = Comment(f"Fabory value: {fval}", "fill_excel.py")
                        stats["conflict"] += 1
                        log_rows.append([code, label, existing, fval, "differs from Fabory (not changed)"])

        if r in extra_rows:
            for c in range(1, N_TOTAL_COLS + 1):
                ws.cell(row=r, column=c).fill = LIGHT_BLUE
            stats["new_variant"] += 1
            log_rows.append([code, "-", "-", "-", "new variant"])

    for row in log_rows:
        log_ws.append(row)
    for i, w in enumerate([16, 18, 20, 20, 24], start=1):
        log_ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(DST)

    print("Klaar. Kopie geschreven naar:", DST)
    print("Extra variant-rijen toegevoegd:", len(extra_rows))
    print("Preferred Supplier ingevuld:", stats["pref_supplier"])
    print("Supplier Article Nr ingevuld:", stats["supplier_artnr"])
    print("MOQ ingevuld:", stats["moq"])
    print("Clamping range ingevuld:", stats["clamping"])
    print("Drill hole ingevuld:", stats["drill_hole"])
    print("Drive type ingevuld:", stats["drive_type"])
    print("Cellen aangevuld (geel):", stats["filled"])
    print("Afwijkingen t.o.v. Fabory (oranje):", stats["conflict"])
    print("Nieuwe-variant rijen gelogd:", stats["new_variant"])


if __name__ == "__main__":
    main()
