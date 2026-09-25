"""Moeren & ringen: DIN 985, DIN 1587, DIN 6923, DIN 935-1, DIN 464, DIN 9021, DIN 127 B,
DIN 6798 A, DIN 2093. (DIN 934 / DIN 125-1A zitten al in core.py.)

Nachtelijke SW-test (1 code per 'kind', via `python build_parts.py <code>`):
  lock_nut    (DIN 985)   12348.100.001
  cap_nut     (DIN 1587)  55400.100.001
  flange_nut  (DIN 6923)  11610.080.001
  crown_nut   (DIN 935-1) 12010.160.001
  knurl_screw (DIN 464)   100800550
  washer      (DIN 9021)  38210.100.001   (kind hergebruikt uit core.py)
  washer      (DIN 127 B) 37020.100.001   (idem)
  washer      (DIN 6798 A)37420.060.001   (idem)
  belleville  (DIN 2093)  36307.188.010

Alle bovenstaande codes gedroogvalideerd met `python build_parts.py --check <code>` (geen SW,
zie coordinator-wijziging: overdag geen SolidWorks gebruiken). Massacontrole t.o.v. Fabory-
gewicht moet 's nachts nog gebeuren.

Aannames / normwaarden (Fabory-data geeft niet alles):
- DIN 985 / DIN 935-1: hex-deel hoogte 'm' niet apart in Fabory-data -> tabel HEX_M
  (= standaard DIN 934 moerhoogtes per d). Nylon-ring (985) resp. kroon (935) = totaal 'h'/'m1'
  min die hex-hoogte. Kroon-sleuven: vereenvoudigd als 6 rechte radiale sleuven (breedte 'n'),
  ipv gebogen sleuf-profiel.
- DIN 6923: flensdikte 'c' niet in Fabory-data -> tabel FLANGE_C (typische DIN 6923-waarden).
- DIN 464 (code 100800550, M6x35): geen Fabory-pagina voor dit interne artikelnummer
  (fetch = NO-PDF). Kopmaten (dk, h, ds, k) hergebruikt van Fabory M6-referentie
  (art. 51829060016, RVS M6x16 DIN 464): dk=24, h=15, ds=12, k=5. Kartel weggelaten -> gladde
  cilinder (conform briefing). Lengte L=35 uit Excel-kolom Length/naam "M6x35".
- DIN 127 B / DIN 6798 A: gemodelleerd als vlakke ring (kind "washer", core.py), met dikte =
  's' resp. 's1' (materiaaldikte), NIET 'h' (dat is de opengesperde/getande hoogte, niet
  representatief voor een vlakke ring). d1/d2 = min./max. Fabory-maten.
- DIN 2093: kegelring (revolve, 4 punten), h0 = Lo - t (vrije hoogte min. dikte); dikte
  verticaal benaderd i.p.v. loodrecht op het conische vlak (vereenvoudiging).
"""
import math
from swlib import num

HEX_M = {3: 2.4, 4: 3.2, 5: 4.7, 6: 5.2, 7: 5.7, 8: 6.8, 10: 8.4, 12: 10.8, 14: 12.8,
         16: 14.8, 18: 15.8, 20: 18.0, 22: 19.4, 24: 21.5, 27: 23.8, 30: 25.6}
FLANGE_C = {5: 0.6, 6: 0.7, 8: 0.8, 10: 0.9, 12: 1.1, 14: 1.2, 16: 1.4, 20: 1.7}
KNURL464 = {6: dict(dk=24, h=15, ds=12, k=5)}  # M6 uit Fabory-referentie art. 51829060016


def _get(src, *keys):
    return next((num(src[k]) for k in keys if k in src), None)


def geometry(d, row):
    norm = str(row[28] or "")
    title = d.get("title") or ""
    a = d.get("Afmetingen") or {}
    spec = d.get("Specificaties") or {}
    size = row[25]
    try:
        nominal = num(size) if size else None
    except ValueError:
        nominal = None

    if "DIN 985" in norm or "DIN 985" in title:
        if not a:
            return None
        dnom, s, h = _get(a, "d-D"), _get(a, "s"), _get(a, "h")
        if None in (dnom, s, h):
            return None
        m = min(HEX_M.get(round(dnom), 0.8 * dnom), h - 0.5)
        return "lock_nut", dict(d=dnom, s=s, h=h, m=m)

    if "DIN 1587" in norm or "DIN 1587" in title:
        if not a:
            return None
        dnom, s, h = _get(a, "d-D"), _get(a, "s"), _get(a, "h")
        t, dk = _get(a, "t (min.)", "t"), _get(a, "dk(max.)", "dk")
        if None in (dnom, s, h, t, dk):
            return None
        return "cap_nut", dict(d=dnom, s=s, h=h, t=t, dk=dk)

    if norm == "DIN 6923" or "DIN 6923" in title:
        if not a:
            return None
        dnom, s = _get(a, "d-D"), _get(a, "s")
        m, dc = _get(a, "m (max.)", "m"), _get(a, "dc(max.)", "dc")
        if None in (dnom, s, m, dc):
            return None
        c = min(FLANGE_C.get(round(dnom), 0.1 * dnom), m - 1)
        return "flange_nut", dict(d=dnom, s=s, m=m, dc=dc, c=c)

    if "DIN 935" in norm or "DIN 935" in title:
        if not a:
            return None
        dnom, s = _get(a, "d-D"), _get(a, "s")
        m1, n = _get(a, "m"), _get(a, "n")
        if None in (dnom, s, m1, n):
            return None
        slots = 6
        for k, v in a.items():
            if k.lower().startswith("aantal gleuven"):
                try:
                    slots = int(num(v))
                except ValueError:
                    pass
        m = min(HEX_M.get(round(dnom), 0.6 * dnom), m1 - 1)
        return "crown_nut", dict(d=dnom, s=s, m1=m1, m=m, n=n, slots=slots)

    if norm == "DIN 464" or "DIN 464" in title:
        dnom = nominal or _get(a, "d-D")
        if dnom is None:
            return None
        L = num(row[29]) if row[29] else _get(a, "L (mm)", "L")
        if L is None:
            return None
        dk, h, ds, k = _get(a, "dk"), _get(a, "h"), _get(a, "ds"), _get(a, "k")
        if None in (dk, h, ds, k):
            ref = KNURL464.get(round(dnom))
            if ref is None:
                return None
            dk, h, ds, k = ref["dk"], ref["h"], ref["ds"], ref["k"]
        return "knurl_screw", dict(d=dnom, L=L, dk=dk, h=h, ds=ds, k=k)

    if "DIN 9021" in norm or "DIN 9021" in title:
        if not a:
            return None
        d1, d2, h = _get(a, "d1"), _get(a, "d2"), _get(a, "h")
        if None in (d1, d2, h):
            return None
        return "washer", dict(d1=d1, d2=d2, h=h)

    if "DIN 127" in norm or "DIN 127" in title:
        if not a:
            return None
        d1, d2, s = _get(a, "d1(min.)", "d1"), _get(a, "d2(max.)", "d2"), _get(a, "s")
        if None in (d1, d2, s):
            return None
        return "washer", dict(d1=d1, d2=d2, h=s)

    if "DIN 6798" in norm or "DIN 6798" in title:
        if not a:
            return None
        d1, d2, s1 = _get(a, "d1(min.)", "d1"), _get(a, "d2(max.)", "d2"), _get(a, "s1")
        if None in (d1, d2, s1):
            return None
        return "washer", dict(d1=d1, d2=d2, h=s1)

    if "DIN 2093" in norm or "DIN 2093" in title:
        de = _get(spec, "Buitendiameter (mm)", "dₑ") or _get(a, "D (mm)")
        di = _get(spec, "di") or _get(a, "di (H12)")
        t = _get(a, "Dikte", "Wanddikte t")
        Lo = _get(a, "Lo")
        if None in (de, di, t, Lo):
            return None
        h0 = max(Lo - t, 0.1)
        return "belleville", dict(de=de, di=di, t=t, h0=h0)

    return None


# ------------------------------------------------------------ builders
def _slot(b, r_out, width, theta):
    ca, sa = math.cos(theta), math.sin(theta)
    rot = lambda x, y: (x * ca - y * sa, x * sa + y * ca)
    b.polyline([rot(-1, -width / 2), rot(r_out + 1, -width / 2),
                rot(r_out + 1, width / 2), rot(-1, width / 2)])


def lock_nut(b, p):
    """DIN 985: hex body (height m) + nylon collar (cylinder inscribed in flats) to total h."""
    b.plane_extrude(lambda: (b.hexagon(p["s"]), b.circle(p["d"])), p["m"], down=True)
    b.plane_extrude(lambda: (b.circle(p["s"]), b.circle(p["d"])), p["h"] - p["m"], down=False)


def cap_nut(b, p):
    """DIN 1587: hex body (height t) + solid dome cylinder (blind) to total h."""
    b.plane_extrude(lambda: (b.hexagon(p["s"]), b.circle(p["d"])), p["t"], down=True)
    b.plane_extrude(lambda: b.circle(p["dk"]), p["h"] - p["t"], down=False)


def flange_nut(b, p):
    """DIN 6923: hex body (height m - c) + flange disc (dia dc, thickness c)."""
    b.plane_extrude(lambda: (b.hexagon(p["s"]), b.circle(p["d"])), p["m"] - p["c"], down=True)
    b.plane_extrude(lambda: (b.circle(p["dc"]), b.circle(p["d"])), p["c"], down=False)


def crown_nut(b, p):
    """DIN 935-1: hex body (height m1) with 'slots' radial slots (width n) cut into the top."""
    b.plane_extrude(lambda: (b.hexagon(p["s"]), b.circle(p["d"])), p["m1"], down=True)
    r_out = p["s"] / math.sqrt(3)
    n, width = p["slots"], p["n"]

    # one cut per slot: overlapping slot rectangles in a single sketch give an invalid profile
    for i in range(n):
        b.plane_extrude(lambda a=2 * math.pi * i / n: _slot(b, r_out, width, a), p["m1"] - p["m"], cut=True, down=True)


def knurl_screw(b, p):
    """DIN 464: knurled cylindrical head (kartel weggelaten) + neck + smooth shank at d."""
    d, L, dk, h, ds, k = p["d"], p["L"], p["dk"], p["h"], p["ds"], p["k"]
    shank_len = max(L - k, 1.0)
    b.revolve([(0, 0), (dk / 2, 0), (dk / 2, -h), (ds / 2, -h), (ds / 2, -h - k),
               (d / 2, -h - k), (d / 2, -h - k - shank_len), (0, -h - k - shank_len)])


def belleville(b, p):
    """DIN 2093: conical ring (schotelveer), revolve, thickness t approximated vertically."""
    ri, re, t, h0 = p["di"] / 2, p["de"] / 2, p["t"], p["h0"]
    b.revolve([(ri, 0), (ri, t), (re, t + h0), (re, h0)])


BUILDERS = {
    "lock_nut": lock_nut,
    "cap_nut": cap_nut,
    "flange_nut": flange_nut,
    "crown_nut": crown_nut,
    "knurl_screw": knurl_screw,
    "belleville": belleville,
}
