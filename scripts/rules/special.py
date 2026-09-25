"""Regels voor Fabory/Nord-Lock/fischer-onderdelen zonder Norm (popnagels, blindklinkmoeren, inserts,
inpersmoeren/-bouten, Nord-Lock (zie core.py) en fischer pluggen).

Testcodes voor de nachtelijke bouwrun (1 per kind, `python build_parts.py <code>`):
  rivet_dome  -> 34110.040.006   (Platbolkop blindklinknagel 4x6 ALU)
  rivet_cs    -> 34184.048.010   (Verzonkenkop blindklinknagel 4.8x10 A2, ISO 15983)
  rivet_nut   -> 69045.100.045   (Blindklinkmoer verzonkenkop M10 A2) en 69315.100.030 (cilinderkop) delen dezelfde builder
  insert      -> 71752.030.004   (Zelftappende schroefdraadbus M3x4)
  stud        -> 14425.060.030   (Self-clinching stud M6x30)
  clinch_nut  -> 14420.030.001   (Self-clinching nut M3)
  anchor      -> 70010           (fischer plug SX Plus 10x50, dims in Specificaties)
                 70016           (fischer Plug S 16, dims in Afmetingen -- andere familie, geen SX Plus in 16mm)
  (Nord-Lock washers lopen al via rules/core.py, zodra data/<code>.json bestaat -- fabory_map.json erbij gemaakt.)

Aannames / vereenvoudigingen:
  - Geen draad: klinkmoer/insert-boring en stud/moer-boring op nominale d (excel Size-kolom, row[25]).
  - Blindklinknagel-kop (rond of verzonken) en blindklinkmoer-kop: vlakke cilinder i.p.v. bolvormige/conische
    kop -- dk en kophoogte k kloppen, alleen de afronding/conus is vereenvoudigd.
  - Blindklinkmoer: buitendiameter schacht = Fabory "Boor ø" (drilgat), dat is de werkelijke schachtdiameter.
  - Self-clinching nut: vierkante plaatvoet (B, dikte H-h) + kartelboss (D, hoogte h), boring d doorlopend.
  - fischer plug: volle cilinder d x L (geen kanaal/lamellen) -- massa komt daardoor ca. 20-30% te hoog uit,
    dat is bekend en acceptabel (regel: maatafwijking telt, niet massa-afwijking).
  - fischer SX 16x80 (code 70016) heeft geen SX Plus opvolger bij Fabory; vervangen door fischer Plug S 16
    (63100.160.001, D=16 L=80 uit de norm-tabel van de S-plug, geen SX-vorm/kraag).
  - fischer SX 6x50 R (code 78185) is bij Fabory dezelfde SX Plus 6x50 (63125.060.002) als code 24827;
    de "R"-variant kon niet los teruggevonden worden op Fabory (twijfelgeval, zie eindrapport).
"""
import math, re
from swlib import num
from rules.core import shank


def _dims(d, key="Afmetingen"):
    a = d.get(key) or {}
    g = lambda *keys: next(num(a[k]) for k in keys if k in a)
    return a, g


def geometry(d, row):
    manu = row[24]
    if manu not in ("Fabory", "Nord-Lock", "fischer"):
        return None
    title = d.get("title") or ""
    tl = title.lower()
    try:
        if "blindklinknagel" in tl:
            a, g = _dims(d)
            if not a:
                return None
            p = dict(d=g("D (mm)"), L=g("L (mm)"), dk=g("ø dk (max.)", "ø dk"), k=g("k (max.)", "k"))
            return ("rivet_cs" if "verzonkenkop" in tl else "rivet_dome"), p

        if "blindklinkmoer" in tl:
            a, g = _dims(d)
            if not a:
                return None
            p = dict(d_out=g("Boor ø"), d_in=num(row[25]), L=g("L (mm)"),
                      dk=g("ø dk (max.)", "ø dk"), k=g("k (max.)", "k"))
            return "rivet_nut", p

        if "schroefdraadbus" in tl:
            a, g = _dims(d)
            if not a:
                return None
            p = dict(d_out=g("d2"), d_in=num(row[25]), L=g("L (mm)"))
            return "insert", p

        if "inpers draadeinden" in tl:
            a, g = _dims(d)
            if not a:
                return None
            p = dict(d=num(row[25]), D=g("D"), H=g("H (max.)", "H"), L=g("L (mm)"))
            return "stud", p

        if "inpersmoer" in tl:
            a, g = _dims(d)
            if not a:
                return None
            p = dict(B=g("B"), D=g("D", "D (max.)"), d=g("d"), H=g("H"), h=g("h (max.)", "h"))
            return "clinch_nut", p

        if "fischer plug" in tl or re.search(r"\bplug s\b", tl):
            spec = d.get("Specificaties") or {}
            a = d.get("Afmetingen") or {}
            dia_raw = spec.get("Diameter") or a.get("D (mm)")
            len_raw = spec.get("Lengte") or a.get("L (mm)")
            if not dia_raw or not len_raw:
                return None
            return "anchor", dict(d=num(dia_raw), L=num(len_raw))
    except StopIteration:
        return None
    return None


def _rivet_bore(d):
    # set rivet without mandrel: open sleeve, bore ~ mandrel dia (ISO 15977: 2.45 for d=4 -> ~0.6 d)
    return 0.6 * d / 2


def rivet_dome(b, p):
    """Dome (platbolkop) head as spherical cap, hollow sleeve, no mandrel."""
    d, L, dk, k = p["d"], p["L"], p["dk"], p["k"]
    ri, R0 = _rivet_bore(d), dk / 2
    yc = -(R0 ** 2 + k ** 2) / (2 * k)  # sphere centre on axis through (0,0) and (R0,-k)
    R = -yc
    y_ri = yc + math.sqrt(R ** 2 - ri ** 2)
    a0, a1 = math.atan2(y_ri - yc, ri), math.atan2(-k - yc, R0)
    am = (a0 + a1) / 2
    b.revolve([(ri, y_ri), ("arc", (R0, -k), (R * math.cos(am), yc + R * math.sin(am))),
               (d / 2, -k), (d / 2, -k - L), (ri, -k - L)])


def rivet_cs(b, p):
    """Countersunk head, hollow sleeve, no mandrel."""
    d, L, dk = p["d"], p["L"], p["dk"]
    ri = _rivet_bore(d)
    cone = (dk - d) / 2  # theoretical sharp countersink
    b.revolve([(ri, 0), (dk / 2, 0), (d / 2, -cone), (d / 2, -L), (ri, -L)])


def rivet_nut(b, p):
    d_in, d_out, L, dk, k = p["d_in"], p["d_out"], p["L"], p["dk"], p["k"]
    b.plane_extrude(lambda: (b.circle(dk), b.circle(d_in)), k, down=False)
    b.plane_extrude(lambda: (b.circle(d_out), b.circle(d_in)), L - k, down=True)


def insert(b, p):
    b.plane_extrude(lambda: (b.circle(p["d_out"]), b.circle(p["d_in"])), p["L"])


def stud(b, p):
    d, D, H, L = p["d"], p["D"], p["H"], p["L"]
    b.revolve([(0, 0), (D / 2, 0), (D / 2, -H)] + shank(p, -H, -L))


def clinch_nut(b, p):
    B, D, dd, H, h = p["B"], p["D"], p["d"], p["H"], p["h"]
    b.plane_extrude(lambda: (b.square(B), b.circle(dd)), H - h, down=True)
    b.plane_extrude(lambda: (b.circle(D), b.circle(dd)), h, down=False)


def anchor(b, p):
    b.plane_extrude(lambda: b.circle(p["d"]), p["L"])


BUILDERS = {
    "rivet_dome": rivet_dome,
    "rivet_cs": rivet_cs,
    "rivet_nut": rivet_nut,
    "insert": insert,
    "stud": stud,
    "clinch_nut": clinch_nut,
    "anchor": anchor,
}
