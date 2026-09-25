"""Extra rules: DIN 933, DIN 913, DIN 916, DIN 603, DIN 7981 C-H, ISO 14585 C/F, DIN 571, DIN 94.

Nachtelijke SW-test (1 code per 'kind', droog al gevalideerd met --check):
    python build_parts.py 51010.100.012 07860.060.040 07850.080.008 08320.100.100 \
        51660.063.050 26505.048.038 08280.080.055 50390.032.032
    kind -> testcode: hexbolt=51010.100.012 (DIN933) / 08280.080.055 (DIN571),
    setscrew=07860.060.040 (DIN913) / 07850.080.008 (DIN916), carriagebolt=08320.100.100,
    panhead=51660.063.050 (DIN7981) / 26505.048.038 (ISO14585 F), splitpen=50390.032.032.
    Controleer massa SW t.o.v. Fabory-gewicht na de build (zie build_log.csv).

Vereenvoudigingen (zie AGENT_BRIEF.md regel 1/2):
- Alle schroefdraad wordt weggelaten: gladde schacht op nominale d.
- Kruis-/Torx-uitsparingen in platte-/bolkoppen worden weggelaten (mag volgens opdracht).
- DIN 916 (komvormige punt) wordt gemodelleerd als platte punt (mag volgens opdracht).
- DIN 603: hoogte van het vierkant onder de kop staat niet in de Fabory-data -> aangenomen
  gelijk aan de sleutelwijdte van het vierkant (fn), normwaarde-aanname.
- DIN 94: splitpen vereenvoudigd tot cilinder Ø d met een oog (ring) aan de kop; oogdikte
  aangenomen = d, oog-gatdiameter aangenomen = 0.4 * c (niet in Fabory-data, aannames).
"""
from swlib import num


def _dims(d):
    a = d.get("Afmetingen") or {}
    g = lambda *keys: next(num(a[k]) for k in keys if k in a)
    return a, g


def geometry(d, row):
    a, g = _dims(d)
    title = d.get("title") or ""
    if not a:
        return None
    try:
        if "DIN 933" in title:
            return "hexbolt", dict(d=g("d-D"), L=g("L (mm)"), k=g("k (max.)", "k"), s=g("s"))
        if "DIN 571" in title:
            return "hexbolt", dict(d=g("D (mm)", "d-D"), L=g("L (mm)"), k=g("k (max.)", "k"), s=g("s"))
        if "DIN 913" in title:
            return "setscrew", dict(d=g("d-D"), L=g("L (mm)"), s=g("s"), t=g("t (min.)", "t"))
        if "DIN 916" in title:
            return "setscrew", dict(d=g("d-D"), L=g("L (mm)"), s=g("s"), t=g("t (min.)", "t"))
        if "DIN 603" in title:
            return "carriagebolt", dict(d=g("d-D"), L=g("L (mm)"), dk=g("dk(max.)", "dk"),
                                        k=g("k (max.)", "k"), fn=g("fn (max.)", "fn"))
        if "DIN 7981" in title:
            return "panhead", dict(d=g("D (mm)"), L=g("L (mm)"), dk=g("dk(max.)", "dk"), k=g("k (max.)", "k"))
        if "ISO 14585" in title:
            return "panhead", dict(d=g("D (mm)"), L=g("L (mm)"), dk=g("dk(max.)", "dk"), k=g("k (max.)", "k"))
        if "DIN 94" in title:
            return "splitpen", dict(d=g("d (max.)", "d (min.)"), L=g("L (mm)"), c=g("c (max.)"))
    except StopIteration:
        return None
    return None


def shank(p, y0, y1):
    r = p["d"] / 2
    return [(r, y0), (r, y1), (0, y1)]


def hexbolt(b, p):
    """External hex head bolt (DIN 933 volledig draad, DIN 571 houtdraadbout): kop boven Y=0."""
    d, L, k, s = p["d"], p["L"], p["k"], p["s"]
    b.revolve([(0, -k), (d / 2, -k)] + shank(p, -k, -(k + L)))
    b.plane_extrude(lambda: b.hexagon(s), k, cut=False, down=True)


def setscrew(b, p):
    """Stelschroef (DIN 913/916), volle cilinder + binnenzeskant aan de kop, platte punt."""
    d, L, s, t = p["d"], p["L"], p["s"], p["t"]
    b.revolve([(0, 0), (d / 2, 0)] + shank(p, 0, -L))
    b.plane_extrude(lambda: b.hexagon(s), t, cut=True)


def carriagebolt(b, p):
    """DIN 603 slotbout: bolkop (vereenvoudigd als cilinder) + vierkant onder de kop + ronde schacht."""
    d, L, dk, k, fn = p["d"], p["L"], p["dk"], p["k"], p["fn"]
    qh = fn  # hoogte vierkant: aanname (niet in Fabory-data)
    b.revolve([(0, k), (dk / 2, k), (dk / 2, 0), (d / 2, 0), (d / 2, -L), (0, -L)])
    b.plane_extrude(lambda: b.square(fn), qh, cut=False, down=True)


def panhead(b, p):
    """Pancilinderplaatschroef (DIN 7981 C-H, ISO 14585 C/F): bolkop, kruis/Torx weggelaten."""
    d, L, dk, k = p["d"], p["L"], p["dk"], p["k"]
    b.revolve([(0, 0), (dk / 2, 0), (dk / 2, -k)] + shank(p, -k, -(k + L)))


def splitpen(b, p):
    """DIN 94 splitpen: vereenvoudigd tot cilinder Ø d met een oog (ring) aan de kop."""
    d, L, c = p["d"], p["L"], p["c"]
    eye_h = d  # aanname: oogdikte ~ draaddiameter
    hole = 0.4 * c  # aanname: ooggat
    b.revolve([(0, 0), (c / 2, 0), (c / 2, -eye_h), (d / 2, -eye_h), (d / 2, -L), (0, -L)])
    b.plane_extrude(lambda: b.circle(hole), eye_h, cut=True, down=True)


BUILDERS = {
    "hexbolt": hexbolt,
    "setscrew": setscrew,
    "carriagebolt": carriagebolt,
    "panhead": panhead,
    "splitpen": splitpen,
}
