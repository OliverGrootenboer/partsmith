"""Core rules (tested): DIN 912, ISO 10642, ISO 7380-2, DIN 934 nut, DIN 125 washer, Nord-Lock."""
import math
from swlib import num, hex_socket_depth


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
        if "DIN 912" in title or "ISO 4762" in title:
            return "din912", dict(d=g("d-D"), L=g("L (mm)"), k=g("k (max.)", "k"), s=g("s"), dk=g("dk(max.)", "dk"))
        if "ISO 10642" in title:
            return "iso10642", dict(d=g("d-D"), L=g("L (mm)"), k=g("k (max.)", "k"), s=g("s"), dk=g("dk(max.)", "dk"))
        if "ISO 7380-2" in title:
            return "iso7380_2", dict(d=g("d-D"), L=g("L (mm)"), k=g("k (max.)", "k"), s=g("s"), dc=g("dc(max.)", "dc"))
        if "DIN 934" in title:
            return "nut", dict(d=g("d-D"), m=g("m"), s=g("s"))
        if "NORD-LOCK" in title.upper():
            # set thickness differs per material: 's (Set) Del pro' / 's (Set) A4' / 's (Set) 254 SMO'
            tl = title.upper()
            want = "254 SMO" if "254 SMO" in tl else "A4" if " A4" in tl else "Del pro"
            t = next((num(v) for k, v in a.items() if k.startswith("s (Set)") and want in k),
                     next(num(v) for k, v in a.items() if k.startswith("s (Set)")))
            return "washer", dict(d1=g("d1"), d2=g("d2"), h=t)
        if "DIN 125" in title:
            return "washer", dict(d1=g("d1"), d2=g("d2"), h=g("h"))
    except StopIteration:
        return None
    return None


def shank(p, y0, y1):
    r = p["d"] / 2
    return [(r, y0), (r, y1), (0, y1)]


def din912(b, p):
    d, L, dk, k = p["d"], p["L"], p["dk"], p["k"]
    b.revolve([(0, 0), (dk / 2, 0), (dk / 2, -k)] + shank(p, -k, -k - L))
    b.plane_extrude(lambda: b.hexagon(p["s"]), 0.6 * d, cut=True)


def iso10642(b, p):
    d, L, dk = p["d"], p["L"], p["dk"]
    cone = (dk - d) / 2  # theoretical sharp 90° head
    b.revolve([(0, 0), (dk / 2, 0)] + shank(p, -cone, -L))
    b.plane_extrude(lambda: b.hexagon(p["s"]), hex_socket_depth(p["k"]), cut=True)


def iso7380_2(b, p):
    d, L, dc, k = p["d"], p["L"], p["dc"], p["k"]
    dk, c = 1.75 * d, 0.15 * d  # dome dia / flange thickness (ISO 7380 approx.)
    yc = (k * k - c * c - (dk / 2) ** 2) / (2 * (k - c))
    R = k - yc
    am = (math.pi / 2 + math.atan2(c - yc, dk / 2)) / 2
    top = lambda y: y - k
    b.revolve([(0, top(k)), ("arc", (dk / 2, top(c)), (R * math.cos(am), top(yc + R * math.sin(am)))),
               (dc / 2, top(c)), (dc / 2, top(0))] + shank(p, top(0), top(-L)))
    b.plane_extrude(lambda: b.hexagon(p["s"]), hex_socket_depth(k), cut=True)


def nut(b, p):
    b.plane_extrude(lambda: (b.hexagon(p["s"]), b.circle(p["d"])), p["m"])


def washer(b, p):
    b.plane_extrude(lambda: (b.circle(p["d2"]), b.circle(p["d1"])), p["h"])


BUILDERS = {"din912": din912, "iso10642": iso10642, "iso7380_2": iso7380_2, "nut": nut, "washer": washer}
