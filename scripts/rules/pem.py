"""PennEngineering (PEM) self-clinching standoffs (SO/BSO) and studs (FH/HFH).

Not on Fabory: dimensions come straight from PEM's own metric datasheets, so this
module ignores the (empty) Fabory `d` dict and loads its own JSON per code from
data/pem/<code>.json (see that JSON's "source" field for the exact PEM page).

Nightly build test codes (1 per 'kind'):
    pem_so    -> SO-M5-4ZI
    pem_bso   -> BSO-M4-8ZI
    pem_stud  -> FH-M5-30ZI   (also covers HFH-M10-20ZI, same builder)

Assumptions (simplification allowed per brief; mass may deviate, main dims must not):
- Standoff body modelled as a regular hexagon prism, across-flats "s" = PEM's
  "H Nom." column (body width) for the SO/BSO metric table. No knurl/thread cut;
  axial bore is a plain hole at the nominal thread diameter "d".
  SO = through bore. BSO = blind bore; PEM only publishes a *minimum* full-thread
  depth "F" (kept in the JSON for reference), so the model instead leaves a fixed
  1.5 mm solid base (blind_depth = L - 1.5) -- this happens to equal F exactly for
  the shortest BSO-M4 length (8 mm) and is conservative (deeper) for longer parts.
- FH/HFH studs modelled as: head (dia H, thickness T) + shank at nominal d (no
  thread cut) down to overall length L. The unthreaded shank length "S" from the
  datasheet is stored for reference but not drawn separately (same diameter as
  the threaded portion in this simplified model).
  FH (flush head, sits within the sheet): head thickness T taken as PEM's
  "Min. Sheet Thickness" for that thread size (a flush head is only as thick as
  the thinnest sheet it can sit flush in).
  HFH (heavy-duty, non-flush, projects above the sheet): head diameter H and
  thickness T taken directly from PEM's "S Max." / "T Max." columns.

Sources: PEM bulletins
  SO  - https://www.pemnet.com/wp-content/uploads/sites/2/2022/06/sodata.pdf
  FH  - https://www.pemnet.com/wp-content/uploads/sites/2/2022/06/fhdata.pdf (also covers HFH)
plus individual pemnet.com product-finder pages used to cross-check a few codes.
"""
import json, os
from swlib import DATA_DIR

PEM_DIR = os.path.join(DATA_DIR, "pem")


def _load(code):
    f = os.path.join(PEM_DIR, f"{code}.json")
    if not os.path.exists(f):
        return None
    return json.load(open(f, encoding="utf-8"))


def geometry(d, row):
    if (row[24] or "") != "PennEngineering (PEM)":
        return None
    pj = _load(str(row[23]))
    if not pj:
        return None
    a = pj["Afmetingen"]
    kind = pj["type"]
    if kind == "SO":
        return "pem_so", dict(d=a["d"], s=a["s"], L=a["L"])
    if kind == "BSO":
        return "pem_bso", dict(d=a["d"], s=a["s"], L=a["L"])
    if kind in ("FH", "HFH"):
        return "pem_stud", dict(d=a["d"], L=a["L"], H=a["H"], T=a["T"])
    return None


def pem_so(b, p):
    """Through-hole threaded standoff: hex tube, across-flats s, bore d, length L."""
    b.plane_extrude(lambda: (b.hexagon(p["s"]), b.circle(p["d"])), p["L"], down=True)


def pem_bso(b, p):
    """Blind threaded standoff: solid hex body, then a blind bore cut from the top."""
    wall = 1.5
    depth = max(1.0, p["L"] - wall)
    b.plane_extrude(lambda: b.hexagon(p["s"]), p["L"], down=True)
    b.plane_extrude(lambda: b.circle(p["d"]), depth, cut=True, down=True)


def pem_stud(b, p):
    """FH/HFH stud: head (dia H, thickness T) + plain shank at nominal d to length L."""
    d, L, H, T = p["d"], p["L"], p["H"], p["T"]
    b.revolve([(0, 0), (H / 2, 0), (H / 2, -T), (d / 2, -T), (d / 2, -L), (0, -L)])


BUILDERS = {"pem_so": pem_so, "pem_bso": pem_bso, "pem_stud": pem_stud}
