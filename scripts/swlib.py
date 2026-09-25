"""Shared SolidWorks helpers: modelling primitives, material/save, and a cross-process lock.

Conventions (all rule modules follow these):
- dimensions in mm; the API wants metres (MM factor applied here)
- part axis = model Y; top of head / top face at Y=0; body extends towards -Y
- revolve profiles are closed polygons of (r, y) points on the Front Plane;
  an item ("arc", (r_end, y_end), (r_mid, y_mid)) draws a 3-point arc
- true dimensions take priority over mass: no thread, smooth shank at nominal d
"""
import contextlib, math, os, time
import pythoncom, win32com.client as w32

ROOT = os.path.abspath(os.environ.get("FASTENER_PROJECT", os.getcwd()))  # project folder (config.json, data, SLDPRT, PDF)
DATA_DIR = os.path.join(ROOT, "data")
OUT_DIR = os.path.join(ROOT, "SLDPRT")



def _config():
    """Project settings from <project>/config.json (see config.example.json)."""
    import json
    f = os.path.join(ROOT, "config.json")
    if not os.path.exists(f):
        raise SystemExit(f"config.json ontbreekt in {ROOT} - kopieer config.example.json en vul de paden in")
    return json.load(open(f, encoding="utf-8"))


CONFIG = _config()
XLSX = os.path.join(ROOT, CONFIG["excel"])            # import list (read only, never modified)
TEMPLATE = CONFIG["part_template"]                    # company .PRTDOT
LOCK = os.path.join(DATA_DIR, "solidworks.lock")

DB = "SOLIDWORKS Materials"
MATERIALS = {
    "Steel, zinc plated": "Plain Carbon Steel",
    "Steel": "Plain Carbon Steel",
    "Steel, Delta Protekt": "Plain Carbon Steel",
    "Spring steel, zinc plated": "Plain Carbon Steel",
    "Spring steel, phosphated": "Plain Carbon Steel",
    "Stainless steel A2": "AISI 304",
    "Stainless steel": "AISI 304",
    "Stainless steel A4": "AISI 316 Annealed Stainless Steel Bar (SS)",
    "Stainless spring steel A4": "AISI 316 Annealed Stainless Steel Bar (SS)",
    "Stainless steel 254 SMO": "AISI 316 Annealed Stainless Steel Bar (SS)",
    "Polyamide (nylon)": "Nylon 101",
    "Aluminium": "6061 Alloy",
    "Aluminium / steel mandrel zinc plated": "6061 Alloy",
    "Copper": "Copper",
}
NOTHING = w32.VARIANT(pythoncom.VT_DISPATCH, None)
MM = 0.001


@contextlib.contextmanager
def sw_lock(timeout=1800):
    """Only one process drives SolidWorks at a time (parallel agents share one instance)."""
    t0 = time.time()
    while True:
        try:
            fd = os.open(LOCK, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
            break
        except FileExistsError:
            if time.time() - os.path.getmtime(LOCK) > 600:  # stale lock
                os.remove(LOCK)
            elif time.time() - t0 > timeout:
                raise TimeoutError("SolidWorks lock")
            time.sleep(1)
    try:
        yield
    finally:
        os.remove(LOCK)
        time.sleep(1.5)  # let waiting processes (tests, snapshots) take a turn


def num(v):
    """'M8' / '8,4' / '12 mm' -> float"""
    return float(str(v).replace(",", ".").replace("mm", "").strip().lstrip("M"))


def hex_socket_depth(k):
    return 0.6 * k


def connect():
    """Attach to a running SolidWorks, or start one (night runs)."""
    try:
        return w32.GetActiveObject("SldWorks.Application")
    except pythoncom.com_error:
        sw = w32.Dispatch("SldWorks.Application")
        sw.Visible = True
        time.sleep(20)  # let add-ins/licence settle
        return sw


class Builder:
    def __init__(self):
        self.sw = connect()

    def new(self):
        self.doc = self.sw.NewDocument(TEMPLATE, 0, 0, 0)
        self.sm = self.doc.SketchManager
        self.fm = self.doc.FeatureManager
        self.sm.AddToDB = True

    def sketch_on(self, plane):
        self.doc.ClearSelection2(True)
        self.doc.Extension.SelectByID2(plane, "PLANE", 0, 0, 0, False, 0, NOTHING, 0)
        self.sm.InsertSketch(True)

    def close_sketch(self):
        self.sm.InsertSketch(True)
        self.doc.ClearSelection2(True)

    def volume(self):
        mp = self.doc.Extension.CreateMassProperty
        return mp.Volume if mp else 0.0

    # ------------------------------------------------------------ features
    def revolve(self, pts):
        """Revolve closed (r, y) profile around Y. Can be called more than once (merges)."""
        self.sketch_on("Front Plane")
        ys = [p[1] for p in pts if p[0] != "arc"]
        self.sm.CreateCenterLine(0, (max(ys) + 1) * MM, 0, 0, (min(ys) - 1) * MM, 0)
        prev = pts[0]
        for p in pts[1:] + [pts[0]]:
            if p[0] == "arc":
                end, mid = p[1], p[2]
                self.sm.Create3PointArc(prev[0] * MM, prev[1] * MM, 0, end[0] * MM, end[1] * MM, 0,
                                        mid[0] * MM, mid[1] * MM, 0)
                prev = end
            else:
                self.sm.CreateLine(prev[0] * MM, prev[1] * MM, 0, p[0] * MM, p[1] * MM, 0)
                prev = p
        self.close_sketch()
        self.doc.FeatureByPositionReverse(0).Select2(False, 0)
        f = self.fm.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 2 * math.pi, 0,
                                    False, False, 0, 0, 0, 0, 0, True, False, True)
        if f is None:
            raise RuntimeError("revolve mislukt")
        return f

    def _extrude(self, depth, cut, flip):
        if cut:
            return self.fm.FeatureCut4(True, False, flip, 0, 0, depth * MM, 0, False, False, False, False,
                                       0, 0, False, False, False, False, False, False, True, False, False,
                                       False, 0, 0, False, False)
        return self.fm.FeatureExtrusion3(True, False, flip, 0, 0, depth * MM, 0, False, False, False, False,
                                         0, 0, False, False, False, False, True, True, True, 0, 0, False)

    def plane_extrude(self, draw, depth, cut=False, plane="Top Plane", down=None):
        """Sketch on a plane (default Top Plane, Y=0) with draw(), then extrude/cut `depth` mm.
        Cuts try both directions and keep the one that removes material.
        down=True/False forces the extrusion direction (True = towards -Y for Top Plane)."""
        if down is None:
            flips = (True, False)
        elif cut:  # cut direction semantics differ from extrude: try requested first, then the other
            flips = (down, not down)
        else:
            flips = (down,)
        for flip in flips:
            v0 = self.volume()
            self.sketch_on(plane)
            draw()
            self.close_sketch()
            sk = self.doc.FeatureByPositionReverse(0)
            sk.Select2(False, 0)
            f = self._extrude(depth, cut, flip)
            if f is not None and (not cut or self.volume() < v0 - 1e-12):
                sk.Select2(False, 0)
                self.doc.BlankSketch  # hide consumed sketch in previews
                self.doc.ClearSelection2(True)
                return f
            if f is not None:
                f.Select2(False, 0)
                self.doc.EditDelete
            sk.Select2(False, 0)
            self.doc.EditDelete
        raise RuntimeError("extrusie/snede mislukt")

    top_sketch_extrude = plane_extrude  # backwards compat

    # ------------------------------------------------ sketch entities (sketch coords, mm)
    def hexagon(self, s, x=0, y=0):
        """Hexagon with across-flats s (CreatePolygon's point is a vertex -> circumradius s/sqrt3)."""
        self.sm.CreatePolygon(x * MM, y * MM, 0, (x + s / math.sqrt(3)) * MM, y * MM, 0, 6, True)

    def square(self, s, x=0, y=0):
        """Square with across-flats s."""
        h = s / 2
        self.sm.CreateCenterRectangle(x * MM, y * MM, 0, (x + h) * MM, (y + h) * MM, 0)

    def circle(self, dia, x=0, y=0):
        self.sm.CreateCircleByRadius(x * MM, y * MM, 0, dia / 2 * MM)

    def rect(self, x1, y1, x2, y2):
        self.sm.CreateCornerRectangle(x1 * MM, y1 * MM, 0, x2 * MM, y2 * MM, 0)

    def polyline(self, pts, closed=True):
        seq = pts + [pts[0]] if closed else pts
        for a, b in zip(seq, seq[1:]):
            self.sm.CreateLine(a[0] * MM, a[1] * MM, 0, b[0] * MM, b[1] * MM, 0)

    # ------------------------------------------------------------ output
    def save_as(self, code, material):
        cfg = self.doc.ConfigurationManager.ActiveConfiguration.Name
        self.doc.SetMaterialPropertyName2(cfg, DB, material)
        self.doc.ForceRebuild3(False)
        self.doc.ShowNamedView2("*Isometric", 7)
        self.doc.ViewZoomtofit2()
        path = os.path.join(OUT_DIR, f"{code}.SLDPRT")
        err = self.doc.SaveAs3(path, 0, 3)  # silent + copy
        mass = self.doc.Extension.CreateMassProperty.Mass * 1000  # g
        return path, err, mass

    def close(self):
        self.sw.CloseDoc(self.doc.GetTitle)
