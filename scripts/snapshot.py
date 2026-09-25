"""Export framed JPG snapshots of built parts to data/snapshots/<code>.jpg (for the progress page).

Usage: python snapshot.py <code> ...
"""
import os, sys
import pythoncom, win32com.client as w32
from swlib import connect, sw_lock, OUT_DIR, DATA_DIR

SNAP_DIR = os.path.join(DATA_DIR, "snapshots")


def snap(sw, code):
    os.makedirs(SNAP_DIR, exist_ok=True)
    path = os.path.join(OUT_DIR, f"{code}.SLDPRT")
    doc = sw.OpenDoc(path, 1)
    if doc is None:
        return None
    title = doc.GetTitle
    err = w32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    sw.ActivateDoc3(title, False, 0, err)
    doc.ShowNamedView2("*Isometric", 7)
    doc.ViewZoomtofit2()
    doc.GraphicsRedraw2()
    out = os.path.join(SNAP_DIR, f"{code}.jpg")
    doc.SaveAs3(out, 0, 3)
    sw.CloseDoc(title)
    return out


if __name__ == "__main__":
    sw = connect()
    with sw_lock():
        for c in sys.argv[1:]:
            print(snap(sw, c))
