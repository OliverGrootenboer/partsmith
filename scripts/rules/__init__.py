"""Geometry rule registry.

Each module in this package defines:

    def geometry(d, row):
        '''d = parsed Fabory data (dict, may be {}), row = Excel row tuple
        (row[5]=Name, row[21]=Type, row[22]=Material, row[23]=Code, row[24]=Manufacturer,
         row[25]=Size, row[28]=Norm, row[29]=Length).
        Return (kind, params_dict_in_mm) when this module handles the part, else None.'''

    BUILDERS = {"kind": build_fn}   # build_fn(b: swlib.Builder, p: dict) -> None (b.new() already called)

Modules are tried in alphabetical order; the first non-None geometry() wins.
"""
import importlib, pkgutil

MODULES = [importlib.import_module(f"{__name__}.{m.name}") for m in sorted(pkgutil.iter_modules(__path__), key=lambda m: m.name)]


def geometry(d, row):
    for m in MODULES:
        g = m.geometry(d, row)
        if g:
            return g
    return None


def builder(kind):
    for m in MODULES:
        if kind in m.BUILDERS:
            return m.BUILDERS[kind]
    raise KeyError(kind)
