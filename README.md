# partsmith

A Claude skill that forges a SolidWorks library of standard parts from an Excel PLM import list.
It starts with fasteners and grows one rule module per part family:
one SLDPRT per article number, correct main dimensions, material assigned so SolidWorks computes the mass,
and an enriched copy of the import list.

First run: 658 list rows + 1,143 supplier variants → 1,801 parts, 26 part types, ~3 hours unattended, 0 build errors.

## Install

Copy this folder to `~/.claude/skills/partsmith/` (or add it to a plugin). Claude loads `SKILL.md`
when you ask for SolidWorks parts from a fastener list.

## Use without Claude

```powershell
pip install openpyxl pywin32
copy config.example.json C:\my-project\config.json   # fill in excel + part_template
cd C:\my-project
python <skill>\scripts\fetch_all.py
python <skill>\scripts\build_parts.py --check --all
python <skill>\scripts\build_parts.py --all
python <skill>\scripts\fill_excel.py
```

## What is not in this repo

No import lists, datasheets, SolidWorks parts or supplier data: the `.gitignore` keeps `config.json`, `data/`,
`PDF/`, `SLDPRT/` and `*.xlsx` out. Datasheets belong to their suppliers; download them yourself.

## Scripts

| Script | Does |
| --- | --- |
| `fetch_fabory.py`, `fetch_all.py` | Download Fabory datasheet PDF + parse the product page per article |
| `variants.py`, `extend_list.py` | List all Fabory variants and create extra rows for missing lengths/sizes |
| `build_parts.py` | Group by geometry, model via the SolidWorks API, save per material, log mass |
| `rules/` | One parametric rule per norm/type (DIN 912, ISO 10642, DIN 933, nuts, washers, rivets, PEM, …) |
| `swlib.py` | Modelling primitives, materials, SolidWorks lock |
| `apply_materials.py` | Assign material to downloaded manufacturer CAD |
| `fill_excel.py` | Enriched copy of the import list with a fill log |
| `snapshot.py`, `progress.py` | Part snapshots and a progress summary for a dashboard |
| `night_build.ps1` | Full unattended run without AI |
