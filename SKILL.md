---
name: partsmith
description: Build a SolidWorks library of standard/catalog parts from an Excel PLM import list - currently fasteners (bolts, nuts, washers, rivets, rivet nuts, PEM, plugs), extensible with one rule module per part family. Downloads Fabory datasheets, models each norm with a parametric rule via the SolidWorks API, assigns material so SolidWorks computes the mass, saves one SLDPRT per article number and fills a copy of the import list (supplier, MOQ, drill hole, drive type, clamping range). Use when the user wants SolidWorks parts, a CAD/part library or PLM parts created from a parts list, or wants to add lengths, materials, variants or new part families to such a library.
---

# partsmith (Excel → SolidWorks library parts)

Turns an Excel import list of fasteners into one SolidWorks part per article number, with correct main
dimensions and a SolidWorks-computed mass. Proven on 658 list rows + 1,143 Fabory variants = 1,801 parts
(26 part types, ~3 h unattended build, 0 build errors).

## Requirements

- Windows, SolidWorks (licensed) on the same machine, Python 3 with `openpyxl` and `pywin32`, `curl.exe`.
- A project folder with `config.json` (copy `config.example.json`):
  - `excel` – the import list, relative to the project folder. **Read only; the scripts never modify it.**
  - `part_template` – the company part template (`.PRTDOT`). Ask the user for the exact path.
- Run every script with the project folder as working directory, or set `FASTENER_PROJECT`.
  Scripts live in this skill's `scripts/` folder: `python "<skill>/scripts/build_parts.py" --all`.

## Import list layout

Sheet `Sheet compact`, data from row 3. 0-based columns: 5 Name, 21 Type, 22 Material, 23 Manufacturer Code,
24 Manufacturer, 25 Size (M), 26 Strength Class, 27 Thread Pitch, 28 Norm, 29 Length (mm).
If the user's list differs, adapt the indexes (they are used in `build_parts.py`, `fetch_all.py`,
`extend_list.py`, `fill_excel.py`, `rules/*.py`) before anything else, and confirm with the user.

## Workflow

1. **Spar first, build nothing.** Confirm goal and constraints with the user (never edit the Excel without asking;
   file name = article code; dimensions over mass). Check SolidWorks is reachable
   (`win32com.client.GetActiveObject("SldWorks.Application")`) and one datasheet downloads.
2. **Fetch data**: `fetch_all.py` → `PDF/<code>.pdf` + `data/<code>.json` (title, Specificaties, Afmetingen,
   Artikelnummers, moq). Non-Fabory codes that Fabory sells (e.g. Nord-Lock) go in `data/fabory_map.json`
   (`{"<excel code>": "<fabory code>"}`). Be polite: the script pauses between requests.
3. **Optional variants**: `variants.py` (all Fabory variants per family) → `extend_list.py` →
   `data/extra_rows.json`: every length of the M-sizes already in the list, all rivet/rivet-nut sizes.
   All scripts pick up these extra rows automatically.
4. **Dry run**: `build_parts.py --check --all` → geometry per code, "geen regel" = no rule yet.
5. **Rules**: one module per group in `scripts/rules/` (`geometry(d, row)` → `(kind, params)`, `BUILDERS`).
   Write a new rule for any norm without one. Conventions are in the `swlib.py` docstring
   (axis = Y, head top at Y=0, mm, no thread, smooth shank at nominal d).
6. **Test one part per kind** in SolidWorks, compare mass with the datasheet weight
   (> ~30–40 % off → suspect a dimension error; rivets and plugs excepted). Look at a snapshot
   (`snapshot.py <code>` → `data/snapshots/<code>.jpg`).
7. **Build**: `build_parts.py --all` (skips existing SLDPRT; identical geometry is modelled once and saved per
   material). Prefer running at night: SolidWorks is one instance and the user needs the licence by day.
8. **Downloaded CAD** (e.g. Elesa-Ganter via the browser): put it in `SLDPRT/<code>.SLDPRT`, list
   `{"<code>": "<Excel material>"}` in `data/elesa.json`, run `apply_materials.py`.
9. **Excel copy**: `fill_excel.py` → `<excel>_aangevuld.xlsx`: empty cells filled (yellow), differences from
   the supplier flagged (orange, not changed), extra variants (light blue), new columns Preferred Supplier,
   Supplier Article Nr, MOQ, Clamping range, Drill hole, Drive type, and a *Fill log* sheet.
10. **Report**: counts, codes the supplier does not know, mass outliers, open questions for the user.

`progress.py "<phase>" "<message>"` writes `data/progress.json` for a progress page if the user wants one.
`night_build.ps1 -Project <folder>` runs steps 2, 7, 8, 9 without AI.

## Rules of thumb (learned the hard way)

- **Unattended runs stall on permission prompts.** Before a night run, have the user allow `python *` and
  file edits in the project folder, or run the build from an attended session in the background.
- Keep the machine awake: `build_parts.py` requests it from Windows while running; a closed laptop lid still sleeps.
- `sw_lock()` serialises SolidWorks between processes; a killed build can leave `data/solidworks.lock` behind
  (remove it only when its PID is dead) and an open document (`CloseAllDocuments(True)`).
- pywin32 late binding: zero-argument COM methods are properties (`doc.EditDelete`, `ext.CreateMassProperty`).
  `CreatePolygon` takes a **vertex** point: circumradius = s/√3 for across-flats s.
- Cuts: try both directions; overlapping contours in one sketch fail → one cut per contour.
- Missing dimension table: `borrow()` copies head dimensions from the same norm + size and takes the length
  from the list — only for Fabory rows without own data (PEM/Elesa/fischer keep their own).
- Supplier weights can be wrong (Nord-Lock stainless listed ~11× too low). The mass check catches model errors
  and data errors; report both, never "fix" the model to match bad data.
- Parallel sub-agents speed up rule writing but cost a lot of tokens; give them one shared brief, their own
  rule file each, and no SolidWorks access during the day.
