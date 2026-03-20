# CLAUDE.md

## IMPORTANT: Read API Reference Files First
**Before making ANY code changes**, you MUST read the following API reference files. ETABS and SAFE API documentation is extremely scarce online and not reliably in training data. These skill files are the authoritative source for correct API usage:
- `etabs-api.skill` — ETABS COM API reference (methods, parameters, return values, table names)
- `safe-api.skill` — SAFE COM API reference (methods, parameters, return values, table names)

**Always consult these files** when working with COM API calls, database table names, method signatures, parameter types, or return value structures. Do not guess or rely on training data for ETABS/SAFE API details.

## Project Overview
ETABS to SAFE Shell Uniform Load Exporter — a Python tool that transfers shell uniform loads from CSI ETABS to CSI SAFE via their COM APIs.

## Repository Structure
- `etabs_to_safe.py` — CLI script (no GUI, prints to stdout)
- `etabs_to_safe_gui.py` — Tkinter GUI with logging, debug toggle, progress bar, log save
- `etabs-api.skill` — ETABS COM API reference documentation (MUST READ before coding)
- `safe-api.skill` — SAFE COM API reference documentation (MUST READ before coding)
- `requirements.txt` — Python dependencies
- `.github/workflows/build.yml` — GitHub Actions workflow to build Windows .exe and commit it to repo

## Key Concepts
- Both ETABS and SAFE expose COM APIs via `comtypes` on Windows
- ETABS connection: `ETABSv1.Helper` → `CSI.ETABS.API.ETABSObject` → `SapModel`
- SAFE connection: `SAFEv1.Helper` → `CSI.SAFE.API.SAFEObject` → `SapModel`
- Slab matching: ETABS slab label (`GetLabelFromName`) → SAFE slab unique name (`GetNameList`)
- Shell uniform loads: `AreaObj.GetLoadUniform` (read from ETABS), `AreaObj.SetLoadUniform` (write to SAFE)

## Caching Strategy
All read-heavy data is pre-cached from database tables before the per-slab export loop:
- `build_table_load_cache()` — ETABS uniform load assignments (avoids per-slab `GetLoadUniform` calls)
- `build_label_cache()` — ETABS area object labels/stories (avoids per-slab `GetLabelFromName` calls)
- `build_safe_load_cache()` — SAFE existing load assignments (avoids per-slab SAFE table reads)

Only write operations (delete/assign loads in SAFE) use COM per slab.

## Development
- Python 3.10+ on Windows (COM APIs are Windows-only)
- Install deps: `pip install -r requirements.txt`
- Run CLI: `python etabs_to_safe.py`
- Run GUI: `python etabs_to_safe_gui.py`
- Build exe: `pyinstaller --onefile --windowed --name ETABStoSAFE etabs_to_safe_gui.py`

## Testing
No automated tests — the tool requires running ETABS and SAFE instances with loaded models. Manual testing: select slabs in ETABS, run the tool, verify loads appear in SAFE.

## Code Style
- Simple, minimal code. No over-engineering.
- Use `logging` module (not print) in GUI version.
- Keep CLI version using print for simplicity.
