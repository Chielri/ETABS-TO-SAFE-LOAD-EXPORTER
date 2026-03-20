# CLAUDE.md

## Project Overview
ETABS to SAFE Shell Uniform Load Exporter — a Python tool that transfers shell uniform loads from CSI ETABS to CSI SAFE via their COM APIs.

## Repository Structure
- `etabs_to_safe.py` — CLI script (no GUI, prints to stdout)
- `etabs_to_safe_gui.py` — Tkinter GUI with logging, debug toggle, progress bar, log save
- `etabs-api.skill` — ETABS API reference (ZIP archive, extract to read)
- `safe-api.skill` — SAFE API reference (ZIP archive, extract to read)
- `requirements.txt` — Python dependencies
- `.github/workflows/build.yml` — GitHub Actions workflow to build Windows .exe

## Key Concepts
- Both ETABS and SAFE expose COM APIs via `comtypes` on Windows
- ETABS connection: `ETABSv1.Helper` → `CSI.ETABS.API.ETABSObject` → `SapModel`
- SAFE connection: `SAFEv1.Helper` → `CSI.SAFE.API.ETABSObject` → `SapModel`
  - **CRITICAL**: SAFE reuses ETABS API infrastructure. The ProgID is `ETABSObject`, NOT `SAFEObject`.
- ETABS exposes rich COM interfaces: `AreaObj.GetLoadUniform`, `AreaObj.SetLoadUniform`, `LoadPatterns`, etc.
- SAFE's primary API is **database tables** (`cDatabaseTables`): `GetTableForDisplayArray`, `GetTableForEditingArray`, `SetTableForEditingArray`, `ApplyEditedTables`
  - SAFE does NOT expose `AreaObj`, `LoadPatterns`, or other ETABS-specific COM interfaces
  - All SAFE model manipulation (loads, geometry, properties) goes through database tables
  - Key table: `"Area Load Assignments - Uniform"` for reading/writing uniform area loads
- Slab matching: ETABS slab label (`GetLabelFromName`) → SAFE slab unique name (via database tables)
- Shell uniform loads: `AreaObj.GetLoadUniform` (read from ETABS), database tables (write to SAFE)

## API Reference
- Consult `etabs-api.skill` for ETABS API details (extract with `unzip etabs-api.skill`)
- Consult `safe-api.skill` for SAFE API details (extract with `unzip safe-api.skill`)
- ETABS COM return values: `[OutParam1, OutParam2, ..., ret]` — ret is LAST
- SAFE database table data is a flat 1D array, row-by-row (N fields × M records)
- All API methods return int: 0 = success, nonzero = failure

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
