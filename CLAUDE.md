# CLAUDE.md

## Project Overview
ETABS to SAFE Shell Uniform Load Exporter — a Python tool that transfers shell uniform loads from CSI ETABS to CSI SAFE via their COM APIs.

## Repository Structure
- `etabs_to_safe.py` — CLI script (no GUI, prints to stdout)
- `etabs_to_safe_gui.py` — Tkinter GUI with logging, debug toggle, progress bar, log save
- `requirements.txt` — Python dependencies
- `.github/workflows/build.yml` — GitHub Actions workflow to build Windows .exe

## Key Concepts
- Both ETABS and SAFE expose COM APIs via `comtypes` on Windows
- ETABS connection: `ETABSv1.Helper` → `CSI.ETABS.API.ETABSObject` → `SapModel`
- SAFE connection: `SAFEv1.Helper` → `CSI.SAFE.API.SAFEObject` → `SapModel`
- Slab matching: ETABS slab label (`GetLabelFromName`) → SAFE slab unique name (`GetNameList`)
- Shell uniform loads: `AreaObj.GetLoadUniform` (read from ETABS), `AreaObj.SetLoadUniform` (write to SAFE)

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
