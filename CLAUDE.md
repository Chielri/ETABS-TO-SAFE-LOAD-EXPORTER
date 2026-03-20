# CLAUDE.md

## Project Overview
ETABS to SAFE Shell Uniform Load Exporter — a Python tool that transfers shell uniform loads from CSI ETABS to CSI SAFE via their COM APIs.

## Repository Structure
- `etabs_to_safe.py` — CLI script (no GUI, prints to stdout)
- `etabs_to_safe_gui.py` — Tkinter GUI with logging, debug toggle, progress bar, log save
- `CSI API ETABS v1.chm` — Official ETABS API reference (CHM help file)
- `CSI API SAFE v1.chm` — Official SAFE API reference (CHM help file)
- `etabs-api.skill` — ETABS API reference (ZIP archive, extract to read)
- `safe-api.skill` — SAFE API reference (ZIP archive, extract to read)
- `requirements.txt` — Python dependencies
- `.github/workflows/build.yml` — GitHub Actions workflow to build Windows .exe

## Key Concepts
- Both ETABS and SAFE expose COM APIs via `comtypes` on Windows
- ETABS connection: `ETABSv1.Helper` → `CSI.ETABS.API.ETABSObject` → `SapModel`
- SAFE connection: `SAFEv1.Helper` → `CSI.SAFE.API.ETABSObject` → `SapModel`
  - **CRITICAL**: SAFE reuses ETABS API infrastructure. The ProgID is `ETABSObject`, NOT `SAFEObject`.
- ETABS exposes rich COM interfaces: `cAreaObj`, `cAreaElm`, `cLoadPatterns`, `cSelect`, `cDatabaseTables`, etc.
- SAFE has exactly 14 interfaces (per CHM): `cAnalysisResults`, `cAnalysisResultsSetup`, `cAnalyze`, `cDatabaseTables`, `cDesignCompositeBeam`, `cDesignConcrete`, `cDesignConcreteSlab`, `cDesignSteel`, `cFile`, `cHelper`, `cOAPI`, `cSapModel`, `cSelect`, `cView`
  - **SAFE does NOT expose `cAreaObj`, `cLoadPatterns`, `cAreaElm`** — these are ETABS-only
  - All SAFE model manipulation (loads, geometry, properties) goes through `cDatabaseTables`
  - Key table: `"Area Load Assignments - Uniform"` for reading/writing uniform area loads
- Slab matching: ETABS slab label (`GetLabelFromName`) → SAFE slab unique name (via database tables)
- Shell uniform loads: `AreaObj.GetLoadUniform` (read from ETABS), database tables (write to SAFE)

## API Reference
- Authoritative source: `CSI API ETABS v1.chm` and `CSI API SAFE v1.chm` (extract with `extract_chmLib`)
- Supplementary: `etabs-api.skill` and `safe-api.skill` (ZIP archives, extract with `unzip`)
- COM return values (comtypes): `[OutParam1, OutParam2, ..., ret]` — ref params become outputs, return code is LAST
- SAFE database table data is a flat 1D string array, row-by-row (N fields × M records)
- All API methods return int: 0 = success, nonzero = failure
- Key ETABS signatures (verified from CHM):
  - `cAreaObj.GetLoadUniform(Name, ref NumberItems, ref AreaName[], ref LoadPat[], ref CSys[], ref Dir[], ref Value[], ItemType)`
  - `cAreaObj.GetLabelFromName(Name, ref Label, ref Story)`
  - `cSelect.GetSelected(ref NumberItems, ref ObjectType[], ref ObjectName[])`
  - `cLoadPatterns.Add(Name, LoadPatternType, SelfWTMultiplier, AddToLoadCombination)`
- Key SAFE signatures (verified from CHM):
  - `cDatabaseTables.GetTableForDisplayArray(TableKey, ref FieldKeyList[], GroupName, ref TableVersion, ref FieldsKeysIncluded[], ref NumberRecords, ref TableData[])`
  - `cDatabaseTables.GetTableForEditingArray(TableKey, GroupName, ref TableVersion, ref FieldsKeysIncluded[], ref NumberRecords, ref TableData[])`
  - `cDatabaseTables.SetTableForEditingArray(TableKey, ref TableVersion, ref FieldsKeysIncluded[], NumberRecords, ref TableData[])`
  - `cDatabaseTables.ApplyEditedTables(FillImportLog, ref NumFatalErrors, ref NumErrorMsgs, ref NumWarnMsgs, ref NumInfoMsgs, ref ImportLog)`

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
