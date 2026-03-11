# ETABS to SAFE - Shell Uniform Load Exporter

Transfer shell uniform loads from selected slabs in **CSI ETABS** to matching slabs in **CSI SAFE** via their COM APIs.

## Features

- Connects to running ETABS and SAFE instances automatically
- **API connection status panel** showing ETABS/SAFE connection state, model filename, and process ID (PID)
- **PID-based connection** — connect to a specific ETABS or SAFE instance when multiple are running
- **Refresh Status** button to re-check connections (e.g. after changing the active API instance in ETABS/SAFE)
- Reads all shell uniform loads from selected slabs in ETABS
- Matches slabs by ETABS label to SAFE unique name
- **Logs slab level/story** for each selected slab during export
- Creates missing load patterns in SAFE automatically
- **CSV export report** with ETABS slab label, unique name, level, all loads, SAFE slab name, and assignment status (toggle via checkbox)
- GUI with real-time log, progress bar, debug mode, and log export
- CLI version also available for scripting

## Requirements

- Windows (COM APIs are Windows-only)
- Python 3.10+
- CSI ETABS (v18+) and CSI SAFE (v20+) installed
- `comtypes` Python package

## Installation

```bash
pip install -r requirements.txt
```

Or just:

```bash
pip install comtypes
```

## Usage

### GUI

```bash
python etabs_to_safe_gui.py
```

1. Open your model in ETABS and your model in SAFE
2. Click **Refresh Status** to verify both connections (the panel shows connection state, PID, and model filename)
3. *(Optional)* Enter a specific **PID** to target a particular ETABS or SAFE instance
4. Select the slabs in ETABS you want to export loads from
5. Click **Run Export**
6. If **CSV Report** is checked, a save dialog appears after export with the full report
7. Check the **Debug** checkbox for verbose output
8. Use **Save Log** to export the log to a file

### CLI

```bash
python etabs_to_safe.py
```

### Pre-built Executable

Download `ETABStoSAFE.exe` from [Releases](../../releases) — no Python installation needed.

## How It Works

1. Connects to ETABS via `ETABSv1.Helper` COM object
2. Connects to SAFE via `SAFEv1.Helper` COM object
3. Gets selected area objects (type 5) from ETABS
4. For each slab, reads uniform loads (`AreaObj.GetLoadUniform`)
5. Gets the ETABS slab label via `GetLabelFromName`
6. Looks up the matching slab in SAFE by unique name
7. Creates any missing load patterns in SAFE
8. Assigns loads via `AreaObj.SetLoadUniform`

## Building the Executable

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name ETABStoSAFE etabs_to_safe_gui.py
```

The `.exe` will be in the `dist/` folder.

A GitHub Actions workflow is included that builds and uploads the executable automatically on each release.

## CSV Report Columns

| Column | Description |
|---|---|
| ETABS_UniqueName | Area object unique name in ETABS |
| ETABS_Label | Slab label from `GetLabelFromName` |
| Level | Story/level the slab belongs to |
| LoadPattern | Load pattern name (e.g. Dead, Live) |
| Direction | Load direction (Gravity, Global-X, etc.) |
| Value | Load magnitude |
| CSys | Coordinate system |
| SAFE_SlabName | Matched slab name in SAFE |
| Assignment_Status | OK, FAILED, Unmatched, or No loads |

## Acknowledgements

This project relies on the following software and APIs:

- **[CSI ETABS](https://www.csiamerica.com/products/etabs)** and **[CSI SAFE](https://www.csiamerica.com/products/safe)** by Computers and Structures, Inc. (CSI) — This tool uses their public COM/API interfaces for automation. ETABS and SAFE are proprietary commercial software; a valid license for each is required. This project is not affiliated with or endorsed by CSI.

### Python Dependencies

| Package | License | Description |
|---|---|---|
| [Python](https://www.python.org/) | [PSF License](https://docs.python.org/3/license.html) | Python interpreter (3.10+) |
| [comtypes](https://github.com/enthought/comtypes) | [MIT License](https://github.com/enthought/comtypes/blob/master/LICENSE.txt) | COM interface library used to communicate with ETABS and SAFE APIs |
| [PyInstaller](https://pyinstaller.org/) | [GPL-2.0 with bootloader exception](https://github.com/pyinstaller/pyinstaller/blob/develop/COPYING.txt) | Bundles the app into a standalone `.exe`. The bootloader exception permits distribution of non-GPL applications |

### Python Standard Library (no additional license)

The following standard library modules are used and ship with Python under the PSF License: `tkinter`, `logging`, `threading`, `csv`, `os`, `subprocess`, `sys`, `traceback`, `datetime`.

### Build & CI

| Tool | License | Description |
|---|---|---|
| [GitHub Actions](https://github.com/features/actions) | GitHub Terms of Service | CI/CD for building releases |
| [actions/checkout](https://github.com/actions/checkout) | [MIT License](https://github.com/actions/checkout/blob/main/LICENSE) | Checks out the repository |
| [actions/setup-python](https://github.com/actions/setup-python) | [MIT License](https://github.com/actions/setup-python/blob/main/LICENSE) | Sets up Python in CI |
| [actions/upload-artifact](https://github.com/actions/upload-artifact) | [MIT License](https://github.com/actions/upload-artifact/blob/main/LICENSE) | Uploads build artifacts |
| [softprops/action-gh-release](https://github.com/softprops/action-gh-release) | [MIT License](https://github.com/softprops/action-gh-release/blob/master/LICENSE) | Attaches binaries to GitHub releases |

## License

[The Unlicense](LICENSE) — public domain.
