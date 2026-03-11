# ETABS to SAFE - Shell Uniform Load Exporter

Transfer shell uniform loads from selected slabs in **CSI ETABS** to matching slabs in **CSI SAFE** via their COM APIs.

## Features

- Connects to running ETABS and SAFE instances automatically
- Reads all shell uniform loads from selected slabs in ETABS
- Matches slabs by ETABS label to SAFE unique name
- Creates missing load patterns in SAFE automatically
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
2. Select the slabs in ETABS you want to export loads from
3. Click **Run Export**
4. Check the **Debug** checkbox for verbose output
5. Use **Save Log** to export the log to a file

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

## License

[The Unlicense](LICENSE) — public domain.
