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

## Quick Start (for the Pre-built .exe)

If you just want to run the tool without installing Python:

1. Download **`ETABStoSAFE.exe`** from the [Releases](../../releases) page
2. Open your structural model in **ETABS** and your foundation model in **SAFE**
3. In ETABS, **select the slabs** you want to transfer loads from
4. Double-click `ETABStoSAFE.exe` — the GUI opens
5. Click **Run Export** — loads are transferred and you'll see results in the log

No Python installation needed.

---

## Usage — GUI (Recommended)

The GUI version provides a visual interface with real-time feedback. This is the recommended way to use the tool.

```bash
python etabs_to_safe_gui.py
```

### Step-by-Step

1. **Open both programs** — Launch ETABS with your structural model and SAFE with your foundation/slab model. Both must have their models open and visible.

2. **Launch the tool** — Run `ETABStoSAFE.exe` or `python etabs_to_safe_gui.py`. The main window appears.

3. **Check connections** — Click **Refresh Status**. The status panel shows whether ETABS and SAFE are connected, the model filename, and the process ID (PID).
   - If you have **multiple ETABS or SAFE instances** open, enter the PID of the one you want in the PID field, then click Refresh Status again. You can find the PID in Windows Task Manager (Details tab).
   - If you only have one instance of each, leave PID fields empty — it connects automatically.

4. **Select slabs in ETABS** — Switch to ETABS and select the slabs (floor/shell objects) whose loads you want to export. You can select one slab, multiple slabs, or an entire floor.

5. **Click Run Export** — The tool:
   - Reads all shell uniform loads from each selected slab in ETABS
   - Shows the **slab label, level/story, and all assigned loads** in the log
   - Matches each ETABS slab to the corresponding slab in SAFE by label name
   - Creates any load patterns in SAFE that don't already exist (e.g. Dead, Live, SDL)
   - Assigns the loads to the matched slabs in SAFE
   - Shows a progress bar and summary when done

6. **Review the results** — The log panel shows every slab processed, its level, which loads were assigned, and any warnings (e.g. unmatched slabs). The status bar shows a final summary.

7. **Save the CSV report** *(optional)* — If the **CSV Report** checkbox is checked, a Save dialog appears after export. The CSV file contains a complete record of every slab and load processed — useful for documentation or QA review.

8. **Save the log** *(optional)* — Click **Save Log** to save the full text log to a `.txt` file.

### GUI Options

| Control | What it does |
|---|---|
| **Run Export** | Starts the load transfer process |
| **Clear Log** | Clears the log panel |
| **Save Log** | Saves the log text to a file |
| **Debug** | Shows detailed/verbose messages in the log (for troubleshooting) |
| **CSV Report** | When checked, prompts to save a CSV summary after export |
| **Refresh Status** | Re-checks ETABS and SAFE connections and updates the status panel |
| **PID fields** | Enter a specific process ID to connect to a particular ETABS/SAFE instance |

---

## Usage — CLI (Command Line)

The CLI version runs entirely in the terminal with no graphical interface. It prints results directly to the console. This is useful for scripting or when you prefer a text-based workflow.

```bash
python etabs_to_safe.py
```

### Step-by-Step

1. **Open both programs** — Same as GUI: open ETABS with your structural model and SAFE with your foundation model.

2. **Select slabs in ETABS** — In ETABS, select the slabs you want to export loads from. This must be done **before** running the script.

3. **Open a Command Prompt** — Press `Win + R`, type `cmd`, press Enter. Navigate to the folder where the script is located:
   ```
   cd C:\path\to\ETABS-TO-SAFE-LOAD-EXPORTER
   ```

4. **Run the script**:
   ```
   python etabs_to_safe.py
   ```

5. **Read the output** — The script prints:
   - Connection status for ETABS and SAFE
   - For each selected slab: label, story, and all uniform loads found
   - Whether each slab was matched in SAFE
   - Whether each load was assigned successfully
   - A final summary with counts

### Example CLI Output

```
============================================================
  ETABS to SAFE - Shell Uniform Load Exporter
============================================================

Connected to ETABS: MyBuilding.EDB
Connected to SAFE: MyFoundation.fdb

Found 3 selected area object(s) in ETABS.
Found 45 area object(s) in SAFE.
Existing load patterns in SAFE: {'Dead', 'Live'}

ETABS slab: 'F1' (Label: 'F1', Story: 'Level 1')
  Load: Pattern='Dead', Dir=Gravity, Value=2.5000
  Load: Pattern='Live', Dir=Gravity, Value=3.0000
  Matched to SAFE slab: 'F1'
  Assigned: Pattern='Dead', Value=2.5000 -> OK
  Assigned: Pattern='Live', Value=3.0000 -> OK

ETABS slab: 'F2' (Label: 'F2', Story: 'Level 1')
  Load: Pattern='Dead', Dir=Gravity, Value=2.5000
  Matched to SAFE slab: 'F2'
  Assigned: Pattern='Dead', Value=2.5000 -> OK

ETABS slab: 'F3' (Label: 'F3', Story: 'Level 2')
  No uniform loads assigned. Skipping.

============================================================
  SUMMARY
  Selected slabs in ETABS:  3
  Matched to SAFE:          2
  Unmatched:                0
  Loads assigned:           3
============================================================
Done!
```

### CLI vs GUI — Which to Use?

| | GUI | CLI |
|---|---|---|
| Best for | Day-to-day use, visual feedback | Scripting, batch workflows |
| Progress bar | Yes | No |
| CSV report | Yes (toggle via checkbox) | No |
| PID selection | Yes (entry fields) | No (uses active instance) |
| Log export | Yes (Save Log button) | Copy/paste from terminal |
| Debug mode | Yes (checkbox) | No |

---

## Slab Matching — How ETABS Slabs Map to SAFE Slabs

Understanding how the tool matches slabs is important for getting correct results:

1. The tool reads each selected slab's **label** from ETABS (via `GetLabelFromName`) — for example, a slab with unique name `"4"` on Story `"Level 1"` might have label `"F1"`.
2. It looks for a slab in SAFE with that **same name** (`"F1"`).
3. If no match is found by label, it tries the ETABS **unique name** (`"4"`) as a fallback.
4. If neither matches, the slab is reported as **Unmatched** and skipped.

**Tip:** For best results, make sure your SAFE slab names match your ETABS slab labels. You can check slab labels in ETABS under `Assign > Shell > Labels`.

## Troubleshooting

| Problem | Solution |
|---|---|
| **"Could not connect to ETABS"** | Make sure ETABS is running and a model is open (not just the start screen). If you have multiple ETABS instances, enter the correct PID. |
| **"Could not connect to SAFE"** | Same as above but for SAFE. Make sure the SAFE model file (`.fdb`) is open. |
| **"No area/shell objects selected in ETABS"** | Go to ETABS and select at least one slab/shell before running. Use click, box select, or `Select > Group` to select slabs. |
| **Slabs show as "Unmatched"** | The slab label in ETABS doesn't match any slab name in SAFE. Check that naming is consistent between both models. See "Slab Matching" above. |
| **Loads show as "FAILED"** | The load could not be assigned in SAFE. Check that the slab exists and is not locked. Enable Debug mode for more detail. |
| **Wrong ETABS/SAFE instance connected** | If multiple instances are open, use **Refresh Status** and enter the correct PID in the GUI. Find PIDs in Windows Task Manager > Details tab. In ETABS/SAFE you can also go to `Tools > Active Instance for API` to set which instance responds. |
| **Script says `comtypes` not found** | Run `pip install comtypes` in Command Prompt. Not needed if using the `.exe` version. |

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
