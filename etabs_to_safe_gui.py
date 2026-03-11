"""
ETABS to SAFE Shell Uniform Load Exporter - GUI Version

Tkinter GUI with logging panel, debug toggle, and export functionality.
"""

import csv
import logging
import os
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext
import sys
import traceback
from datetime import datetime

# ---------------------------------------------------------------------------
# Core logic (identical to etabs_to_safe.py but uses logging instead of print)
# ---------------------------------------------------------------------------

logger = logging.getLogger("etabs_to_safe")

DIR_NAMES = {
    1: "Local-1", 2: "Local-2", 3: "Local-3",
    4: "Global-X", 5: "Global-Y", 6: "Gravity",
    7: "Projected-X", 8: "Projected-Y", 9: "Projected-Z",
    10: "Gravity Projected", 11: "Gravity Projected",
}


def find_running_processes(exe_name):
    """Find running processes by executable name. Returns list of (pid, name) tuples."""
    try:
        output = subprocess.check_output(
            ["tasklist", "/fi", f"imagename eq {exe_name}", "/fo", "csv", "/nh"],
            text=True, stderr=subprocess.DEVNULL, creationflags=0x08000000,
        )
        results = []
        for line in output.strip().splitlines():
            if not line.strip() or "No tasks" in line or "INFO:" in line:
                continue
            parts = line.strip().strip('"').split('","')
            if len(parts) >= 2:
                try:
                    results.append((int(parts[1]), parts[0]))
                except ValueError:
                    pass
        return results
    except Exception:
        return []


def connect_to_etabs(pid=None):
    import comtypes.client
    helper = comtypes.client.CreateObject("ETABSv1.Helper")
    helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
    if pid:
        logger.info("Connecting to ETABS with PID %s...", pid)
        etabs_object = helper.GetObjectProcess("CSI.ETABS.API.ETABSObject", pid)
    else:
        etabs_object = helper.GetObject("CSI.ETABS.API.ETABSObject")
    if etabs_object is None:
        raise RuntimeError(
            "Could not connect to ETABS. Make sure ETABS is running with a model open."
        )
    sap_model = etabs_object.SapModel
    logger.info("Connected to ETABS: %s", sap_model.GetModelFilename())
    return etabs_object, sap_model


def connect_to_safe(pid=None):
    import comtypes.client
    helper = comtypes.client.CreateObject("SAFEv1.Helper")
    helper = helper.QueryInterface(comtypes.gen.SAFEv1.cHelper)
    # NOTE: SAFE reuses ETABS API infrastructure — the ProgID is "ETABSObject", not "SAFEObject"
    if pid:
        logger.info("Connecting to SAFE with PID %s...", pid)
        safe_object = helper.GetObjectProcess("CSI.SAFE.API.ETABSObject", pid)
    else:
        safe_object = helper.GetObject("CSI.SAFE.API.ETABSObject")
    if safe_object is None:
        raise RuntimeError(
            "Could not connect to SAFE. Make sure SAFE is running with a model open."
        )
    sap_model = safe_object.SapModel
    logger.info("Connected to SAFE: %s", sap_model.GetModelFilename())
    return safe_object, sap_model


def get_selected_area_names(etabs_model):
    ret = etabs_model.SelectObj.GetSelected(0, [], [])
    # ret: (NumberItems, ObjectType, ObjectName, retcode)
    retcode = ret[-1]
    if retcode != 0:
        raise RuntimeError(f"Failed to get selection from ETABS (ret={retcode}).")
    number_items, object_type, object_name = ret[0], ret[1], ret[2]
    area_names = [object_name[i] for i in range(number_items) if object_type[i] == 5]
    if not area_names:
        raise RuntimeError("No area/shell objects selected in ETABS.")
    logger.info("Found %d selected area object(s) in ETABS.", len(area_names))
    return area_names


def get_etabs_label(etabs_model, area_name):
    ret = etabs_model.AreaObj.GetLabelFromName(area_name, "", "")
    # ret: (Label, Story, retcode)
    retcode = ret[-1]
    if retcode != 0:
        logger.warning("  GetLabelFromName failed for '%s' (ret=%s).", area_name, retcode)
        return area_name, ""
    return ret[0], ret[1]


def get_shell_uniform_loads(etabs_model, area_name):
    """Get shell uniform loads — tries direct API, then database tables for Load Sets."""
    # 1) Try the standard direct API call
    try:
        ret = etabs_model.AreaObj.GetLoadUniform(area_name, 0, [], [], [], [], [], 0)
        logger.debug("  GetLoadUniform raw return: %s", ret)
        retcode = ret[-1]
        number_items = ret[0]
        if retcode == 0 and number_items > 0:
            loads = []
            for i in range(number_items):
                pat = str(ret[2][i])
                if pat.startswith("~"):
                    continue
                loads.append({
                    "load_pattern": pat,
                    "direction": int(ret[4][i]),
                    "value": float(ret[5][i]),
                    "csys": str(ret[3][i]),
                })
            if loads:
                return loads
        logger.debug("  GetLoadUniform: retcode=%s, items=%s", retcode, number_items)
    except Exception as e:
        logger.debug("  GetLoadUniform exception: %s", e)

    # 2) Try element-level query (cAreaElm) — may see loads the object-level misses
    #    NOTE: Returns one entry per mesh element, so we must deduplicate.
    try:
        ret = etabs_model.AreaElm.GetLoadUniform(area_name, 0, [], [], [], [], [], 0)
        logger.debug("  AreaElm.GetLoadUniform raw return: %s", ret)
        retcode = ret[-1]
        number_items = ret[0]
        if retcode == 0 and number_items > 0:
            seen = set()
            loads = []
            for i in range(number_items):
                pat = str(ret[2][i])
                if pat.startswith("~"):
                    continue
                direction = int(ret[4][i])
                value = float(ret[5][i])
                csys = str(ret[3][i])
                key = (pat, direction, value, csys)
                if key not in seen:
                    seen.add(key)
                    loads.append({
                        "load_pattern": pat,
                        "direction": direction,
                        "value": value,
                        "csys": csys,
                    })
            if loads:
                return loads
        logger.debug("  AreaElm.GetLoadUniform: retcode=%s, items=%s", retcode, number_items)
    except Exception as e:
        logger.debug("  AreaElm.GetLoadUniform exception: %s", e)

    # 3) Fallback: database tables (catches loads assigned via Load Sets)
    logger.debug("  Trying database tables fallback for '%s'...", area_name)
    return _get_uniform_loads_from_tables(etabs_model, area_name)


# Direction string-to-int mapping for database table values
_DIR_STR_TO_INT = {
    "gravity": 6, "grav": 6,
    "local-1": 1, "local 1": 1, "1": 1,
    "local-2": 2, "local 2": 2, "2": 2,
    "local-3": 3, "local 3": 3, "3": 3,
    "global-x": 4, "global x": 4, "x": 4,
    "global-y": 5, "global y": 5, "y": 5,
    "global-z": 6, "global z": 6, "z": 6,
    "projected-x": 7, "projected x": 7,
    "projected-y": 8, "projected y": 8,
    "projected-z": 9, "projected z": 9,
    "gravity projected": 10,
}


def _filter_internal_patterns(loads):
    """Remove internal load patterns (those starting with '~')."""
    return [ld for ld in loads if not str(ld["load_pattern"]).startswith("~")]


def _parse_direction(raw):
    """Convert a direction value (int, float-string, or descriptive string) to int."""
    if isinstance(raw, (int, float)):
        return int(raw)
    try:
        return int(float(raw))
    except (ValueError, TypeError):
        pass
    return _DIR_STR_TO_INT.get(str(raw).strip().lower(), 6)


def _find_column(fields, *candidates):
    """Find the index of a column by trying multiple candidate names (case-insensitive)."""
    lower_fields = [f.lower().strip() for f in fields]
    for c in candidates:
        cl = c.lower()
        for idx, fl in enumerate(lower_fields):
            if fl == cl:
                return idx
    # Partial match fallback
    for c in candidates:
        cl = c.lower()
        for idx, fl in enumerate(lower_fields):
            if cl in fl:
                return idx
    return None


def _read_table(db, table_name):
    """Read a database table and return (fields, num_fields, num_records, table_data) or None."""
    try:
        ret = db.GetTableForDisplayArray(table_name, [], "", 0, [], 0, [])
        if ret[-1] != 0:
            return None
        fields = list(ret[2]) if ret[2] else []
        num_records = ret[3]
        table_data = list(ret[4]) if ret[4] else []
        if not fields or num_records == 0:
            return None
        logger.debug("  Table '%s': fields=%s, records=%d", table_name, fields, num_records)
        return fields, len(fields), num_records, table_data
    except Exception as e:
        logger.debug("  Table '%s' read error: %s", table_name, e)
        return None


def _get_uniform_loads_from_tables(etabs_model, area_name):
    """Retrieve shell uniform loads via ETABS database tables API.

    Strategy:
    1) Try direct uniform load tables (have LoadPattern + Load columns).
    2) Try Load Set resolution: join the assignment table (slab -> LoadSet name)
       with the definition table (LoadSet name -> LoadPattern + LoadValue).
    """
    db = etabs_model.DatabaseTables

    # Discover candidate table names
    all_tables = []
    try:
        ret = db.GetAvailableTables(0, [], [], [])
        if ret[-1] == 0 and ret[1]:
            for t in ret[1]:
                tl = t.lower()
                if "uniform" in tl and ("area" in tl or "shell" in tl):
                    all_tables.append(t)
                elif "load set" in tl and ("area" in tl or "shell" in tl):
                    all_tables.append(t)
            logger.debug("  Discovered load tables: %s", all_tables)
    except Exception as e:
        logger.debug("  GetAvailableTables error: %s", e)

    # --- Step 1: Try direct uniform load tables ---
    for table_name in all_tables:
        tdata = _read_table(db, table_name)
        if tdata is None:
            continue
        fields, num_fields, num_records, table_data = tdata

        name_col = _find_column(fields, "UniqueName", "Unique Name", "AreaName")
        pat_col = _find_column(fields, "LoadPat", "Load Pattern", "LoadPattern", "Pattern")
        val_col = _find_column(fields, "UnifLoad", "Uniform Load", "Value", "Load")

        # Need name + pattern + value to be a direct-load table
        if name_col is None or pat_col is None or val_col is None:
            continue

        dir_col = _find_column(fields, "Dir", "Direction")
        csys_col = _find_column(fields, "CSys", "CoordSys", "Coord Sys")

        loads = []
        for row in range(num_records):
            start = row * num_fields
            row_data = table_data[start:start + num_fields]
            if len(row_data) < num_fields:
                continue
            if row_data[name_col] != area_name:
                continue
            load = {
                "load_pattern": row_data[pat_col],
                "direction": _parse_direction(row_data[dir_col]) if dir_col is not None else 6,
                "value": float(row_data[val_col]),
                "csys": row_data[csys_col] if csys_col is not None else "Global",
            }
            loads.append(load)
            logger.debug("  Direct table row match: %s", load)

        loads = [ld for ld in loads if not str(ld["load_pattern"]).startswith("~")]
        if loads:
            logger.info("  Found %d load(s) via direct table '%s'", len(loads), table_name)
            return loads

    # --- Step 2: Load Set resolution (two-table join) ---
    # Find the assignment table: maps UniqueName -> LoadSet name
    # Find the definition table: maps LoadSet Name -> LoadPattern + LoadValue
    assign_table = None
    defn_table = None
    for t in all_tables:
        tl = t.lower()
        if "load set" in tl and ("assignment" in tl or "area load" in tl):
            assign_table = t
        elif "load set" in tl and "shell" in tl and "assignment" not in tl and "area load" not in tl:
            defn_table = t

    logger.debug("  Load Set tables: assign='%s', defn='%s'", assign_table, defn_table)

    if not assign_table or not defn_table:
        return []

    # Read the assignment table to find which LoadSet(s) this slab uses
    tdata = _read_table(db, assign_table)
    if tdata is None:
        return []
    fields, num_fields, num_records, table_data = tdata

    name_col = _find_column(fields, "UniqueName", "Unique Name", "AreaName")
    set_col = _find_column(fields, "LoadSet", "Load Set")
    if name_col is None or set_col is None:
        logger.debug("  Assignment table missing UniqueName or LoadSet column")
        return []

    load_set_names = set()
    for row in range(num_records):
        start = row * num_fields
        row_data = table_data[start:start + num_fields]
        if len(row_data) < num_fields:
            continue
        if row_data[name_col] == area_name:
            load_set_names.add(row_data[set_col])

    if not load_set_names:
        logger.debug("  No Load Set assignments found for '%s'", area_name)
        return []

    logger.debug("  Slab '%s' uses Load Set(s): %s", area_name, load_set_names)

    # Read the definition table to resolve LoadSet -> LoadPattern + LoadValue
    tdata = _read_table(db, defn_table)
    if tdata is None:
        return []
    fields, num_fields, num_records, table_data = tdata

    set_name_col = _find_column(fields, "Name", "LoadSet", "Load Set")
    pat_col = _find_column(fields, "LoadPat", "Load Pattern", "LoadPattern", "Pattern")
    val_col = _find_column(fields, "LoadValue", "Load Value", "UnifLoad", "Uniform Load", "Value", "Load")
    dir_col = _find_column(fields, "Dir", "Direction")
    csys_col = _find_column(fields, "CSys", "CoordSys", "Coord Sys")

    if set_name_col is None or pat_col is None or val_col is None:
        logger.debug("  Definition table missing required columns (Name/LoadPattern/LoadValue)")
        return []

    loads = []
    for row in range(num_records):
        start = row * num_fields
        row_data = table_data[start:start + num_fields]
        if len(row_data) < num_fields:
            continue
        if row_data[set_name_col] not in load_set_names:
            continue
        load = {
            "load_pattern": row_data[pat_col],
            "direction": _parse_direction(row_data[dir_col]) if dir_col is not None else 6,
            "value": float(row_data[val_col]),
            "csys": row_data[csys_col] if csys_col is not None else "Global",
        }
        loads.append(load)
        logger.debug("  Load Set resolved: Set='%s', %s", row_data[set_name_col], load)

    loads = _filter_internal_patterns(loads)
    if loads:
        logger.info("  Found %d load(s) via Load Set tables", len(loads))
    return loads


def get_safe_area_names(safe_model):
    # Try COM AreaObj first (works in some SAFE versions via ETABS COM layer)
    try:
        ret = safe_model.AreaObj.GetNameList(0, [])
        retcode = ret[-1]
        if retcode == 0 and ret[1]:
            name_set = set(ret[1])
            logger.info("Found %d area object(s) in SAFE.", len(name_set))
            return name_set
    except Exception as e:
        logger.debug("AreaObj.GetNameList not available: %s", e)

    # Fallback: database tables (required for SAFE v22+)
    logger.debug("Trying database tables to get SAFE area names...")
    try:
        db = safe_model.DatabaseTables
        ret = db.GetTableForDisplayArray("Objects and Elements - Areas", [], "", 0, [], 0, [])
        if ret[-1] == 0 and ret[4]:
            fields = list(ret[2]) if ret[2] else []
            num_records = ret[3]
            table_data = list(ret[4])
            name_col = _find_column(fields, "UniqueName", "Unique Name", "Name")
            if name_col is not None:
                num_fields = len(fields)
                name_set = set()
                for row in range(num_records):
                    start = row * num_fields
                    if start + name_col < len(table_data):
                        name_set.add(table_data[start + name_col])
                logger.info("Found %d area object(s) in SAFE (via tables).", len(name_set))
                return name_set
    except Exception as e:
        logger.debug("SAFE database table fallback failed: %s", e)

    logger.warning("Failed to get area names from SAFE.")
    return set()


def get_existing_load_patterns(safe_model):
    ret = safe_model.LoadPatterns.GetNameList(0, [])
    # ret: (NumberNames, MyName, retcode)
    retcode = ret[-1]
    if retcode != 0:
        logger.warning("Failed to get load patterns from SAFE (ret=%s).", retcode)
        return set()
    names = ret[1]
    return set(names) if names else set()


def ensure_load_pattern_exists(safe_model, pattern_name, existing_patterns):
    if pattern_name not in existing_patterns:
        ret = safe_model.LoadPatterns.Add(pattern_name, 8, 0, True)
        retcode = ret[-1] if isinstance(ret, (tuple, list)) else ret
        if retcode == 0:
            existing_patterns.add(pattern_name)
            logger.info("  Created load pattern '%s' in SAFE.", pattern_name)
        else:
            logger.warning("  Failed to create load pattern '%s' (ret=%s).", pattern_name, retcode)
    return existing_patterns


def assign_load_to_safe(safe_model, slab_name, load):
    # Try COM AreaObj first (works in some SAFE versions via ETABS COM layer)
    try:
        ret = safe_model.AreaObj.SetLoadUniform(
            slab_name, load["load_pattern"], load["value"],
            load["direction"], True, load["csys"],
        )
        retcode = ret[-1] if isinstance(ret, (tuple, list)) else ret
        if retcode == 0:
            return 0
        logger.debug("  AreaObj.SetLoadUniform returned %s, trying database tables...", retcode)
    except Exception as e:
        logger.debug("  AreaObj.SetLoadUniform not available: %s", e)

    # Fallback: database tables (required for SAFE v22+)
    return _assign_load_via_tables(safe_model, slab_name, load)


def _assign_load_via_tables(safe_model, slab_name, load):
    """Assign a uniform load to SAFE via database tables API."""
    try:
        db = safe_model.DatabaseTables
        table_key = "Area Load Assignments - Uniform"

        # Get current table structure
        ret = db.GetTableForEditingArray(table_key, "", 0, [], 0, [])
        if ret[-1] != 0:
            logger.debug("  GetTableForEditingArray failed (ret=%s)", ret[-1])
            return ret[-1]

        table_version = ret[0]
        fields = list(ret[1]) if ret[1] else []
        num_records = ret[2]
        table_data = list(ret[3]) if ret[3] else []

        if not fields:
            logger.debug("  No fields in '%s' table", table_key)
            return -1

        num_fields = len(fields)

        # Build a new row with empty values
        new_row = [""] * num_fields
        for idx, f in enumerate(fields):
            fl = f.lower().strip()
            if fl in ("uniquename", "unique name", "name"):
                new_row[idx] = slab_name
            elif fl in ("loadpat", "load pattern", "loadpattern"):
                new_row[idx] = load["load_pattern"]
            elif fl in ("dir", "direction"):
                new_row[idx] = str(load["direction"])
            elif fl in ("unifload", "uniform load", "value"):
                new_row[idx] = str(load["value"])
            elif fl in ("csys", "coordsys", "coord sys"):
                new_row[idx] = load["csys"]

        # Append the new row
        num_records += 1
        table_data.extend(new_row)

        ret = db.SetTableForEditingArray(table_key, table_version, fields, num_records, table_data)
        if isinstance(ret, (tuple, list)):
            retcode = ret[-1]
        else:
            retcode = ret
        if retcode != 0:
            logger.debug("  SetTableForEditingArray failed (ret=%s)", retcode)
            return retcode

        ret = db.ApplyEditedTables(True, 0, 0, 0, 0, "")
        retcode = ret[-1] if isinstance(ret, (tuple, list)) else ret
        if retcode != 0:
            logger.debug("  ApplyEditedTables failed (ret=%s)", retcode)
        return retcode
    except Exception as e:
        logger.debug("  Database table load assignment failed: %s", e)
        return -1


def run_export(progress_callback=None, etabs_pid=None, safe_pid=None):
    """Main export logic. Returns a summary dict. Raises on error."""
    etabs_obj, etabs_model = connect_to_etabs(pid=etabs_pid)
    safe_obj, safe_model = connect_to_safe(pid=safe_pid)

    selected_areas = get_selected_area_names(etabs_model)
    safe_area_names = get_safe_area_names(safe_model)
    existing_patterns = get_existing_load_patterns(safe_model)
    logger.debug("Existing SAFE load patterns: %s", existing_patterns)

    matched = 0
    unmatched = 0
    loads_assigned = 0
    total = len(selected_areas)
    csv_rows = []

    for idx, area_name in enumerate(selected_areas):
        label, story = get_etabs_label(etabs_model, area_name)
        logger.info("ETABS slab: '%s' (Label: '%s', Story: '%s')", area_name, label, story)
        logger.info("  Level: %s", story if story else "N/A")

        loads = get_shell_uniform_loads(etabs_model, area_name)
        if not loads:
            logger.info("  No uniform loads assigned. Skipping.")
            csv_rows.append({
                "ETABS_UniqueName": area_name,
                "ETABS_Label": label,
                "Level": story,
                "LoadPattern": "",
                "Direction": "",
                "Value": "",
                "CSys": "",
                "SAFE_SlabName": "",
                "Assignment_Status": "No loads",
            })
            if progress_callback:
                progress_callback(idx + 1, total)
            continue

        for load in loads:
            dir_name = DIR_NAMES.get(load["direction"], f"Dir-{load['direction']}")
            logger.debug("  Load: Pattern='%s', Dir=%s, Value=%.4f, CSys='%s'",
                         load["load_pattern"], dir_name, load["value"], load["csys"])

        # Match: ETABS label -> SAFE unique name
        safe_slab_name = label
        if safe_slab_name not in safe_area_names:
            if area_name in safe_area_names:
                safe_slab_name = area_name
            else:
                logger.warning("  No matching slab in SAFE (tried '%s' and '%s'). Skipping.",
                               label, area_name)
                unmatched += 1
                for load in loads:
                    dir_name = DIR_NAMES.get(load["direction"], f"Dir-{load['direction']}")
                    csv_rows.append({
                        "ETABS_UniqueName": area_name,
                        "ETABS_Label": label,
                        "Level": story,
                        "LoadPattern": load["load_pattern"],
                        "Direction": dir_name,
                        "Value": load["value"],
                        "CSys": load["csys"],
                        "SAFE_SlabName": "",
                        "Assignment_Status": "Unmatched",
                    })
                if progress_callback:
                    progress_callback(idx + 1, total)
                continue

        logger.info("  Matched to SAFE slab: '%s'", safe_slab_name)
        matched += 1

        for load in loads:
            existing_patterns = ensure_load_pattern_exists(
                safe_model, load["load_pattern"], existing_patterns)
            ret = assign_load_to_safe(safe_model, safe_slab_name, load)
            dir_name = DIR_NAMES.get(load["direction"], f"Dir-{load['direction']}")
            if ret == 0:
                loads_assigned += 1
                logger.info("  Assigned: Pattern='%s', Value=%.4f -> OK",
                            load["load_pattern"], load["value"])
                status = "OK"
            else:
                logger.error("  FAILED: Pattern='%s' (ret=%s)", load["load_pattern"], ret)
                status = f"FAILED (ret={ret})"
            csv_rows.append({
                "ETABS_UniqueName": area_name,
                "ETABS_Label": label,
                "Level": story,
                "LoadPattern": load["load_pattern"],
                "Direction": dir_name,
                "Value": load["value"],
                "CSys": load["csys"],
                "SAFE_SlabName": safe_slab_name,
                "Assignment_Status": status,
            })

        if progress_callback:
            progress_callback(idx + 1, total)

    safe_model.View.RefreshView(0, False)

    summary = {
        "selected": total,
        "matched": matched,
        "unmatched": unmatched,
        "loads_assigned": loads_assigned,
        "csv_rows": csv_rows,
    }
    logger.info("SUMMARY: Selected=%d, Matched=%d, Unmatched=%d, Loads assigned=%d",
                total, matched, unmatched, loads_assigned)
    return summary


def write_export_csv(csv_rows, filepath):
    """Write the export report CSV file."""
    fieldnames = [
        "ETABS_UniqueName", "ETABS_Label", "Level",
        "LoadPattern", "Direction", "Value", "CSys",
        "SAFE_SlabName", "Assignment_Status",
    ]
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(csv_rows)
    logger.info("CSV report saved to %s (%d rows)", filepath, len(csv_rows))


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class TextHandler(logging.Handler):
    """Logging handler that writes to a Tkinter ScrolledText widget."""

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record) + "\n"
        self.text_widget.after(0, self._append, msg, record.levelno)

    def _append(self, msg, levelno):
        self.text_widget.configure(state="normal")
        tag = "DEBUG" if levelno <= logging.DEBUG else \
              "WARNING" if levelno == logging.WARNING else \
              "ERROR" if levelno >= logging.ERROR else "INFO"
        self.text_widget.insert(tk.END, msg, tag)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state="disabled")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ETABS to SAFE - Shell Uniform Load Exporter")
        self.geometry("780x620")
        self.resizable(True, True)
        self._running = False
        self._build_ui()
        self._setup_logging()

    # -- UI ------------------------------------------------------------------

    def _build_ui(self):
        # Top frame: buttons & options
        top = ttk.Frame(self, padding=10)
        top.pack(fill=tk.X)

        self.run_btn = ttk.Button(top, text="Run Export", command=self._on_run)
        self.run_btn.pack(side=tk.LEFT)

        self.clear_btn = ttk.Button(top, text="Clear Log", command=self._clear_log)
        self.clear_btn.pack(side=tk.LEFT, padx=(8, 0))

        self.save_btn = ttk.Button(top, text="Save Log", command=self._save_log)
        self.save_btn.pack(side=tk.LEFT, padx=(8, 0))

        self.debug_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Debug", variable=self.debug_var,
                        command=self._toggle_debug).pack(side=tk.LEFT, padx=(16, 0))

        # --- API Status Frame ---
        status_frame = ttk.LabelFrame(self, text="API Connection Status", padding=8)
        status_frame.pack(fill=tk.X, padx=10, pady=(0, 4))

        # ETABS row
        etabs_row = ttk.Frame(status_frame)
        etabs_row.pack(fill=tk.X, pady=2)
        ttk.Label(etabs_row, text="ETABS:", width=7).pack(side=tk.LEFT)
        self.etabs_status_var = tk.StringVar(value="Not checked")
        ttk.Label(etabs_row, textvariable=self.etabs_status_var,
                  width=50, anchor=tk.W).pack(side=tk.LEFT, padx=(4, 8))
        ttk.Label(etabs_row, text="PID:").pack(side=tk.LEFT)
        self.etabs_pid_var = tk.StringVar(value="")
        self.etabs_pid_entry = ttk.Entry(etabs_row, textvariable=self.etabs_pid_var, width=8)
        self.etabs_pid_entry.pack(side=tk.LEFT, padx=(2, 0))

        # SAFE row
        safe_row = ttk.Frame(status_frame)
        safe_row.pack(fill=tk.X, pady=2)
        ttk.Label(safe_row, text="SAFE:", width=7).pack(side=tk.LEFT)
        self.safe_status_var = tk.StringVar(value="Not checked")
        ttk.Label(safe_row, textvariable=self.safe_status_var,
                  width=50, anchor=tk.W).pack(side=tk.LEFT, padx=(4, 8))
        ttk.Label(safe_row, text="PID:").pack(side=tk.LEFT)
        self.safe_pid_var = tk.StringVar(value="")
        self.safe_pid_entry = ttk.Entry(safe_row, textvariable=self.safe_pid_var, width=8)
        self.safe_pid_entry.pack(side=tk.LEFT, padx=(2, 0))

        # Refresh button
        btn_row = ttk.Frame(status_frame)
        btn_row.pack(fill=tk.X, pady=(4, 0))
        self.refresh_btn = ttk.Button(btn_row, text="Refresh Status",
                                       command=self._on_refresh_status)
        self.refresh_btn.pack(side=tk.LEFT)
        ttk.Label(btn_row, text="(Leave PID empty for active instance)",
                  foreground="#888888").pack(side=tk.LEFT, padx=(8, 0))

        # Progress bar
        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill=tk.X, padx=10, pady=(0, 4))

        # Status label
        self.status_var = tk.StringVar(value="Ready. Select slabs in ETABS then click Run Export.")
        ttk.Label(self, textvariable=self.status_var).pack(fill=tk.X, padx=10)

        # Log area
        self.log_text = scrolledtext.ScrolledText(self, state="disabled", wrap=tk.WORD,
                                                   font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(4, 10))

        # Tag colours
        self.log_text.tag_config("DEBUG", foreground="#888888")
        self.log_text.tag_config("INFO", foreground="#000000")
        self.log_text.tag_config("WARNING", foreground="#CC8800")
        self.log_text.tag_config("ERROR", foreground="#CC0000")

    # -- Logging -------------------------------------------------------------

    def _setup_logging(self):
        self.text_handler = TextHandler(self.log_text)
        self.text_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                                                          datefmt="%H:%M:%S"))
        logger.addHandler(self.text_handler)
        logger.setLevel(logging.INFO)

    def _toggle_debug(self):
        level = logging.DEBUG if self.debug_var.get() else logging.INFO
        logger.setLevel(level)
        logger.info("Log level set to %s", logging.getLevelName(level))

    # -- Actions -------------------------------------------------------------

    def _get_etabs_pid(self):
        """Get ETABS PID from entry field, or None for active instance."""
        val = self.etabs_pid_var.get().strip()
        if val:
            try:
                return int(val)
            except ValueError:
                logger.warning("Invalid ETABS PID '%s', using active instance.", val)
        return None

    def _get_safe_pid(self):
        """Get SAFE PID from entry field, or None for active instance."""
        val = self.safe_pid_var.get().strip()
        if val:
            try:
                return int(val)
            except ValueError:
                logger.warning("Invalid SAFE PID '%s', using active instance.", val)
        return None

    def _on_refresh_status(self):
        if self._running:
            return
        self._running = True
        self.refresh_btn.configure(state="disabled")
        self.run_btn.configure(state="disabled")
        self.etabs_status_var.set("Checking...")
        self.safe_status_var.set("Checking...")
        threading.Thread(target=self._refresh_worker, daemon=True).start()

    def _refresh_worker(self):
        try:
            import comtypes
            comtypes.CoInitialize()
            try:
                self._check_etabs_status()
                self._check_safe_status()
            finally:
                comtypes.CoUninitialize()
        except Exception as e:
            logger.error("Status check failed: %s", e)
            logger.debug(traceback.format_exc())
        finally:
            self.after(0, self._on_refresh_done)

    def _check_etabs_status(self):
        # Show running ETABS processes
        procs = find_running_processes("ETABS.exe")
        if procs:
            pid_list = ", ".join(str(p[0]) for p in procs)
            logger.info("ETABS processes found: %s", pid_list)

        try:
            etabs_pid = self._get_etabs_pid()
            _, etabs_model = connect_to_etabs(pid=etabs_pid)
            model_file = etabs_model.GetModelFilename()
            # Find matching PID from process list
            display_pid = etabs_pid if etabs_pid else (procs[0][0] if len(procs) == 1 else None)
            pid_text = f"PID {display_pid}" if display_pid else "Active instance"
            status = f"Connected ({pid_text}) - {os.path.basename(model_file)}"
            self.after(0, self.etabs_status_var.set, status)
            if display_pid and not self.etabs_pid_var.get().strip():
                self.after(0, self.etabs_pid_var.set, str(display_pid))
        except Exception as e:
            self.after(0, self.etabs_status_var.set, f"Disconnected - {e}")
            logger.warning("ETABS connection check failed: %s", e)

    def _check_safe_status(self):
        # Show running SAFE processes
        procs = find_running_processes("SAFE.exe")
        if procs:
            pid_list = ", ".join(str(p[0]) for p in procs)
            logger.info("SAFE processes found: %s", pid_list)

        try:
            safe_pid = self._get_safe_pid()
            _, safe_model = connect_to_safe(pid=safe_pid)
            model_file = safe_model.GetModelFilename()
            display_pid = safe_pid if safe_pid else (procs[0][0] if len(procs) == 1 else None)
            pid_text = f"PID {display_pid}" if display_pid else "Active instance"
            status = f"Connected ({pid_text}) - {os.path.basename(model_file)}"
            self.after(0, self.safe_status_var.set, status)
            if display_pid and not self.safe_pid_var.get().strip():
                self.after(0, self.safe_pid_var.set, str(display_pid))
        except Exception as e:
            self.after(0, self.safe_status_var.set, f"Disconnected - {e}")
            logger.warning("SAFE connection check failed: %s", e)

    def _on_refresh_done(self):
        self._running = False
        self.refresh_btn.configure(state="normal")
        self.run_btn.configure(state="normal")

    def _on_run(self):
        if self._running:
            return
        self._running = True
        self.run_btn.configure(state="disabled")
        self.refresh_btn.configure(state="disabled")
        self.progress["value"] = 0
        self.status_var.set("Running...")
        threading.Thread(target=self._run_worker, daemon=True).start()

    def _run_worker(self):
        try:
            import comtypes
            comtypes.CoInitialize()
            try:
                etabs_pid = self._get_etabs_pid()
                safe_pid = self._get_safe_pid()
                summary = run_export(
                    progress_callback=self._update_progress,
                    etabs_pid=etabs_pid,
                    safe_pid=safe_pid,
                )
                self.after(0, self._on_done, summary)
            finally:
                comtypes.CoUninitialize()
        except Exception as e:
            logger.error("Export failed: %s", e)
            logger.debug(traceback.format_exc())
            self.after(0, self._on_error, str(e))

    def _update_progress(self, current, total):
        pct = int(current / total * 100) if total else 0
        self.after(0, self._set_progress, pct, current, total)

    def _set_progress(self, pct, current, total):
        self.progress["value"] = pct
        self.status_var.set(f"Processing slab {current}/{total}...")

    def _on_done(self, summary):
        self._running = False
        self.run_btn.configure(state="normal")
        self.refresh_btn.configure(state="normal")
        self.progress["value"] = 100

        # Auto-save CSV report
        csv_rows = summary.get("csv_rows", [])
        csv_path = ""
        if csv_rows:
            csv_path = self._save_csv(csv_rows)

        status_msg = (
            f"Done! Matched: {summary['matched']}, "
            f"Unmatched: {summary['unmatched']}, "
            f"Loads assigned: {summary['loads_assigned']}"
        )
        if csv_path:
            status_msg += f" | CSV: {os.path.basename(csv_path)}"
        self.status_var.set(status_msg)

    def _save_csv(self, csv_rows):
        """Save the CSV export report. Returns the file path or empty string."""
        from tkinter import filedialog
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"ETABS_to_SAFE_Report_{timestamp}.csv"
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=default_name,
            title="Save Export Report CSV",
        )
        if path:
            try:
                write_export_csv(csv_rows, path)
                return path
            except Exception as e:
                logger.error("Failed to save CSV: %s", e)
        return ""

    def _on_error(self, msg):
        self._running = False
        self.run_btn.configure(state="normal")
        self.refresh_btn.configure(state="normal")
        self.progress["value"] = 0
        self.status_var.set(f"Error: {msg}")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

    def _save_log(self):
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("Log files", "*.log")],
            title="Save Log As",
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.log_text.get("1.0", tk.END))
            logger.info("Log saved to %s", path)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
