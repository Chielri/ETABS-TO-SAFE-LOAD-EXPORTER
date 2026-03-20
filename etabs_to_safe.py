"""
ETABS to SAFE Shell Uniform Load Exporter

Connects to running ETABS and SAFE instances, reads shell uniform loads
from selected slabs in ETABS, matches them to slabs in SAFE by label/name,
and assigns the loads in SAFE.

Requirements:
    - ETABS and SAFE must both be running with models open
    - Select the slabs in ETABS before running this script
    - Python comtypes package: pip install comtypes
"""

import sys
import comtypes.client


def connect_to_etabs():
    """Connect to a running ETABS instance and return the SapModel."""
    try:
        helper = comtypes.client.CreateObject("ETABSv1.Helper")
        helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
        etabs_object = helper.GetObject("CSI.ETABS.API.ETABSObject")
        if etabs_object is None:
            raise RuntimeError("ETABS returned no application object.")
        sap_model = etabs_object.SapModel
        print(f"Connected to ETABS: {sap_model.GetModelFilename()}")
        return etabs_object, sap_model
    except Exception as e:
        print(f"ERROR: Could not connect to ETABS. Make sure ETABS is running.\n{e}")
        sys.exit(1)


def connect_to_safe():
    """Connect to a running SAFE instance and return the SapModel."""
    try:
        helper = comtypes.client.CreateObject("SAFEv1.Helper")
        helper = helper.QueryInterface(comtypes.gen.SAFEv1.cHelper)
        # NOTE: SAFE reuses ETABS API infrastructure — the ProgID is "ETABSObject", not "SAFEObject"
        safe_object = helper.GetObject("CSI.SAFE.API.ETABSObject")
        if safe_object is None:
            raise RuntimeError("SAFE returned no application object.")
        sap_model = safe_object.SapModel
        print(f"Connected to SAFE: {sap_model.GetModelFilename()}")
        return safe_object, sap_model
    except Exception as e:
        print(f"ERROR: Could not connect to SAFE. Make sure SAFE is running.\n{e}")
        sys.exit(1)


def get_selected_area_names(etabs_model):
    """Get names of selected area/shell objects in ETABS."""
    ret = etabs_model.SelectObj.GetSelected(0, [], [])
    # ret: (NumberItems, ObjectType, ObjectName, retcode)
    retcode = ret[-1]
    if retcode != 0:
        print(f"ERROR: Failed to get selection from ETABS (ret={retcode}).")
        sys.exit(1)
    number_items = ret[0]
    object_type = ret[1]
    object_name = ret[2]

    # Filter for area objects only (type 5)
    area_names = []
    for i in range(number_items):
        if object_type[i] == 5:
            area_names.append(object_name[i])

    if not area_names:
        print("ERROR: No area/shell objects selected in ETABS.")
        print("Please select the slabs in ETABS and run again.")
        sys.exit(1)

    print(f"Found {len(area_names)} selected area object(s) in ETABS.")
    return area_names


def get_etabs_label(etabs_model, area_name, label_cache=None):
    """Get the label and story for an area object in ETABS."""
    if label_cache is not None and area_name in label_cache:
        return label_cache[area_name]
    ret = etabs_model.AreaObj.GetLabelFromName(area_name, "", "")
    # ret: (Label, Story, retcode)
    retcode = ret[-1]
    if retcode != 0:
        print(f"  WARNING: GetLabelFromName failed for '{area_name}' (ret={retcode}).")
        return area_name, ""
    label = ret[0]
    story = ret[1]
    return label, story


def build_label_cache(etabs_model):
    """Pre-read ETABS area object labels from database tables. Returns dict: unique_name -> (label, story)."""
    cache = {}
    try:
        db = etabs_model.DatabaseTables
        for table_name in [
            "Objects and Elements - Areas",
            "Area Object Connectivity",
            "Objects - Area Objects",
            "Area Section Assignments",
        ]:
            tdata = _read_table(db, table_name)
            if tdata is None:
                continue
            fields, num_fields, num_records, table_data = tdata
            name_col = _find_column(fields, "UniqueName", "Unique Name")
            label_col = _find_column(fields, "Label")
            story_col = _find_column(fields, "Story", "Level")
            if name_col is None or label_col is None:
                continue
            for row in range(num_records):
                start = row * num_fields
                row_data = table_data[start:start + num_fields]
                if len(row_data) < num_fields:
                    continue
                uname = row_data[name_col]
                label = row_data[label_col]
                story = row_data[story_col] if story_col is not None else ""
                cache[uname] = (label, story)
            print(f"Cached {len(cache)} label(s) from '{table_name}'")
            break
    except Exception:
        pass
    return cache


def build_safe_load_cache(safe_model):
    """Pre-read SAFE existing load assignments. Returns dict: slab_name -> [pattern_names]."""
    cache = {}
    try:
        db = safe_model.DatabaseTables
        table_key = "Area Load Assignments - Uniform"
        ret = db.GetTableForDisplayArray(table_key, [], "", 0, [], 0, [])
        if ret[-1] == 0 and ret[4]:
            fields = list(ret[2]) if ret[2] else []
            num_records = ret[3]
            table_data = list(ret[4])
            num_fields = len(fields)
            name_col = _find_column(fields, "UniqueName", "Unique Name", "Name")
            pat_col = _find_column(fields, "LoadPat", "Load Pattern", "LoadPattern")
            if name_col is not None and pat_col is not None:
                for row in range(num_records):
                    start = row * num_fields
                    if start + num_fields <= len(table_data):
                        slab = table_data[start + name_col]
                        pat = table_data[start + pat_col]
                        cache.setdefault(slab, []).append(pat)
                print(f"Cached existing SAFE loads: {len(cache)} slab(s) with loads")
    except Exception:
        pass
    return cache


def get_shell_uniform_loads(etabs_model, area_name, table_cache=None):
    """Get shell uniform loads — tries table cache first (fast), then COM API fallbacks."""
    # 1) Table cache (instant lookup — preferred when available)
    if table_cache is not None:
        loads = table_cache.get(area_name, [])
        if loads:
            print(f"  Cache hit: {len(loads)} load(s) for '{area_name}'")
            return loads
        print(f"  Cache miss for '{area_name}', falling back to COM API")

    # 2) Try the standard direct API call
    print(f"  Fallback 1/2: AreaObj.GetLoadUniform (direct API)...")
    try:
        ret = etabs_model.AreaObj.GetLoadUniform(area_name, 0, [], [], [], [], [], 0)
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
                print(f"  Fallback 1/2: Found {len(loads)} load(s) via direct API.")
                return loads
        print(f"  Fallback 1/2: No loads (retcode={retcode}, items={number_items}).")
    except Exception as e:
        print(f"  Fallback 1/2: Not available ({e}).")

    # 3) Try element-level query (cAreaElm) — may see loads the object-level misses
    print(f"  Fallback 2/2: AreaElm.GetLoadUniform (element-level API)...")
    try:
        ret = etabs_model.AreaElm.GetLoadUniform(area_name, 0, [], [], [], [], [], 0)
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
                print(f"  Fallback 2/2: Found {len(loads)} load(s) via element-level API.")
                return loads
        print(f"  Fallback 2/2: No loads (retcode={retcode}, items={number_items}).")
    except Exception as e:
        print(f"  Fallback 2/2: Not available ({e}).")

    # 4) Last resort: read database tables individually (no cache available)
    if table_cache is None:
        print(f"  Last resort: Reading database tables for '{area_name}'...")
        loads = _get_uniform_loads_from_tables(etabs_model, area_name)
        if loads:
            print(f"  Found {len(loads)} load(s) from database tables.")
        else:
            print(f"  No loads found in database tables.")
        return loads

    return []


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
        return fields, len(fields), num_records, table_data
    except Exception:
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
                elif "load set" in tl:
                    all_tables.append(t)
    except Exception:
        pass

    # --- Step 1: Try direct uniform load tables ---
    for table_name in all_tables:
        tdata = _read_table(db, table_name)
        if tdata is None:
            continue
        fields, num_fields, num_records, table_data = tdata

        name_col = _find_column(fields, "UniqueName", "Unique Name", "AreaName")
        pat_col = _find_column(fields, "LoadPat", "Load Pattern", "LoadPattern", "Pattern")
        val_col = _find_column(fields, "UnifLoad", "Uniform Load", "Value", "Load")

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
            loads.append({
                "load_pattern": row_data[pat_col],
                "direction": _parse_direction(row_data[dir_col]) if dir_col is not None else 6,
                "value": float(row_data[val_col]),
                "csys": row_data[csys_col] if csys_col is not None else "Global",
            })

        loads = _filter_internal_patterns(loads)
        if loads:
            print(f"  Found {len(loads)} load(s) via database table '{table_name}'")
            return loads

    # --- Step 2: Load Set resolution (two-table join) ---
    assign_table = None
    defn_table = None
    for t in all_tables:
        tl = t.lower()
        if "load set" in tl and ("assignment" in tl or "area load" in tl):
            assign_table = t
        elif "load set" in tl and "shell" in tl and "assignment" not in tl and "area load" not in tl:
            defn_table = t

    if not assign_table or not defn_table:
        return []

    # Read assignment table: slab UniqueName -> LoadSet name
    tdata = _read_table(db, assign_table)
    if tdata is None:
        return []
    fields, num_fields, num_records, table_data = tdata

    name_col = _find_column(fields, "UniqueName", "Unique Name", "AreaName")
    set_col = _find_column(fields, "LoadSet", "Load Set")
    if name_col is None or set_col is None:
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
        return []

    # Read definition table: LoadSet Name -> LoadPattern + LoadValue
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
        return []

    loads = []
    for row in range(num_records):
        start = row * num_fields
        row_data = table_data[start:start + num_fields]
        if len(row_data) < num_fields:
            continue
        if row_data[set_name_col] not in load_set_names:
            continue
        loads.append({
            "load_pattern": row_data[pat_col],
            "direction": _parse_direction(row_data[dir_col]) if dir_col is not None else 6,
            "value": float(row_data[val_col]),
            "csys": row_data[csys_col] if csys_col is not None else "Global",
        })

    loads = _filter_internal_patterns(loads)
    if loads:
        print(f"  Found {len(loads)} load(s) via Load Set tables")
    return loads


def build_table_load_cache(etabs_model):
    """Pre-read all database tables for uniform loads and return a dict: area_name -> [loads].

    Reads the tables once and indexes by area name so each slab lookup is O(1).
    Combines both direct uniform load tables AND Load Set resolution tables.
    Returns an empty dict if tables are unavailable.
    """
    cache = {}
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
                elif "load set" in tl:
                    all_tables.append(t)
            print(f"  Discovered {len(all_tables)} candidate load tables: {all_tables}")
    except Exception:
        return cache

    # --- Step 1: Direct uniform load tables ---
    direct_count = 0
    for table_name in all_tables:
        tdata = _read_table(db, table_name)
        if tdata is None:
            continue
        fields, num_fields, num_records, table_data = tdata

        name_col = _find_column(fields, "UniqueName", "Unique Name", "AreaName")
        pat_col = _find_column(fields, "LoadPat", "Load Pattern", "LoadPattern", "Pattern")
        val_col = _find_column(fields, "UnifLoad", "Uniform Load", "Value", "Load")

        if name_col is None or pat_col is None or val_col is None:
            continue

        dir_col = _find_column(fields, "Dir", "Direction")
        csys_col = _find_column(fields, "CSys", "CoordSys", "Coord Sys")

        table_count = 0
        for row in range(num_records):
            start = row * num_fields
            row_data = table_data[start:start + num_fields]
            if len(row_data) < num_fields:
                continue
            pat = row_data[pat_col]
            if str(pat).startswith("~"):
                continue
            name = row_data[name_col]
            cache.setdefault(name, []).append({
                "load_pattern": pat,
                "direction": _parse_direction(row_data[dir_col]) if dir_col is not None else 6,
                "value": float(row_data[val_col]),
                "csys": row_data[csys_col] if csys_col is not None else "Global",
            })
            table_count += 1

        if table_count > 0:
            direct_count += table_count
            print(f"  Direct table '{table_name}': cached {table_count} load(s)")

    if direct_count > 0:
        print(f"  Step 1 total: {direct_count} direct load(s) for {len(cache)} slab(s)")

    # --- Step 2: Load Set resolution (two-table join) — ALWAYS runs ---
    assign_table = None
    defn_table = None
    for t in all_tables:
        tl = t.lower()
        if "load set" in tl and ("assignment" in tl or "area load" in tl):
            assign_table = t
        elif "load set" in tl and ("definition" in tl or "shell" in tl) \
                and "assignment" not in tl and "area load" not in tl:
            defn_table = t

    print(f"  Load Set tables: assign='{assign_table}', defn='{defn_table}'")

    if not assign_table or not defn_table:
        print("  Load Set resolution skipped (tables not found)")
    else:
        _resolve_load_sets(db, assign_table, defn_table, cache)

    total_loads = sum(len(v) for v in cache.values()) if cache else 0
    if total_loads > 0:
        print(f"Cached {total_loads} load(s) for {len(cache)} slab(s) total")

    return cache


def _resolve_load_sets(db, assign_table, defn_table, cache):
    """Resolve Load Set tables and merge results into cache. Modifies cache in place."""
    # Read assignment table
    tdata = _read_table(db, assign_table)
    if tdata is None:
        return
    fields, num_fields, num_records, table_data = tdata

    name_col = _find_column(fields, "UniqueName", "Unique Name", "AreaName")
    set_col = _find_column(fields, "LoadSet", "Load Set")
    if name_col is None or set_col is None:
        return

    area_to_sets = {}
    for row in range(num_records):
        start = row * num_fields
        row_data = table_data[start:start + num_fields]
        if len(row_data) < num_fields:
            continue
        area_to_sets.setdefault(row_data[name_col], set()).add(row_data[set_col])

    if not area_to_sets:
        return

    print(f"  Found {len(area_to_sets)} slab(s) with Load Set assignments")

    # Read definition table
    tdata = _read_table(db, defn_table)
    if tdata is None:
        return
    fields, num_fields, num_records, table_data = tdata

    set_name_col = _find_column(fields, "Name", "LoadSet", "Load Set")
    pat_col = _find_column(fields, "LoadPat", "Load Pattern", "LoadPattern", "Pattern")
    val_col = _find_column(fields, "LoadValue", "Load Value", "UnifLoad", "Uniform Load", "Value", "Load")
    dir_col = _find_column(fields, "Dir", "Direction")
    csys_col = _find_column(fields, "CSys", "CoordSys", "Coord Sys")

    if set_name_col is None or pat_col is None or val_col is None:
        return

    set_to_loads = {}
    for row in range(num_records):
        start = row * num_fields
        row_data = table_data[start:start + num_fields]
        if len(row_data) < num_fields:
            continue
        set_name = row_data[set_name_col]
        pat = row_data[pat_col]
        if str(pat).startswith("~"):
            continue
        set_to_loads.setdefault(set_name, []).append({
            "load_pattern": pat,
            "direction": _parse_direction(row_data[dir_col]) if dir_col is not None else 6,
            "value": float(row_data[val_col]),
            "csys": row_data[csys_col] if csys_col is not None else "Global",
        })

    print(f"  Found {len(set_to_loads)} Load Set definition(s): {list(set_to_loads.keys())}")

    # Join: area_name -> loads via load set names (merge into existing cache)
    loadset_count = 0
    loadset_slabs = 0
    for area_name, set_names in area_to_sets.items():
        added = False
        for set_name in set_names:
            if set_name in set_to_loads:
                cache.setdefault(area_name, []).extend(set_to_loads[set_name])
                loadset_count += len(set_to_loads[set_name])
                added = True
        if added:
            loadset_slabs += 1

    print(f"  Step 2 total: {loadset_count} load(s) for {loadset_slabs} slab(s) via Load Set resolution")


def get_safe_area_names(safe_model):
    """Get all area object names in SAFE and return as a set for fast lookup.

    Primary: COM AreaObj.GetNameList (works via ETABS COM layer inheritance).
    Fallback: database tables for SAFE versions that don't expose AreaObj.
    """
    # Try COM AreaObj first (works in some SAFE versions via ETABS COM layer)
    try:
        ret = safe_model.AreaObj.GetNameList(0, [])
        retcode = ret[-1]
        if retcode == 0 and ret[1]:
            name_set = set(ret[1])
            print(f"Found {len(name_set)} area object(s) in SAFE.")
            return name_set
    except Exception:
        pass

    # Fallback: database tables
    try:
        db = safe_model.DatabaseTables
        for table_name in [
            "Objects and Elements - Areas",
            "Area Object Connectivity",
            "Objects - Area Objects",
            "Area Section Assignments",
        ]:
            tdata = _read_table(db, table_name)
            if tdata is None:
                continue
            fields, num_fields, num_records, table_data = tdata
            name_col = _find_column(fields, "UniqueName", "Unique Name", "Name")
            if name_col is None:
                continue
            name_set = set()
            for row in range(num_records):
                start = row * num_fields
                if start + name_col < len(table_data):
                    name_set.add(table_data[start + name_col])
            print(f"Found {len(name_set)} area object(s) in SAFE (via tables).")
            return name_set
    except Exception:
        pass

    print("WARNING: Failed to get area names from SAFE.")
    return set()


def ensure_load_pattern_exists(safe_model, pattern_name, existing_patterns):
    """Create the load pattern in SAFE if it doesn't already exist. Mutates existing_patterns.

    SAFE may not expose LoadPatterns COM interface — falls back to database tables.
    """
    if pattern_name in existing_patterns:
        return
    # Try COM LoadPatterns (may work if SAFE exposes ETABS-inherited interface)
    try:
        ret = safe_model.LoadPatterns.Add(pattern_name, 8, 0, True)
        retcode = ret[-1] if isinstance(ret, (tuple, list)) else ret
        if retcode == 0:
            existing_patterns.add(pattern_name)
            print(f"  Created load pattern '{pattern_name}' in SAFE.")
            return
    except Exception:
        pass
    # Fallback: database tables
    try:
        db = safe_model.DatabaseTables
        table_key = "Load Pattern Definitions"
        ret = db.GetTableForEditingArray(table_key, "", 0, [], 0, [])
        if ret[-1] == 0:
            table_version = ret[0]
            fields = list(ret[1]) if ret[1] else []
            num_records = ret[2]
            table_data = list(ret[3]) if ret[3] else []
            if fields:
                num_fields = len(fields)
                name_col = _find_column(fields, "Name", "LoadPat", "Load Pattern")
                type_col = _find_column(fields, "Type", "LoadType", "Load Type")
                swm_col = _find_column(fields, "SelfWtMult", "Self Weight Multiplier")
                new_row = [""] * num_fields
                if name_col is not None:
                    new_row[name_col] = pattern_name
                if type_col is not None:
                    new_row[type_col] = "Other"
                if swm_col is not None:
                    new_row[swm_col] = "0"
                num_records += 1
                table_data.extend(new_row)
                ret = db.SetTableForEditingArray(table_key, table_version, fields, num_records, table_data)
                retcode = ret[-1] if isinstance(ret, (tuple, list)) else ret
                if retcode == 0:
                    ret = db.ApplyEditedTables(True, 0, 0, 0, 0, "")
                    retcode = ret[-1] if isinstance(ret, (tuple, list)) else ret
                    if retcode == 0:
                        existing_patterns.add(pattern_name)
                        print(f"  Created load pattern '{pattern_name}' in SAFE (via tables).")
                        return
    except Exception:
        pass
    print(f"  WARNING: Failed to create load pattern '{pattern_name}'.")


def get_existing_load_patterns(safe_model):
    """Get all existing load pattern names in SAFE.

    SAFE may not expose LoadPatterns COM interface — falls back to database tables.
    """
    # Try COM LoadPatterns (may work if SAFE exposes ETABS-inherited interface)
    try:
        ret = safe_model.LoadPatterns.GetNameList(0, [])
        retcode = ret[-1]
        if retcode == 0 and ret[1]:
            return set(ret[1])
    except Exception:
        pass
    # Fallback: database tables
    try:
        db = safe_model.DatabaseTables
        for table_name in ["Load Pattern Definitions", "Load Patterns"]:
            tdata = _read_table(db, table_name)
            if tdata is None:
                continue
            fields, num_fields, num_records, table_data = tdata
            name_col = _find_column(fields, "Name", "LoadPat", "Load Pattern")
            if name_col is not None:
                names = set()
                for row in range(num_records):
                    start = row * num_fields
                    if start + name_col < len(table_data):
                        names.add(table_data[start + name_col])
                print(f"Found {len(names)} load pattern(s) in SAFE.")
                return names
    except Exception:
        pass
    print("WARNING: Failed to get load patterns from SAFE.")
    return set()


def get_safe_slab_loads(safe_model, slab_name, safe_load_cache=None):
    """Get existing uniform loads on a SAFE slab. Returns list of load pattern names.

    SAFE does not expose AreaObj.GetLoadUniform — uses database tables exclusively.
    """
    if safe_load_cache is not None:
        return safe_load_cache.get(slab_name, [])

    try:
        db = safe_model.DatabaseTables
        tdata = _read_table(db, "Area Load Assignments - Uniform")
        if tdata is not None:
            fields, num_fields, num_records, table_data = tdata
            name_col = _find_column(fields, "UniqueName", "Unique Name", "Name")
            pat_col = _find_column(fields, "LoadPat", "Load Pattern", "LoadPattern")
            if name_col is not None and pat_col is not None:
                patterns = []
                for row in range(num_records):
                    start = row * num_fields
                    if start + num_fields <= len(table_data):
                        if table_data[start + name_col] == slab_name:
                            patterns.append(table_data[start + pat_col])
                return patterns
    except Exception:
        pass

    return []


def delete_safe_slab_loads(safe_model, slab_name, load_patterns):
    """Delete existing uniform loads from a SAFE slab for the given load patterns.

    SAFE does not expose AreaObj.DeleteLoadUniform — uses database tables exclusively.
    """
    return _delete_loads_via_tables(safe_model, slab_name, load_patterns)


def _delete_loads_via_tables(safe_model, slab_name, load_patterns):
    """Delete uniform loads from SAFE via database tables API (batched)."""
    try:
        db = safe_model.DatabaseTables
        table_key = "Area Load Assignments - Uniform"

        ret = db.GetTableForEditingArray(table_key, "", 0, [], 0, [])
        if ret[-1] != 0:
            return 0

        table_version = ret[0]
        fields = list(ret[1]) if ret[1] else []
        num_records = ret[2]
        table_data = list(ret[3]) if ret[3] else []

        if not fields or num_records == 0:
            return 0

        num_fields = len(fields)
        name_col = _find_column(fields, "UniqueName", "Unique Name", "Name")
        pat_col = _find_column(fields, "LoadPat", "Load Pattern", "LoadPattern")
        if name_col is None or pat_col is None:
            return 0

        patterns_to_delete = set(load_patterns)

        # Rebuild table data excluding rows matching slab_name + any pattern
        new_data = []
        new_records = 0
        deleted = 0
        for row in range(num_records):
            start = row * num_fields
            row_data = table_data[start:start + num_fields]
            if len(row_data) < num_fields:
                continue
            if row_data[name_col] == slab_name and row_data[pat_col] in patterns_to_delete:
                deleted += 1
                continue
            new_data.extend(row_data)
            new_records += 1

        if deleted == 0:
            return 0

        ret = db.SetTableForEditingArray(table_key, table_version, fields, new_records, new_data)
        retcode = ret[-1] if isinstance(ret, (tuple, list)) else ret
        if retcode != 0:
            return 0

        ret = db.ApplyEditedTables(True, 0, 0, 0, 0, "")
        retcode = ret[-1] if isinstance(ret, (tuple, list)) else ret
        return deleted if retcode == 0 else 0
    except Exception:
        return 0


def assign_load_to_safe(safe_model, slab_name, load):
    """Assign a single shell uniform load to a slab in SAFE.

    SAFE does not expose AreaObj.SetLoadUniform — uses database tables exclusively.
    Returns 0 on success, non-zero on failure.
    """
    return assign_loads_batch_to_safe(safe_model, slab_name, [load])


def assign_loads_batch_to_safe(safe_model, slab_name, loads):
    """Assign multiple shell uniform loads to a slab in SAFE in one table operation.

    Batches all loads into a single GetTable/SetTable/Apply cycle to avoid N+1.
    Returns 0 on success, non-zero on failure.
    """
    if not loads:
        return 0
    try:
        db = safe_model.DatabaseTables
        table_key = "Area Load Assignments - Uniform"

        ret = db.GetTableForEditingArray(table_key, "", 0, [], 0, [])
        if ret[-1] != 0:
            return ret[-1]

        table_version = ret[0]
        fields = list(ret[1]) if ret[1] else []
        num_records = ret[2]
        table_data = list(ret[3]) if ret[3] else []

        if not fields:
            return -1

        num_fields = len(fields)

        name_col = _find_column(fields, "UniqueName", "Unique Name", "Name")
        pat_col = _find_column(fields, "LoadPat", "Load Pattern", "LoadPattern")
        dir_col = _find_column(fields, "Dir", "Direction")
        val_col = _find_column(fields, "UnifLoad", "Uniform Load", "Value")
        csys_col = _find_column(fields, "CSys", "CoordSys", "Coord Sys")

        for load in loads:
            new_row = [""] * num_fields
            if name_col is not None:
                new_row[name_col] = slab_name
            if pat_col is not None:
                new_row[pat_col] = load["load_pattern"]
            if dir_col is not None:
                new_row[dir_col] = str(load["direction"])
            if val_col is not None:
                new_row[val_col] = str(load["value"])
            if csys_col is not None:
                new_row[csys_col] = load["csys"]
            table_data.extend(new_row)
            num_records += 1

        ret = db.SetTableForEditingArray(table_key, table_version, fields, num_records, table_data)
        retcode = ret[-1] if isinstance(ret, (tuple, list)) else ret
        if retcode != 0:
            return retcode

        ret = db.ApplyEditedTables(True, 0, 0, 0, 0, "")
        retcode = ret[-1] if isinstance(ret, (tuple, list)) else ret
        return retcode
    except Exception:
        return -1


DIR_NAMES = {
    1: "Local-1",
    2: "Local-2",
    3: "Local-3",
    4: "Global-X",
    5: "Global-Y",
    6: "Gravity",
    7: "Projected-X",
    8: "Projected-Y",
    9: "Projected-Z",
    10: "Gravity Projected",
    11: "Gravity Projected",
}


def main():
    print("=" * 60)
    print("  ETABS to SAFE - Shell Uniform Load Exporter")
    print("=" * 60)
    print()

    # Connect to both applications
    etabs_obj, etabs_model = connect_to_etabs()
    safe_obj, safe_model = connect_to_safe()
    print()

    # Get selected slabs from ETABS
    selected_areas = get_selected_area_names(etabs_model)

    # Get all area names in SAFE for matching
    safe_area_names = get_safe_area_names(safe_model)

    # Get existing load patterns in SAFE
    existing_patterns = get_existing_load_patterns(safe_model)
    print(f"Existing load patterns in SAFE: {existing_patterns}")
    print()

    # Pre-cache everything from database tables (avoids per-slab COM calls)
    print("Building caches from database tables...")
    table_cache = build_table_load_cache(etabs_model)
    label_cache = build_label_cache(etabs_model)
    safe_load_cache = build_safe_load_cache(safe_model)
    print(f"Caches ready: {sum(len(v) for v in table_cache.values()) if table_cache else 0} ETABS loads, "
          f"{len(label_cache)} labels, {len(safe_load_cache)} SAFE load entries")
    print()

    # Process each selected slab
    matched = 0
    unmatched = 0
    loads_assigned = 0

    for area_name in selected_areas:
        label, story = get_etabs_label(etabs_model, area_name, label_cache=label_cache)
        print(f"ETABS slab: '{area_name}' (Label: '{label}', Story: '{story}')")

        # Get loads from ETABS
        loads = get_shell_uniform_loads(etabs_model, area_name, table_cache=table_cache)
        if not loads:
            print(f"  No uniform loads assigned. Skipping.")
            continue

        for load in loads:
            dir_name = DIR_NAMES.get(load["direction"], f"Dir-{load['direction']}")
            print(f"  Load: Pattern='{load['load_pattern']}', "
                  f"Dir={dir_name}, Value={load['value']:.4f}")

        # Match to SAFE slab using the ETABS label as the SAFE unique name
        safe_slab_name = label
        if safe_slab_name not in safe_area_names:
            # Try with the full ETABS name as fallback
            if area_name in safe_area_names:
                safe_slab_name = area_name
            else:
                print(f"  WARNING: No matching slab found in SAFE "
                      f"(tried '{label}' and '{area_name}'). Skipping.")
                unmatched += 1
                continue

        print(f"  Matched to SAFE slab: '{safe_slab_name}'")
        matched += 1

        # Check for existing loads on SAFE slab and delete them before overwriting
        existing_safe_loads = get_safe_slab_loads(safe_model, safe_slab_name, safe_load_cache=safe_load_cache)
        if existing_safe_loads:
            unique_patterns = sorted(set(existing_safe_loads))
            print(f"  Existing loads in SAFE: {unique_patterns}")
            deleted = delete_safe_slab_loads(safe_model, safe_slab_name, unique_patterns)
            print(f"  Deleted {deleted} existing load pattern(s) from SAFE slab")
        else:
            print(f"  No existing loads on SAFE slab (clean)")

        # Ensure load patterns exist in SAFE
        for load in loads:
            ensure_load_pattern_exists(
                safe_model, load["load_pattern"], existing_patterns
            )

        # Assign all loads in one batched table operation
        ret = assign_loads_batch_to_safe(safe_model, safe_slab_name, loads)
        if ret == 0:
            loads_assigned += len(loads)
            for load in loads:
                print(f"  Assigned: Pattern='{load['load_pattern']}', "
                      f"Value={load['value']:.4f} -> OK")
        else:
            print(f"  FAILED to assign {len(loads)} load(s) to '{safe_slab_name}' "
                  f"(ret={ret})")

    # Refresh SAFE view
    safe_model.View.RefreshView(0, False)

    # Summary
    print()
    print("=" * 60)
    print(f"  SUMMARY")
    print(f"  Selected slabs in ETABS:  {len(selected_areas)}")
    print(f"  Matched to SAFE:          {matched}")
    print(f"  Unmatched:                {unmatched}")
    print(f"  Loads assigned:           {loads_assigned}")
    print("=" * 60)
    print("Done!")


if __name__ == "__main__":
    main()
