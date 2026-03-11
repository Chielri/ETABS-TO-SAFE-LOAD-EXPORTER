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


def get_etabs_label(etabs_model, area_name):
    """Get the label and story for an area object in ETABS."""
    ret = etabs_model.AreaObj.GetLabelFromName(area_name, "", "")
    # ret: (Label, Story, retcode)
    retcode = ret[-1]
    if retcode != 0:
        print(f"  WARNING: GetLabelFromName failed for '{area_name}' (ret={retcode}).")
        return area_name, ""
    label = ret[0]
    story = ret[1]
    return label, story


def get_shell_uniform_loads(etabs_model, area_name):
    """Get all shell uniform loads assigned to an area object in ETABS.

    Tries direct API first, then falls back to database tables for Load Sets.
    Returns a list of dicts with load details.
    """
    # 1) Try the standard direct API call
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
                return loads
    except Exception as e:
        print(f"  DEBUG: GetLoadUniform exception: {e}")

    # 2) Try element-level query (cAreaElm) — may see loads the object-level misses
    #    NOTE: Returns one entry per mesh element, so we must deduplicate.
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
                return loads
    except Exception:
        pass

    # 3) Fallback: database tables (catches loads assigned via Load Sets)
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


def get_safe_area_names(safe_model):
    """Get all area object names in SAFE and return as a set for fast lookup."""
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

    # Fallback: database tables (required for SAFE v22+)
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
                print(f"Found {len(name_set)} area object(s) in SAFE (via tables).")
                return name_set
    except Exception:
        pass

    print("WARNING: Failed to get area names from SAFE.")
    return set()


def ensure_load_pattern_exists(safe_model, pattern_name, existing_patterns):
    """Create the load pattern in SAFE if it doesn't already exist."""
    if pattern_name not in existing_patterns:
        # Type 8 = Other (generic). SelfWtMultiplier = 0. AddLoadCase = True.
        ret = safe_model.LoadPatterns.Add(pattern_name, 8, 0, True)
        retcode = ret[-1] if isinstance(ret, (tuple, list)) else ret
        if retcode == 0:
            existing_patterns.add(pattern_name)
            print(f"  Created load pattern '{pattern_name}' in SAFE.")
        else:
            print(f"  WARNING: Failed to create load pattern '{pattern_name}' (ret={retcode}).")
    return existing_patterns


def get_existing_load_patterns(safe_model):
    """Get all existing load pattern names in SAFE."""
    ret = safe_model.LoadPatterns.GetNameList(0, [])
    # ret: (NumberNames, MyName, retcode)
    retcode = ret[-1]
    if retcode != 0:
        print(f"WARNING: Failed to get load patterns from SAFE (ret={retcode}).")
        return set()
    names = ret[1]
    return set(names) if names else set()


def assign_load_to_safe(safe_model, slab_name, load):
    """Assign a single shell uniform load to a slab in SAFE.

    Tries COM AreaObj first, then falls back to database tables.
    Returns 0 on success, non-zero on failure.
    """
    # Try COM AreaObj first (works in some SAFE versions via ETABS COM layer)
    try:
        ret = safe_model.AreaObj.SetLoadUniform(
            slab_name,
            load["load_pattern"],
            load["value"],
            load["direction"],
            True,
            load["csys"],
        )
        retcode = ret[-1] if isinstance(ret, (tuple, list)) else ret
        if retcode == 0:
            return 0
    except Exception:
        pass

    # Fallback: database tables (required for SAFE v22+)
    return _assign_load_via_tables(safe_model, slab_name, load)


def _assign_load_via_tables(safe_model, slab_name, load):
    """Assign a uniform load to SAFE via database tables API."""
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

        num_records += 1
        table_data.extend(new_row)

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

    # Process each selected slab
    matched = 0
    unmatched = 0
    loads_assigned = 0

    for area_name in selected_areas:
        label, story = get_etabs_label(etabs_model, area_name)
        print(f"ETABS slab: '{area_name}' (Label: '{label}', Story: '{story}')")

        # Get loads from ETABS
        loads = get_shell_uniform_loads(etabs_model, area_name)
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

        # Ensure load patterns exist in SAFE and assign loads
        for load in loads:
            existing_patterns = ensure_load_pattern_exists(
                safe_model, load["load_pattern"], existing_patterns
            )
            ret = assign_load_to_safe(safe_model, safe_slab_name, load)
            if ret == 0:
                loads_assigned += 1
                print(f"  Assigned: Pattern='{load['load_pattern']}', "
                      f"Value={load['value']:.4f} -> OK")
            else:
                print(f"  FAILED to assign: Pattern='{load['load_pattern']}' "
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
