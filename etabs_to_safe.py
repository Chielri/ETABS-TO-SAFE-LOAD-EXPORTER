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

    Returns a list of dicts with load details.
    """
    ret = etabs_model.AreaObj.GetLoadUniform(area_name, 0, [], [], [], [], [], 0)
    # ret layout: (NumberItems, AreaName, LoadPat, CSys, Dir, Value, retcode)
    retcode = ret[-1]
    number_items = ret[0]
    if retcode != 0 or number_items == 0:
        return []
    load_pats = ret[2]
    csys = ret[3]
    dirs = ret[4]
    values = ret[5]

    loads = []
    for i in range(number_items):
        loads.append({
            "load_pattern": load_pats[i],
            "direction": dirs[i],
            "value": values[i],
            "csys": csys[i],
        })
    return loads


def get_safe_area_names(safe_model):
    """Get all area object names in SAFE and return as a set for fast lookup."""
    ret = safe_model.AreaObj.GetNameList(0, [])
    # ret: (NumberNames, MyName, retcode)
    retcode = ret[-1]
    if retcode != 0:
        print(f"WARNING: Failed to get area names from SAFE (ret={retcode}).")
        return set()
    names = ret[1]
    name_set = set(names) if names else set()
    print(f"Found {len(name_set)} area object(s) in SAFE.")
    return name_set


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

    SetLoadUniform(Name, LoadPat, Value, Dir, Replace, CSys)
    Replace=True replaces existing load of same pattern, False adds to it.
    Returns 0 on success, non-zero on failure.
    """
    ret = safe_model.AreaObj.SetLoadUniform(
        slab_name,
        load["load_pattern"],
        load["value"],
        load["direction"],
        True,  # Replace existing load for this pattern
        load["csys"],
    )
    # COM may return a tuple; the status code is the last element
    if isinstance(ret, (tuple, list)):
        return ret[-1]
    return ret


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
