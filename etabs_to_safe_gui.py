"""
ETABS to SAFE Shell Uniform Load Exporter - GUI Version

Tkinter GUI with logging panel, debug toggle, and export functionality.
"""

import logging
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext
import sys
import traceback

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


def connect_to_etabs():
    import comtypes.client
    helper = comtypes.client.CreateObject("ETABSv1.Helper")
    helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
    etabs_object = helper.GetObject("CSI.ETABS.API.ETABSObject")
    if etabs_object is None:
        raise RuntimeError(
            "Could not connect to ETABS. Make sure ETABS is running with a model open."
        )
    sap_model = etabs_object.SapModel
    logger.info("Connected to ETABS: %s", sap_model.GetModelFilename())
    return etabs_object, sap_model


def connect_to_safe():
    import comtypes.client
    helper = comtypes.client.CreateObject("SAFEv1.Helper")
    helper = helper.QueryInterface(comtypes.gen.SAFEv1.cHelper)
    # NOTE: SAFE reuses ETABS API infrastructure — the ProgID is "ETABSObject", not "SAFEObject"
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
    ret = etabs_model.AreaObj.GetLoadUniform(area_name, 0, [], [], [], [], [], 0)
    retcode = ret[-1]
    number_items = ret[0]
    if retcode != 0 or number_items == 0:
        return []
    loads = []
    for i in range(number_items):
        loads.append({
            "load_pattern": ret[2][i],
            "direction": ret[4][i],
            "value": ret[5][i],
            "csys": ret[3][i],
        })
    return loads


def get_safe_area_names(safe_model):
    ret = safe_model.AreaObj.GetNameList(0, [])
    # ret: (NumberNames, MyName, retcode)
    retcode = ret[-1]
    if retcode != 0:
        logger.warning("Failed to get area names from SAFE (ret=%s).", retcode)
        return set()
    names = ret[1]
    name_set = set(names) if names else set()
    logger.info("Found %d area object(s) in SAFE.", len(name_set))
    return name_set


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
    ret = safe_model.AreaObj.SetLoadUniform(
        slab_name, load["load_pattern"], load["value"],
        load["direction"], True, load["csys"],
    )
    if isinstance(ret, (tuple, list)):
        return ret[-1]
    return ret


def run_export(progress_callback=None):
    """Main export logic. Returns a summary dict. Raises on error."""
    etabs_obj, etabs_model = connect_to_etabs()
    safe_obj, safe_model = connect_to_safe()

    selected_areas = get_selected_area_names(etabs_model)
    safe_area_names = get_safe_area_names(safe_model)
    existing_patterns = get_existing_load_patterns(safe_model)
    logger.debug("Existing SAFE load patterns: %s", existing_patterns)

    matched = 0
    unmatched = 0
    loads_assigned = 0
    total = len(selected_areas)

    for idx, area_name in enumerate(selected_areas):
        label, story = get_etabs_label(etabs_model, area_name)
        logger.info("ETABS slab: '%s' (Label: '%s', Story: '%s')", area_name, label, story)

        loads = get_shell_uniform_loads(etabs_model, area_name)
        if not loads:
            logger.info("  No uniform loads assigned. Skipping.")
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
                if progress_callback:
                    progress_callback(idx + 1, total)
                continue

        logger.info("  Matched to SAFE slab: '%s'", safe_slab_name)
        matched += 1

        for load in loads:
            existing_patterns = ensure_load_pattern_exists(
                safe_model, load["load_pattern"], existing_patterns)
            ret = assign_load_to_safe(safe_model, safe_slab_name, load)
            if ret == 0:
                loads_assigned += 1
                logger.info("  Assigned: Pattern='%s', Value=%.4f -> OK",
                            load["load_pattern"], load["value"])
            else:
                logger.error("  FAILED: Pattern='%s' (ret=%s)", load["load_pattern"], ret)

        if progress_callback:
            progress_callback(idx + 1, total)

    safe_model.View.RefreshView(0, False)

    summary = {
        "selected": total,
        "matched": matched,
        "unmatched": unmatched,
        "loads_assigned": loads_assigned,
    }
    logger.info("SUMMARY: %s", summary)
    return summary


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
        self.geometry("750x520")
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

    def _on_run(self):
        if self._running:
            return
        self._running = True
        self.run_btn.configure(state="disabled")
        self.progress["value"] = 0
        self.status_var.set("Running...")
        threading.Thread(target=self._run_worker, daemon=True).start()

    def _run_worker(self):
        try:
            import comtypes
            comtypes.CoInitialize()
            try:
                summary = run_export(progress_callback=self._update_progress)
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
        self.progress["value"] = 100
        self.status_var.set(
            f"Done! Matched: {summary['matched']}, "
            f"Unmatched: {summary['unmatched']}, "
            f"Loads assigned: {summary['loads_assigned']}"
        )

    def _on_error(self, msg):
        self._running = False
        self.run_btn.configure(state="normal")
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
