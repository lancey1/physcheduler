import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date, datetime
import logging


class PreAssignmentWindow:
    """
    A standalone Toplevel window that shows a date × shift grid.
    Each cell is a dropdown listing physicians who have a preference
    for that shift on that date (preference 1/2/3), plus an UNASSIGNED option.

    Usage:
        window = PreAssignmentWindow(parent, controller)
        # controller must expose:
        #   .employees  – dict of emp → {preferences, shift_ranges}
        #   .shifts     – dict of shift_name → start_time
        #   .pre_assignments – dict (may be empty) to read/write
    """

    PREF_COLORS = {1: "#90EE90", 2: "#FFFF99", 3: "#FFB6B6", None: "#FFFFFF"}

    def __init__(self, parent, controller):
        self.controller = controller
        self.win = tk.Toplevel(parent)
        self.win.title("Pre-Assignment Editor")
        self.win.geometry("1000x1050")
        self.win.grab_set()

        self.pre_assignments = {}
        self._cell_vars = {}
        self._combos = {}

        self._load_existing()
        self._build_ui()

    def _load_existing(self):
        existing = getattr(self.controller, "pre_assignments", {})
        for d, shifts in existing.items():
            self.pre_assignments[d] = dict(shifts)

    def _build_ui(self):
        top = ttk.Frame(self.win, padding=(8, 6))
        top.pack(fill=tk.X)

        ttk.Label(top, text="Pre-Assignment Editor", font=("Georgia", 13, "bold")).pack(side=tk.LEFT)
        ttk.Button(top, text="✓  Save & Close", command=self._save_and_close).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="✗  Clear All", command=self._clear_all).pack(side=tk.RIGHT, padx=4)

        legend = ttk.Frame(top)
        legend.pack(side=tk.RIGHT, padx=16)
        for label, color in [("Pref 1", "#90EE90"), ("Pref 2", "#FFFF99"),
                              ("Pref 3", "#FFB6B6"), ("No pref", "#FFFFFF")]:
            tk.Label(legend, text=f" {label} ", bg=color,
                     relief="solid", bd=1, font=("Helvetica", 8)).pack(side=tk.LEFT, padx=2)

        ttk.Separator(self.win, orient="horizontal").pack(fill=tk.X)

        canvas_frame = ttk.Frame(self.win)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_frame, bg="#f4f4f4", highlightthickness=0)
        self.vsb = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.inner = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Bind mousewheel only to canvas itself, not globally
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.inner.bind("<MouseWheel>", self._on_mousewheel)

        self._populate_grid()

        self.status_var = tk.StringVar(value="Use the dropdowns to assign physicians.")
        ttk.Label(self.win, textvariable=self.status_var,
                  relief="sunken", anchor="w", padding=(6, 2)).pack(fill=tk.X, side=tk.BOTTOM)

    def _populate_grid(self):
        shifts = [s for s in self.controller.shifts if s != "Charlottetown"]
        employees = self.controller.employees
        dates = sorted(set(
            d for emp_data in employees.values()
            for d in emp_data.get("preferences", {}).keys()
        ))

        if not dates:
            ttk.Label(self.inner, text="No employee data loaded yet.\nPlease load physician Excel files first.",
                      font=("Helvetica", 11), foreground="gray").grid(row=0, column=0, padx=30, pady=30)
            return

        # Header row
        corner = ttk.Label(self.inner, text="Date", font=("Helvetica", 9, "bold"),
                           relief="ridge", padding=(3, 2), anchor="center")
        corner.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        corner.bind("<MouseWheel>", self._on_mousewheel)

        for col, shift in enumerate(shifts, start=1):
            header = ttk.Label(self.inner, text=shift, font=("Helvetica", 8, "bold"),
                               relief="ridge", padding=(2, 1), anchor="center", width=12)
            header.grid(row=0, column=col, sticky="nsew", padx=1, pady=1)
            header.bind("<MouseWheel>", self._on_mousewheel)

        # Data rows
        for row, date_obj in enumerate(dates, start=1):
            bg = "#dde0f0" if date_obj.weekday() >= 5 else "#f0f0f0"
            date_lbl = tk.Label(self.inner,
                                text=f"{date_obj.strftime('%a')}  {date_obj.strftime('%Y-%m-%d')}",
                                font=("Courier", 8, "bold"),
                                bg=bg, relief="ridge", padx=4, pady=2, anchor="w", width=16)
            date_lbl.grid(row=row, column=0, sticky="nsew", padx=1, pady=1)
            date_lbl.bind("<MouseWheel>", self._on_mousewheel)

            for col, shift in enumerate(shifts, start=1):
                self._build_cell(row, col, date_obj, shift, employees)

        for col in range(len(shifts) + 1):
            self.inner.columnconfigure(col, weight=1)

    def _build_cell(self, row, col, date_obj, shift_name, employees):
        options_with_pref = []
        for emp, emp_data in employees.items():
            pref = emp_data.get("preferences", {}).get(date_obj, {}).get(shift_name)
            if pref in {1, 2, 3}:
                options_with_pref.append((emp, pref))

        options_with_pref.sort(key=lambda x: (x[1], x[0]))
        option_names = ["UNASSIGNED"] + [emp for emp, _ in options_with_pref]
        pref_map = {emp: pref for emp, pref in options_with_pref}

        existing = self.pre_assignments.get(date_obj, {}).get(shift_name, "UNASSIGNED")
        if existing not in option_names:
            existing = "UNASSIGNED"

        var = tk.StringVar(value=existing)
        self._cell_vars[(date_obj, shift_name)] = var

        cell_frame = tk.Frame(self.inner, relief="ridge", bd=1)
        cell_frame.grid(row=row, column=col, sticky="nsew", padx=1, pady=1)
        cell_frame.bind("<MouseWheel>", self._on_mousewheel)

        def get_bg(v):
            if v == "UNASSIGNED":
                return "#f0f0f0"
            return self.PREF_COLORS.get(pref_map.get(v), "#ffffff")

        band = tk.Frame(cell_frame, height=4, bg=get_bg(existing))
        band.pack(fill=tk.X)
        band.bind("<MouseWheel>", self._on_mousewheel)

        combo = ttk.Combobox(cell_frame, textvariable=var, values=option_names,
                             state="readonly", width=12, font=("Helvetica", 8))
        combo.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        self._combos[(date_obj, shift_name)] = combo

        # Block scroll on combobox — redirect to canvas
        combo.bind("<MouseWheel>", self._on_mousewheel_break)

        def on_change(*args, _band=band, _var=var, _pmap=pref_map):
            _band.config(bg=get_bg(_var.get()))
            self._update_status()

        var.trace_add("write", on_change)

        if not options_with_pref:
            combo.config(state="disabled", foreground="gray")

    def _save_and_close(self):
        new_pa = {}
        for (date_obj, shift_name), var in self._cell_vars.items():
            val = var.get()
            if val and val != "UNASSIGNED":
                new_pa.setdefault(date_obj, {})[shift_name] = val

        self.controller.pre_assignments = new_pa
        total = sum(len(v) for v in new_pa.values())
        logging.info(f"Pre-assignments saved: {total} assignments across {len(new_pa)} dates")
        self.status_var.set(f"Saved {total} pre-assignments.")
        messagebox.showinfo("Saved", f"{total} pre-assignments saved.\nClose this window and click Generate Schedule.", parent=self.win)
        self.win.destroy()

    def _clear_all(self):
        if not messagebox.askyesno("Clear All", "Clear all pre-assignments?", parent=self.win):
            return
        for var in self._cell_vars.values():
            var.set("UNASSIGNED")
        self.controller.pre_assignments = {}
        self.status_var.set("All pre-assignments cleared.")

    def _update_status(self):
        count = sum(1 for var in self._cell_vars.values() if var.get() != "UNASSIGNED")
        self.status_var.set(f"{count} pre-assignment(s) set.")

    def _on_inner_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _needs_scroll(self):
        """Return True only if content is taller than the visible canvas area."""
        self.canvas.update_idletasks()
        top, bottom = self.canvas.yview()
        return not (top == 0.0 and bottom == 1.0)

    def _on_mousewheel(self, event):
        if self._needs_scroll():
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_break(self, event):
        """Used on comboboxes: scroll canvas if needed, always block combo from changing."""
        if self._needs_scroll():
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"