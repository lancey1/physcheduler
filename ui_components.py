import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import logging


class FileSelectionFrame(ttk.LabelFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, text="1. Select Physician Preference Files", padding="10")
        self.controller = controller
        self._setup_widgets()

    def _setup_widgets(self):
        self.file_listbox = tk.Listbox(self, height=6, selectmode=tk.EXTENDED)
        self.file_listbox.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=(0, 10))

        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        button_frame = ttk.Frame(self)
        button_frame.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Button(button_frame, text="Add Files", command=self._add_files).pack(fill=tk.X, pady=2)
        ttk.Button(button_frame, text="Remove Selected", command=self._remove_files).pack(fill=tk.X, pady=2)
        ttk.Button(button_frame, text="Clear All", command=self._clear_files).pack(fill=tk.X, pady=2)

    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="Select Physician Preference Files",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if files:
            for file in files:
                if file not in self.controller.files:
                    self.controller.files.append(file)
                    self.file_listbox.insert(tk.END, os.path.basename(file))
                    logging.info(f"Added file: {file}")

    def _remove_files(self):
        selected = self.file_listbox.curselection()
        for i in reversed(selected):
            removed_file = self.controller.files.pop(i)
            self.file_listbox.delete(i)
            logging.info(f"Removed file: {removed_file}")

    def _clear_files(self):
        self.controller.files = []
        self.file_listbox.delete(0, tk.END)
        logging.info("Cleared all files")


class OutputLocationFrame(ttk.LabelFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, text="2. Output Location", padding="10")
        self.controller = controller
        self._setup_widgets()

    def _setup_widgets(self):
        self.output_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.output_var).pack(fill=tk.X, expand=True, side=tk.LEFT, padx=(0, 10))
        ttk.Button(self, text="Browse...", command=self._browse_output).pack(side=tk.RIGHT)

    def _browse_output(self):
        output_file = filedialog.asksaveasfilename(
            title="Save Schedule As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if output_file:
            self.output_var.set(output_file)
            logging.info(f"Set output path: {output_file}")


class ControlFrame(ttk.Frame):
    """
    Button bar – workflow order (left → right):
      1. Load Files → parses employee Excel files
      2. Edit Pre-Assignments → opens the pre-assignment grid (requires files loaded)
      3. Generate Schedule → runs the solver
      4. Load Schedule → load a previously saved .xlsx schedule
      5. Save Edited Schedule → re-save after manual tweaks
    """

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self._setup_widgets()

    def _setup_widgets(self):
        # ── Step 1+2: Load files then open pre-assignment editor ──── #
        ttk.Button(
            self,
            text="1. Load & Pre-Assign",
            command=self._load_and_preassign,
            width=22,
        ).pack(side=tk.LEFT, padx=(0, 4))

        ttk.Separator(self, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6)

        # ── Step 3: Generate schedule ─────────────────────────────────── #
        ttk.Button(
            self,
            text="3. Generate Schedule",
            command=self.controller.run_scheduler,
            width=20,
        ).pack(side=tk.LEFT, padx=(0, 4))

        ttk.Separator(self, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6)

        # ── Utility buttons ───────────────────────────────────────────── #
        ttk.Button(
            self,
            text="Load Schedule",
            command=self.controller.load_schedule,
            width=14,
        ).pack(side=tk.LEFT, padx=(0, 4))

        ttk.Button(
            self,
            text="Save Edited Schedule",
            command=self.controller.save_edited_schedule,
            width=20,
        ).pack(side=tk.LEFT, padx=(0, 4))

        # Legacy "Load Pre-Assignments from file" kept for backwards-compat
        ttk.Button(
            self,
            text="Load Pre-Assignments (xlsx)",
            command=self.controller.set_pre_assignments_from_file,
            width=26,
        ).pack(side=tk.LEFT, padx=(0, 4))

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status_var).pack(side=tk.LEFT, padx=8)

    # ------------------------------------------------------------------ #
    #  Button handlers                                                   #
    # ------------------------------------------------------------------ #
    def _load_and_preassign(self):
        """Load employee files (if needed) then open the pre-assignment editor."""
        if not self.controller.files:
            messagebox.showerror("Error", "Please add physician preference files first.")
            return

        if not self.controller.employees:
            self.status_var.set("Loading employee files…")
            self.controller.root.update()
            ok = self.controller.load_employee_data_from_files()
            if not ok:
                self.status_var.set("Failed to load files – check scheduler.log")
                return
            n = len(self.controller.employees)
            self.status_var.set(f"Loaded {n} physicians.")
            self.controller.root.update()

        from pre_assignment_ui import PreAssignmentWindow
        PreAssignmentWindow(self.controller.root, self.controller)


# ====================================================================== #
#  SchedulePreviewFrame                                                   #
# ====================================================================== #

class SchedulePreviewFrame(ttk.LabelFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, text="Schedule Preview", padding="5")
        self.controller = controller
        self.schedule = {}
        self.sidebar_window = None
        self.edit_window = None
        self._setup_widgets()

    def _setup_widgets(self):
        self.toggle_button = ttk.Button(self, text="Show Shift Summary", command=self._open_sidebar)
        self.toggle_button.pack(anchor="ne", padx=2, pady=2)

        style = ttk.Style()
        style.configure(
            "Custom.Treeview.Heading",
            font=('Helvetica', 9, 'bold'),
            padding=(0, 10, 0, 10)
        )

        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=(10, 5))

        columns = ["Date", "Day"] + list(self.controller.shifts.keys()) + ["TALLY"]
        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            style="Custom.Treeview"
        )

        self.tree.column("Date", width=100, anchor=tk.CENTER, stretch=False)
        self.tree.column("Day", width=50, anchor=tk.CENTER, stretch=False)
        for shift in self.controller.shifts:
            self.tree.column(shift, width=90, anchor=tk.CENTER, stretch=False)
        self.tree.column("TALLY", width=90, anchor=tk.CENTER, stretch=False)

        self.tree.heading("Date", text="Date")
        self.tree.heading("Day", text="Day")
        for shift in self.controller.shifts:
            self.tree.heading(shift, text=f"{shift}\n({self.controller.shifts[shift]})")
        self.tree.heading("TALLY", text="TALLY")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.tree.tag_configure('unassigned', background='#ffdddd')
        self.tree.bind('<Double-Button-1>', self._on_double_click)

    def _on_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        column = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return

        col_idx = int(column.replace('#', '')) - 1
        columns = ["Date", "Day"] + list(self.controller.shifts.keys()) + ["TALLY"]

        if col_idx < 2 or col_idx >= len(columns) - 1:
            return

        shift_name = columns[col_idx]
        values = self.tree.item(row_id)['values']
        date_str = values[0]
        date_obj = datetime.strptime(date_str, '%Y-%m-%d').date()

        current_assignment = self.schedule[date_obj].get(shift_name, "UNASSIGNED")
        if current_assignment != "UNASSIGNED":
            current_assignment = current_assignment.split(" ")[0]

        self._open_edit_dialog(date_obj, shift_name, current_assignment, row_id, col_idx)

    def _open_edit_dialog(self, date_obj, shift_name, current_assignment, row_id, col_idx):
        if self.edit_window is not None and tk.Toplevel.winfo_exists(self.edit_window):
            self.edit_window.destroy()

        self.edit_window = tk.Toplevel(self)
        self.edit_window.title(f"Edit {shift_name} - {date_obj.strftime('%Y-%m-%d')}")
        self.edit_window.geometry("450x650")
        self.edit_window.transient(self)
        self.edit_window.grab_set()

        ttk.Label(
            self.edit_window,
            text=f"Assign physician for:\n{shift_name} on {date_obj.strftime('%a, %B %d, %Y')}",
            font=('Helvetica', 10, 'bold')
        ).pack(pady=10)

        ttk.Label(
            self.edit_window,
            text=f"Current: {current_assignment}",
            font=('Helvetica', 12, 'bold')
        ).pack(pady=5)

        ttk.Separator(self.edit_window, orient='horizontal').pack(fill='x', pady=10)

        legend_frame = ttk.Frame(self.edit_window)
        legend_frame.pack(pady=5)
        ttk.Label(legend_frame, text="Legend: ", font=('Helvetica', 8, 'bold')).pack(side=tk.LEFT)
        for lbl, bg in [("1", '#90EE90'), ("2/None", '#FFFF99'),
                        ("3", '#FFB6B6'), ("On Shift", '#D3D3D3')]:
            tk.Label(legend_frame, text=lbl, bg=bg,
                     font=('Helvetica', 8)).pack(side=tk.LEFT, padx=2)

        tree_frame = ttk.Frame(self.edit_window)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ("Name", "Caseload", "Preference", "On Shift")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                            selectmode="browse", height=20)
        for col, w in zip(columns, [150, 50, 50, 100]):
            tree.heading(col, text=col)
            tree.column(col, width=w, anchor=tk.CENTER)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tree.tag_configure('green', background='#90EE90')
        tree.tag_configure('yellow', background='#FFFF99')
        tree.tag_configure('red', background='#FFB6B6')
        tree.tag_configure('gray', background='#D3D3D3')
        tree.tag_configure('white', background='white')

        tree.insert("", tk.END, values=("UNASSIGNED", "", "", ""), tags=('white',))

        physicians_on_shift = {}
        for shift, assigned in self.schedule[date_obj].items():
            if shift not in ["TALLY", "CASELOAD"] and assigned != "UNASSIGNED":
                physicians_on_shift[assigned] = shift

        physicians_data = []
        for emp in self.controller.employees.keys():
            pref = self.controller.employees[emp].get('preferences', {}).get(date_obj, {}).get(shift_name)
            caseload = self.controller.PROVIDER_CASELOAD.get(emp, 0)
            is_on_shift = emp in physicians_on_shift
            pref_display = str(pref) if pref else "None"
            pref_sort = pref if pref else 2
            on_shift_display = physicians_on_shift.get(emp, "")
            physicians_data.append((emp, caseload, pref_display, on_shift_display, pref_sort, is_on_shift))

        physicians_data.sort(key=lambda x: (x[4], x[0]))

        current_item = None
        for emp, caseload, pref_display, on_shift_display, pref_sort, is_on_shift in physicians_data:
            if is_on_shift:
                tag = 'gray'
            elif pref_sort == 1:
                tag = 'green'
            elif pref_sort == 3:
                tag = 'red'
            else:
                tag = 'yellow'

            item = tree.insert("", tk.END,
                               values=(emp, caseload, pref_display, on_shift_display),
                               tags=(tag,))
            if emp == current_assignment:
                current_item = item

        if current_item:
            tree.selection_set(current_item)
            tree.see(current_item)
        else:
            tree.selection_set(tree.get_children()[0])

        button_frame = ttk.Frame(self.edit_window)
        button_frame.pack(pady=10)

        def save_assignment():
            selection = tree.selection()
            if not selection:
                return
            values = tree.item(selection[0])['values']
            new_assignment = values[0]
            self.schedule[date_obj][shift_name] = new_assignment
            self.display_schedule(self.schedule)
            self.edit_window.destroy()

        def cancel():
            self.edit_window.destroy()

        ttk.Button(button_frame, text="Save", command=save_assignment).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=cancel).pack(side=tk.LEFT, padx=5)

        tree.bind('<Double-Button-1>', lambda e: save_assignment())
        self.edit_window.bind('<Return>', lambda e: save_assignment())
        self.edit_window.bind('<Escape>', lambda e: cancel())

    def _open_sidebar(self):
        if self.sidebar_window is not None and tk.Toplevel.winfo_exists(self.sidebar_window):
            self.sidebar_window.lift()
            return

        self.sidebar_window = tk.Toplevel(self)
        self.sidebar_window.title("Shift Summary")
        self.sidebar_window.geometry("450x500")

        self.sidebar_tree = ttk.Treeview(
            self.sidebar_window,
            columns=["# Shifts Assigned", "Min Shifts"],
            show="tree headings",
            height=20
        )
        self.sidebar_tree.heading("#0", text="Doctor")
        self.sidebar_tree.column("#0", anchor=tk.W, width=180)
        self.sidebar_tree.heading("# Shifts Assigned", text="# Shifts Assigned")
        self.sidebar_tree.column("# Shifts Assigned", anchor=tk.CENTER, width=120)
        self.sidebar_tree.heading("Min Shifts", text="Min Shifts")
        self.sidebar_tree.column("Min Shifts", anchor=tk.CENTER, width=120)
        self.sidebar_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self._update_sidebar_tree()

    def _update_sidebar_tree(self):
        if self.sidebar_window is None or not tk.Toplevel.winfo_exists(self.sidebar_window):
            return

        for item in self.sidebar_tree.get_children():
            self.sidebar_tree.delete(item)

        shift_counts = {}
        shift_details = {}
        shift_type_counts = {}

        for date, shifts in self.schedule.items():
            for shift_name, doctor in shifts.items():
                if shift_name in ["TALLY", "CASELOAD"]:
                    continue
                if doctor != "UNASSIGNED":
                    shift_counts[doctor] = shift_counts.get(doctor, 0) + 1
                    shift_details.setdefault(doctor, []).append(
                        f"{date.strftime('%Y-%m-%d')} - {shift_name}"
                    )
                    shift_type_counts.setdefault(doctor, {})
                    shift_type_counts[doctor][shift_name] = \
                        shift_type_counts[doctor].get(shift_name, 0) + 1

        for doctor in sorted(self.controller.employees.keys()):
            assigned = shift_counts.get(doctor, 0)
            min_shifts = self.controller.employees[doctor]['shift_ranges'].get('min', 0)

            parent_id = self.sidebar_tree.insert(
                "", tk.END, text=doctor, values=(assigned, min_shifts), open=False
            )
            for shift_name in self.controller.shifts:
                count = shift_type_counts.get(doctor, {}).get(shift_name, 0)
                if count > 0:
                    self.sidebar_tree.insert(parent_id, tk.END,
                                             text=f"  {shift_name}: {count}",
                                             values=("", ""))
            for detail in shift_details.get(doctor, []):
                self.sidebar_tree.insert(parent_id, tk.END,
                                         text=f"  {detail}", values=("", ""))

    def display_schedule(self, schedule):
        self.schedule = schedule

        for item in self.tree.get_children():
            self.tree.delete(item)

        for date in sorted(schedule):
            day = date.strftime('%a')
            values = [date.strftime('%Y-%m-%d'), day]
            tags = []
            daily_caseload = 0

            for shift in self.controller.shifts:
                assigned = schedule[date].get(shift, "UNASSIGNED")
                if assigned != "UNASSIGNED":
                    caseload = self.controller.PROVIDER_CASELOAD.get(assigned, 0)
                    values.append(f"{assigned} ({caseload})")
                    daily_caseload += caseload
                else:
                    values.append("UNASSIGNED")
                    if shift != "Charlottetown":
                        tags.append('unassigned')

            values.append(daily_caseload)
            self.tree.insert("", tk.END, values=values, tags=tuple(tags))

        self._update_sidebar_tree()