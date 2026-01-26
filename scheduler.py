import tkinter as tk
from tkinter import ttk, filedialog
from tkinter import messagebox
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import logging
from schedule_generator import ScheduleGenerator
from ui_components import FileSelectionFrame, OutputLocationFrame, ControlFrame
# NOTE: SchedulePreviewFrame is updated below

class ShiftSchedulerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Physician Shift Scheduler")
        self.root.geometry("1100x1000")
        
        # Initialize data structures
        self.shifts = {
            "St. John's": "06:00",
            "Charlottetown": "08:00",
            "Halifax": "09:30",
            "Montreal": "11:30",
            "Ottawa": "14:01",
            "Toronto": "16:01",
            "Edmonton": "18:01",
            "Vancouver": "20:01",
            "Victoria": "23:01"
        }
        
        self.PROVIDER_CASELOAD = {
        "Adamson": 40,
        "Chan": 30,
        "Charnish": 18,
        "Da Qi": 50,
        "Duerksen": 45,
        "Duic": 25,
        "Feng": 55,
        "Gannage": 40,
        "Grenier": 55,
        "Ho": 45,
        "Hui": 45,
        "Huynh": 20,
        "Ikonnikov": 45,
        "iMak": 45,
        "Kanhai": 45,
        "Lau": 45,
        "Lee": 55,
        "Leung": 65,
        "Puri": 110,
        "Shin": 35,
        "Shum": 25,
        "Silverstein": 35,
        "Sommer": 55,
        "Shankar": 25,
        "Tan": 35,
        "Van Heer": 30,
        "Waghmare": 40,
        "Yue": 30,
        }

        self.employees = {}
        self.dates = set()
        self.files = []
        self.schedule = {}  # will hold final editable schedule
        
        # Setup logging
        self.setup_logging()
        
        # Create GUI
        self.create_widgets()
    
    def setup_logging(self):
        logging.basicConfig(
            filename='scheduler.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        console = logging.StreamHandler()
        console.setLevel(logging.DEBUG)
        logging.getLogger().addHandler(console)
    
    def create_widgets(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create UI components
        from ui_components import SchedulePreviewFrame  # import updated version
        self.file_frame = FileSelectionFrame(main_frame, self)
        self.file_frame.pack(fill=tk.X, pady=5)
        
        self.output_frame = OutputLocationFrame(main_frame, self)
        self.output_frame.pack(fill=tk.X, pady=5)
        
        self.control_frame = ControlFrame(main_frame, self)
        self.control_frame.pack(fill=tk.X, pady=5)
        
        self.results_frame = SchedulePreviewFrame(main_frame, self)
        self.results_frame.pack(fill=tk.BOTH, expand=True)
    
    def run_scheduler(self):
        if not self.files:
            messagebox.showerror("Error", "Please select at least one input file")
            return
        
        output_path = self.output_frame.output_var.get()
        if not output_path:
            messagebox.showerror("Error", "Please specify an output file location")
            return
        
        try:
            self.control_frame.status_var.set("Generating schedule...")
            self.root.update()
            
            # Initialize with empty data
            generator = ScheduleGenerator({}, set(), list(self.shifts.keys()))
            
            # Process each file
            for file in self.files:
                if not generator.process_employee_file(file):
                    logging.warning(f"Failed to process file: {file}")
                    continue
            
            # Now we have populated employees and dates
            self.employees = generator.employees
            self.dates = set(generator.dates)
            
            if not self.validate_data():
                return
            
            # Reinitialize generator with loaded data
            generator = ScheduleGenerator(self.employees, self.dates, list(self.shifts.keys()))
            
            self.control_frame.status_var.set("Generating schedule...")
            self.root.update()
            
            self.schedule, stats = generator.generate_schedule()
            
            # Display results (editable preview)
            self.results_frame.display_schedule(self.schedule)
            
            # Save immediately after generation
            self.save_schedule_to_excel(self.schedule, output_path)
            messagebox.showinfo("Success", f"Schedule saved to:\n{output_path}")
            
        except Exception as e:
            logging.exception("Scheduler failed")
            self.control_frame.status_var.set("Error occurred")
            messagebox.showerror("Error", f"Failed to generate schedule:\n{str(e)}")
    
    def validate_data(self):
        if not self.employees:
            messagebox.showerror("Error", "No physician data found in files")
            return False
        if not self.dates:
            messagebox.showerror("Error", "No valid dates found in files")
            return False
        return True
    
    def save_schedule_to_excel(self, schedule, output_path):
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Schedule"
            
            headers = ["Date", "Day"] + list(self.shifts.keys()) + ["TALLY"]
            ws.append(headers)
            
            header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
            assigned_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
            unassigned_fill = PatternFill(start_color="FFCCCB", end_color="FFCCCB", fill_type="solid")
            
            for cell in ws[1]:
                cell.fill = header_fill
            
            for date in sorted(schedule):
                row = [date, date.strftime('%a')]
                daily_tally = 0
                for shift in self.shifts:
                    assigned = schedule[date].get(shift, "UNASSIGNED")
                    if assigned != "UNASSIGNED":
                        caseload = self.PROVIDER_CASELOAD.get(assigned, 0)
                        row.append(f"{assigned} ({caseload})")
                        daily_tally += caseload
                    else:
                        row.append(assigned)
                row.append(daily_tally)
                ws.append(row)
                
                last_row = ws.max_row
                ws.cell(row=last_row, column=1).number_format = 'YYYY-MM-DD'
                
                for col_idx in range(3, len(headers)+1):
                    cell = ws.cell(row=last_row, column=col_idx)
                    if "UNASSIGNED" in str(cell.value):
                        cell.fill = unassigned_fill
                    else:
                        cell.fill = assigned_fill
            
            for col in ws.columns:
                max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                ws.column_dimensions[col[0].column_letter].width = (max_length + 2) * 1.2
            
            wb.save(output_path)
            logging.info(f"Schedule saved to {output_path}")
        
        except Exception as e:
            logging.error(f"Failed to save Excel file: {str(e)}")
            raise

    def save_edited_schedule(self):
        """Save the currently edited schedule to Excel"""
        if not self.schedule:
            messagebox.showwarning("Warning", "No schedule to save. Please generate a schedule first.")
            return
        
        output_path = self.output_frame.output_var.get()
        if not output_path:
            messagebox.showerror("Error", "Please specify an output file location")
            return
        
        try:
            self.save_schedule_to_excel(self.schedule, output_path)
            self.control_frame.status_var.set("Edited schedule saved")
            messagebox.showinfo("Success", f"Edited schedule saved to:\n{output_path}")
        except Exception as e:
            logging.error(f"Failed to save edited schedule: {str(e)}")
            messagebox.showerror("Error", f"Failed to save edited schedule:\n{str(e)}")

    def load_schedule(self):
        """Load a previously saved schedule from Excel"""
        file_path = filedialog.askopenfilename(
            title="Load Existing Schedule",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if self.files and not self.employees:
            self.load_employee_data_from_files()
        
        try:
            from openpyxl import load_workbook
            from datetime import datetime
            
            self.control_frame.status_var.set("Loading schedule...")
            self.root.update()
            
            wb = load_workbook(file_path, data_only=True)
            ws = wb.active
            
            # Clear existing schedule
            self.schedule = {}
            
            # Read headers to get shift names (row 1)
            headers = []
            for cell in ws[1]:
                if cell.value:
                    headers.append(cell.value)
            
            # Find shift columns (skip Date and Day, stop before TALLY)
            shift_columns = []
            for i, header in enumerate(headers):
                if header not in ["Date", "Day", "TALLY"]:
                    shift_columns.append((i, header))
            
            # Read data rows (starting from row 2)
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]:  # Skip empty rows
                    continue
                
                # Parse date
                date_value = row[0]
                if isinstance(date_value, datetime):
                    date_obj = date_value.date()
                elif isinstance(date_value, str):
                    try:
                        date_obj = datetime.strptime(date_value, '%Y-%m-%d').date()
                    except:
                        logging.warning(f"Could not parse date: {date_value}")
                        continue
                else:
                    continue
                
                # Initialize this date in schedule
                self.schedule[date_obj] = {}
                
                # Read shift assignments
                for col_idx, shift_name in shift_columns:
                    cell_value = row[col_idx] if col_idx < len(row) else None
                    
                    if cell_value and isinstance(cell_value, str):
                        # Extract physician name (remove caseload info)
                        if "(" in cell_value:
                            physician = cell_value.split("(")[0].strip()
                        else:
                            physician = cell_value.strip()
                        
                        self.schedule[date_obj][shift_name] = physician
                    else:
                        self.schedule[date_obj][shift_name] = "UNASSIGNED"
            
            wb.close()
            
            # Set output path to the loaded file for easy re-saving
            self.output_frame.output_var.set(file_path)
            
            # Display the loaded schedule
            self.results_frame.display_schedule(self.schedule)
            
            self.control_frame.status_var.set("Schedule loaded successfully")
            messagebox.showinfo("Success", f"Schedule loaded from:\n{file_path}")
            
            logging.info(f"Loaded schedule from {file_path}")
            
        except Exception as e:
            logging.exception("Failed to load schedule")
            self.control_frame.status_var.set("Error loading schedule")
            messagebox.showerror("Error", f"Failed to load schedule:\n{str(e)}")

    def load_employee_data_from_files(self):
        """Load employee data without generating a new schedule"""
        if not self.files:
            return False
        
        try:
            from schedule_generator import ScheduleGenerator
            
            # Initialize with empty data
            generator = ScheduleGenerator({}, set(), list(self.shifts.keys()))
            
            # Process each file to get employee preferences
            for file in self.files:
                if not generator.process_employee_file(file):
                    logging.warning(f"Failed to process file: {file}")
                    continue
            
            # Store employee data
            self.employees = generator.employees
            self.dates = set(generator.dates)
            
            return True
        except Exception as e:
            logging.error(f"Failed to load employee data: {str(e)}")
            return False









if __name__ == "__main__":
    import tkinter as tk
    root = tk.Tk()
    app = ShiftSchedulerApp(root)
    root.mainloop()