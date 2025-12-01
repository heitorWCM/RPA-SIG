import os
import subprocess
import sys
import tkinter as tk
from tkinter import ttk
from pathlib import Path
from threading import Thread
from queue import Queue
import ctypes
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import customtkinter

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# Set Windows App User Model ID for taskbar icon
try:
    myappid = 'mycompany.rpa.automation.1.0'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except:
    pass

class ScriptSelectorWindow:
    def __init__(self, base_folder):
        self.base_folder = base_folder
        self.selected_scripts = []
        self.selected_period = None
        self.root = tk.Tk()
        self.root.title("RPA Run")
        
        # Set window size
        window_width = 560
        window_height = 500
        
        # Center window
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.resizable(False, False)
        
        # Set icon
        try:
            icon_path = resource_path("assets/icon.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Could not set icon: {e}")
        
        self.setup_ui()
    
    def generate_period_options(self):
        """Generate period options: last 6 months + next 2 months"""
        current_date = datetime.now()
        periods = []
        
        # Generate from 6 months ago to 2 months ahead
        for offset in range(-6, 3):  # -6 to +2
            period_date = current_date + relativedelta(months=offset)
            # Format as "MONTH_NAME/YY" in uppercase
            period_str = period_date.strftime("%B/%y").upper()
            periods.append((period_str, period_date))
        
        # Sort by date (most recent first)
        periods.sort(key=lambda x: x[1], reverse=True)
        
        return periods
        
    def setup_ui(self):
        # Top frame for button and combobox
        top_frame = tk.Frame(self.root)
        top_frame.pack(pady=10)
        
        # Period selection label and combobox
        period_label = tk.Label(
            top_frame,
            text="Period:",
            font=("Arial", 10)
        )
        period_label.grid(row=0, column=0, padx=(0, 5))
        
        # Generate period options
        self.period_options = self.generate_period_options()
        period_values = [period[0] for period in self.period_options]
        
        # Create combobox
        self.period_combo = ttk.Combobox(
            top_frame,
            values=period_values,
            state="readonly",
            width=15,
            font=("Arial", 10)
        )
        self.period_combo.grid(row=0, column=1, padx=(0, 25))
        
        # Set default to current month (should be first in sorted list)
        current_month = datetime.now().strftime("%B/%y").upper()
        if current_month in period_values:
            self.period_combo.set(current_month)
        else:
            self.period_combo.current(0)
        
        # Run Scripts button
        run_btn = customtkinter.CTkButton(
            top_frame,
            text="RUN SCRIPTS",
            command=self.on_run_scripts,
            font=("Arial", 14, "bold"),
            text_color="white",
            fg_color="#4CAF50",
            border_color="#081409",
            border_width=4,
            border_spacing=8,
            corner_radius=20,
            hover_color="#3d8e41"
        )
        run_btn.grid(row=0, column=2)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Dictionary to store checkboxes by subfolder
        self.checkboxes = {}
        self.subfolder_vars = {}
        
        # Find all subfolders and scripts
        self.script_structure = self.find_script_structure()
        
        # Create tabs for each subfolder
        for subfolder, scripts in self.script_structure.items():
            self.create_tab(subfolder, scripts)
    
    def find_script_structure(self):
        """Find all scripts organized by subfolder"""
        structure = {}
        base_path = Path(self.base_folder)
        
        if not base_path.exists():
            return structure
        
        for subfolder in sorted(base_path.iterdir()):
            if subfolder.is_dir() and not subfolder.name.startswith('__'):
                scripts = []
                for file in sorted(subfolder.rglob('*.py')):
                    if file.name != '__init__.py' and file.name != 'Main.py':
                        scripts.append(file)
                
                if scripts:
                    structure[subfolder.name] = scripts
        
        return structure
    
    def create_tab(self, subfolder_name, scripts):
        """Create a tab for a subfolder with its scripts"""
        # Create frame for tab
        tab_frame = tk.Frame(self.notebook)
        self.notebook.add(tab_frame, text=subfolder_name)
        
        # Create canvas with scrollbar
        canvas = tk.Canvas(tab_frame, bg="white")
        scrollbar = tk.Scrollbar(tab_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # Store checkboxes for this subfolder
        self.checkboxes[subfolder_name] = []
        
        # "Select All" checkbox
        select_all_var = tk.BooleanVar(value=True)
        self.subfolder_vars[subfolder_name] = select_all_var
        
        select_all_frame = tk.Frame(scrollable_frame, bg="white")
        select_all_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        select_all_cb = tk.Checkbutton(
            select_all_frame,
            text="All...",
            variable=select_all_var,
            command=lambda: self.toggle_all(subfolder_name),
            font=("Arial", 9, "bold"),
            bg="white"
        )
        select_all_cb.pack(anchor="w")
        
        # Add separator
        separator = tk.Frame(scrollable_frame, height=2, bg="lightgray")
        separator.pack(fill="x", padx=10, pady=3)
        
        # Create checkboxes for each script
        for script in scripts:
            var = tk.BooleanVar(value=True)
            
            # Extract clean script name without numbers and underscores
            script_display = script.stem.replace("_", " ").replace("-", " ")
            
            cb_frame = tk.Frame(scrollable_frame, bg="white")
            cb_frame.pack(fill="x", padx=30, pady=2)
            
            cb = tk.Checkbutton(
                cb_frame,
                text=script_display,
                variable=var,
                font=("Arial", 10),
                bg="white",
                command=lambda sn=subfolder_name: self.update_select_all(sn)
            )
            cb.pack(anchor="w")
            
            self.checkboxes[subfolder_name].append((var, script))
    
    def toggle_all(self, subfolder_name):
        """Toggle all checkboxes in a subfolder"""
        state = self.subfolder_vars[subfolder_name].get()
        for var, _ in self.checkboxes[subfolder_name]:
            var.set(state)
    
    def update_select_all(self, subfolder_name):
        """Update 'Select All' checkbox based on individual selections"""
        all_checked = all(var.get() for var, _ in self.checkboxes[subfolder_name])
        self.subfolder_vars[subfolder_name].set(all_checked)
    
    def on_run_scripts(self):
        """Collect selected scripts and period, then close window"""
        self.selected_scripts = []
        
        for subfolder_name, checkboxes in self.checkboxes.items():
            for var, script_path in checkboxes:
                if var.get():
                    self.selected_scripts.append(script_path)
        
        if not self.selected_scripts:
            tk.messagebox.showwarning(
                "No Scripts Selected",
                "Please select at least one script to run."
            )
            return
        
        # Get selected period
        selected_period_str = self.period_combo.get()
        
        # Find the corresponding date object
        for period_str, period_date in self.period_options:
            if period_str == selected_period_str:
                self.selected_period = period_date
                break
        
        self.root.destroy()
    
    def show(self):
        """Show the window and wait for user interaction"""
        self.root.mainloop()
        return self.selected_scripts, self.selected_period


class FloatingProgressWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("RPA Progress")
        
        # Set window size and position (bottom right corner)
        window_width = 250
        window_height = 70
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        x = screen_width - window_width - 10
        y = screen_height - window_height - 80
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.resizable(False, False)
        
        # Set icon using resource_path
        try:
            icon_path = resource_path("assets/icon.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Could not set icon: {e}")

        # Make window stay on top
        self.root.attributes('-topmost', True)
        
        # Remove window decorations (optional)
        self.root.overrideredirect(False)
        
        # Create labels
        self.status_label = tk.Label(
            self.root, 
            text="Initializing...", 
            font=("Arial", 9),
            anchor="w"
        )
        self.status_label.pack(fill="x", padx=10, pady=(5, 0))
        
        # Create progress bar frame
        self.progress_frame = tk.Frame(self.root, height=20, bg="lightgray")
        self.progress_frame.pack(fill="x", padx=10, pady=(5, 0))
        self.progress_frame.pack_propagate(False)
        
        # Create colored progress bar
        self.progress_bar = tk.Frame(self.progress_frame, bg="#4CAF50", height=20)
        self.progress_bar.place(x=0, y=0, relheight=1, relwidth=0)
        
        self.progress_label = tk.Label(
            self.root, 
            text="0/0 scripts", 
            font=("Arial", 8, "bold"),
            anchor="w"
        )
        self.progress_label.pack(fill="x", padx=10, pady=(2, 5))
        
        self.total_steps = 0
        self.current_step = 0
        
        # Create queue for thread-safe updates
        self.update_queue = Queue()
        
        # Start checking queue
        self.check_queue()
        
    def check_queue(self):
        """Check queue for updates and process them"""
        try:
            while not self.update_queue.empty():
                update_type, data = self.update_queue.get_nowait()
                
                if update_type == "progress":
                    current, total, script_name, substep, total_substeps = data
                    self._update_progress(current, total, script_name, substep, total_substeps)
                elif update_type == "status":
                    self._update_status(data)
        except:
            pass
        
        # Schedule next check
        self.root.after(100, self.check_queue)
    
    def update_progress(self, current, total, script_name, substep=0, total_substeps=0):
        """Queue a progress update"""
        self.update_queue.put(("progress", (current, total, script_name, substep, total_substeps)))
    
    def _update_progress(self, current, total, script_name, substep=0, total_substeps=0):
        """Internal method to actually update the progress display"""
        self.status_label.config(text=f"Running: {script_name[:25]}...")
        
        # Calculate overall progress including substeps
        if total_substeps > 0:
            progress = (current + (substep / total_substeps)) / total
            self.progress_label.config(
                text=f"{current}/{total} scripts | Step {substep}/{total_substeps}"
            )
        else:
            progress = current / total if total > 0 else 0
            self.progress_label.config(text=f"{current}/{total} scripts completed")
        
        # Update progress bar width
        self.progress_bar.place(relwidth=progress)
        self.root.update_idletasks()
    
    def update_status(self, message):
        """Queue a status update"""
        self.update_queue.put(("status", message))
    
    def _update_status(self, message):
        """Internal method to actually update status message"""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def close(self):
        """Close the window"""
        self.root.destroy()


def execute_script(script_path, progress_window=None, script_index=0, total_scripts=0, **kwargs):
    """
    Execute a Python script and return its status
    Can read progress markers from the script output
    """
    print(f"\n{'='*80}")
    print(f"Executing: {script_path}")
    print(f"{'='*80}\n")
    
    try:
        # Build command with optional arguments
        cmd = [sys.executable, '-u', str(script_path)]
        
        # Add any kwargs as command line arguments
        if 'initial_date' in kwargs and 'final_date' in kwargs:
            cmd.extend(['--initial_date', kwargs['initial_date']])
            cmd.extend(['--final_date', kwargs['final_date']])
            cmd.extend(['--path', kwargs['path']])  

        # Execute the script using subprocess with real-time output
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=1,
            universal_newlines=True
        )
        
        current_substep = 0
        total_substeps = 0
        
        # Read output line by line
        for line in process.stdout:
            print(line, end='')
            sys.stdout.flush()
            
            # Check for progress markers in the output
            if line.strip().startswith("PROGRESS:"):
                try:
                    parts = line.strip().split(":")[1].split("/")
                    current_substep = int(parts[0])
                    total_substeps = int(parts[1])
                    if progress_window:
                        progress_window.update_progress(
                            script_index, 
                            total_scripts, 
                            script_path.name,
                            current_substep,
                            total_substeps
                        )
                        print(f"[DEBUG] Progress updated: {current_substep}/{total_substeps}", flush=True)
                except Exception as e:
                    print(f"Error parsing progress: {e}", flush=True)
        
        # Wait for process to complete
        return_code = process.wait()
        
        if return_code == 0:
            print(f"\n✓ Successfully completed: {script_path.name}")
            sys.stdout.flush()
            return True
        else:
            stderr = process.stderr.read()
            print(f"\n✗ Error executing {script_path.name}:")
            print(f"Return code: {return_code}")
            if stderr:
                print(f"Error: {stderr}")
            sys.stdout.flush()
            return False
        
    except Exception as e:
        print(f"\n✗ Unexpected error executing {script_path.name}: {str(e)}")
        sys.stdout.flush()
        return False


def run_scripts(selected_scripts, selected_period, progress_window):
    """
    Execute selected scripts with progress updates
    """
    print(f"{'='*80}")
    print(f"RPA Automation - Batch Executor")
    print(f"{'='*80}\n")
    
    if not selected_scripts:
        print("No scripts selected to execute.")
        progress_window.update_status("No scripts selected")
        return
    
    print(f"Found {len(selected_scripts)} script(s) to execute:\n")
    for i, script in enumerate(selected_scripts, 1):
        print(f"  {i}. {script.name}")
    
    print(f"\nSelected period: {selected_period.strftime('%B/%Y')}")
    print()

    from modules import DeterminaDataECaminho
    
    # Calculate the reference date based on selected period
    # The logic remains the same: start_day=3 means we look at the previous month
    # So if user selects "JANUARY/25", we get data from December 2024
    year = selected_period.year
    month = selected_period.month
    
    # Create a date for the 3rd day of the selected month
    reference_date = datetime(year, month, 3)
    
    # Calculate first day of last month and last day of last month
    first_day_selected_month = reference_date.replace(day=1)
    last_day_prev_month = first_day_selected_month - timedelta(days=1)
    first_day_prev_month = last_day_prev_month.replace(day=1)
    
    # Format dates
    initial_date = first_day_prev_month.strftime("%d%m%Y")
    final_date = last_day_prev_month.strftime("%d%m%Y")
    
    # Create path with selected period
    path = os.path.join(
        r"C:\temp",
        "TesteRPA",
        reference_date.strftime("%Y"),
        f"{reference_date.strftime('%m')}.{reference_date.strftime('%b').upper()}"
    ).replace("/", "\\")
    
    # Create the folder
    os.makedirs(path, exist_ok=True)
    
    print(f"Date range: {initial_date} to {final_date}")
    print(f"Output path: {path}\n")
    
    # Execute each script with date parameters
    results = {}
    for i, script in enumerate(selected_scripts, 1):
        subfolder_name = script.parent.name
        print(f"{subfolder_name}")

        progress_window.update_progress(i-1, len(selected_scripts), script.name)
        success = execute_script(
            script, 
            progress_window, 
            i-1, 
            len(selected_scripts),
            initial_date=initial_date,
            final_date=final_date,
            path=path
        )
        results[script.name] = success
    
    # Update to show completion
    progress_window.update_progress(len(selected_scripts), len(selected_scripts), "Complete")
    
    # Print summary
    print(f"\n{'='*80}")
    print(f"Execution Summary")
    print(f"{'='*80}\n")
    
    successful = sum(1 for success in results.values() if success)
    failed = len(results) - successful
    
    print(f"Total scripts: {len(results)}")
    print(f"Successful: {successful}")
    print(f"Failed: {failed}\n")
    
    if failed > 0:
        print("Failed scripts:")
        for script_name, success in results.items():
            if not success:
                print(f"  ✗ {script_name}")
        progress_window.update_status(f"Done: {failed} failed")
    else:
        progress_window.update_status("All scripts completed!")
    
    print(f"\n{'='*80}")
    print("Batch execution completed!")
    print(f"{'='*80}")


def main():
    """
    Main function to show script selector and execute selected scripts
    """
    base_folder = "relatorios"

    # Show script selector window
    selector = ScriptSelectorWindow(base_folder)
    selected_scripts, selected_period = selector.show()
    
    # If no scripts selected, exit
    if not selected_scripts:
        print("No scripts selected. Exiting.")
        return
    
    # Create progress window
    progress_window = FloatingProgressWindow()
    
    # Run scripts in a separate thread to keep GUI responsive
    script_thread = Thread(target=run_scripts, args=(selected_scripts, selected_period, progress_window))
    script_thread.daemon = True
    script_thread.start()
    
    # Start the GUI event loop
    progress_window.root.mainloop()


if __name__ == "__main__":
    main()