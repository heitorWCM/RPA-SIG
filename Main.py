import os
import subprocess
import sys
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, scrolledtext
from pathlib import Path
from threading import Thread
from queue import Queue
import ctypes
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import customtkinter
import threading



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
        self.show_console = False
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
            border_color="#0F2B11",
            border_width=4,
            border_spacing=8,
            corner_radius=20,
            hover_color="#3d8e41"
        )
        run_btn.grid(row=0, column=2)
        
        # Console checkbox frame
        console_frame = tk.Frame(self.root)
        console_frame.pack(pady=(0, 5))
        
        # Show console checkbox
        self.console_var = tk.BooleanVar(value=False)
        console_cb = tk.Checkbutton(
            console_frame,
            text="Show Console Window",
            variable=self.console_var,
            font=("Arial", 9),
            command=self.toggle_console
        )
        console_cb.pack()
        
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
        select_all_var = tk.BooleanVar(value=False)
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
            var = tk.BooleanVar(value=False)
            
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
    
    def toggle_console(self):
        """Update the show_console flag"""
        self.show_console = self.console_var.get()
    
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
        return self.selected_scripts, self.selected_period, self.show_console


class FloatingConsoleWindow:
    def __init__(self, parent=None):
        if parent:
            self.root = tk.Toplevel(parent)
        else:
            self.root = tk.Toplevel()
        self.root.title("Console Output")
        
        # Set window size and position (bottom right, above progress window)
        window_width = 400
        window_height = 140
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        x = screen_width - window_width - 10
        y = screen_height - window_height - 80  # Above progress window
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.resizable(True, True)
        
        # Set icon
        try:
            icon_path = resource_path("assets/icon.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Could not set icon: {e}")

        # Load custom font properly
        try:
            # Corrected path (use forward slash or raw string)
            font_path = resource_path("assets/Nouveau_IBM_Stretch.ttf")
            
            # For Windows, you can register the font
            if os.name == 'nt':
                from ctypes import windll, byref, create_unicode_buffer, create_string_buffer
                FR_PRIVATE = 0x10
                FR_NOT_ENUM = 0x20
                
                # Add font resource
                if os.path.exists(font_path):
                    if windll.gdi32.AddFontResourceExW(
                        create_unicode_buffer(font_path), 
                        FR_PRIVATE, 
                        0
                    ):
                        # Use the actual font family name
                        # For "Retro Computer", the family name might be different
                        self.retro_font = tkfont.Font(family="Nouveau IBM Stretch", size=12)
                    else:
                        raise Exception("Could not add font resource")
                else:
                    raise FileNotFoundError(f"Font file not found: {font_path}")
            else:
                # For other platforms, tkinter might handle it differently
                self.retro_font = tkfont.Font(family="Nouveau_IBM_Stretch", size=8)
                
        except Exception as e:
            print(f"Warning: could not load custom font: {e}")
            # Fallback to monospace fonts that look retro
            self.retro_font = ("Consolas", 8)
        
        # Make window stay on top and semi-transparent
        self.root.attributes('-topmost', True)
        self.root.attributes('-alpha', 0.65)  # 95% opaque (5% transparent)
        
        # Create main frame with dark background
        main_frame = tk.Frame(self.root, bg="#1e1e1e")
        main_frame.pack(fill="both", expand=True, padx=2, pady=2)

        # Create scrolled text widget for console output
        self.console_text = scrolledtext.ScrolledText(
            main_frame,
            wrap=tk.WORD,
            font=self.retro_font,
            bg="#1e1e1e",
            #fg="#d4d4d4",
            fg="#2CFF05",
            insertbackground="#d4d4d4",
            relief=tk.FLAT,
            state=tk.DISABLED
        )
        self.console_text.pack(fill="both", expand=True)
        
        # Configure text tags for colored output
        self.console_text.tag_config("error", foreground="#f48771")
        self.console_text.tag_config("success", foreground="#4ec9b0")
        self.console_text.tag_config("info", foreground="#569cd6")
        self.console_text.tag_config("progress", foreground="#dcdcaa")
        
        # Queue for thread-safe updates
        self.update_queue = Queue()
        
        # Start checking queue
        self.check_queue()
        
        # Keep track of line count
        self.max_lines = 1000  # Maximum lines to keep in buffer
    
    def check_queue(self):
        """Check queue for updates and process them"""
        try:
            while not self.update_queue.empty():
                line, tag = self.update_queue.get_nowait()
                self._append_line(line, tag)
        except:
            pass
        
        # Schedule next check
        try:
            self.root.after(50, self.check_queue)
        except:
            pass
    
    def append_line(self, line, tag="normal"):
        """Queue a line to be added to console"""
        self.update_queue.put((line, tag))
    
    def _append_line(self, line, tag="normal"):
        """Internal method to actually append line to console"""
        self.console_text.config(state=tk.NORMAL)
        
        # Add the line with appropriate tag
        self.console_text.insert(tk.END, line + "\n", tag)
        
        # Limit buffer size
        line_count = int(self.console_text.index('end-1c').split('.')[0])
        if line_count > self.max_lines:
            self.console_text.delete('1.0', f'{line_count - self.max_lines}.0')
        
        # Auto-scroll to bottom
        self.console_text.see(tk.END)
        
        self.console_text.config(state=tk.DISABLED)
    
    def close(self):
        """Close the window"""
        try:
            self.root.destroy()
        except:
            pass


class FloatingProgressWindow:
    def __init__(self, console_window=None):
        self.console_window = console_window
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


def execute_script(script_path, progress_window=None, console_window=None, script_index=0, total_scripts=0, **kwargs):
    """
    Execute a Python script and return its status
    Can read progress markers from the script output
    """
    header = f"\n{'='*62}\nExecuting: {script_path}\n{'='*62}\n"
    print(header)
    if console_window:
        console_window.append_line(f"{'='*62}", "info")
        console_window.append_line(f"Executing: {script_path}", "info")
        console_window.append_line(f"{'='*62}", "info")
    
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
        
        # Flag to track if we're in a CarregandoDados section
        in_loading_section = False
        loading_thread = None
        loading_stop_event = threading.Event()
        
        def simulate_loading_progress():
            """Simulate progress updates during CarregandoDados"""
            step = 0
            while not loading_stop_event.is_set():
                if progress_window and total_substeps > 0:
                    # Simulate incremental progress (but don't exceed current substep)
                    simulated_step = current_substep + (step % 5) * 0.1
                    progress_window.update_progress(
                        script_index,
                        total_scripts,
                        script_path.name,
                        int(simulated_step),
                        total_substeps
                    )
                step += 1
                loading_stop_event.wait(0.5)  # Update every 500ms
        
        # Read output line by line
        for line in process.stdout:
            print(line, end='')
            sys.stdout.flush()
            
            # Add to console window
            if console_window:
                line_stripped = line.rstrip()
                if line_stripped:
                    # Determine tag based on content
                    tag = "normal"
                    if "error" in line_stripped.lower() or "✗" in line_stripped:
                        tag = "error"
                    elif "success" in line_stripped.lower() or "✓" in line_stripped:
                        tag = "success"
                    elif "PROGRESS:" in line_stripped:
                        tag = "progress"
                    elif line_stripped.startswith("="):
                        tag = "info"
                    
                    console_window.append_line(line_stripped, tag)
            
            # Detect CarregandoDados start
            if "Start searching" in line or "Loaded" in line and "loading screen images" in line:
                if not in_loading_section:
                    in_loading_section = True
                    loading_stop_event.clear()
                    loading_thread = threading.Thread(target=simulate_loading_progress, daemon=True)
                    loading_thread.start()
            
            # Detect CarregandoDados end
            if "Image(s) disappeared" in line or "No loading screen detected" in line:
                if in_loading_section:
                    in_loading_section = False
                    loading_stop_event.set()
                    if loading_thread:
                        loading_thread.join(timeout=1)
            
            # Check for progress markers in the output
            if line.strip().startswith("PROGRESS:"):
                try:
                    parts = line.strip().split(":")[1].split("/")
                    current_substep = int(parts[0])
                    total_substeps = int(parts[1])
                    
                    # Stop loading simulation if active
                    if in_loading_section:
                        in_loading_section = False
                        loading_stop_event.set()
                        if loading_thread:
                            loading_thread.join(timeout=1)
                    
                    if progress_window:
                        progress_window.update_progress(
                            script_index, 
                            total_scripts, 
                            script_path.name,
                            current_substep,
                            total_substeps
                        )
                except Exception as e:
                    error_msg = f"Error parsing progress: {e}"
                    print(error_msg, flush=True)
                    if console_window:
                        console_window.append_line(error_msg, "error")
        
        # Make sure to stop loading thread if still running
        if in_loading_section:
            loading_stop_event.set()
            if loading_thread:
                loading_thread.join(timeout=1)
        
        # Wait for process to complete
        return_code = process.wait()
        
        if return_code == 0:
            success_msg = f"✓ Successfully completed: {script_path.name}"
            print(f"\n{success_msg}")
            if console_window:
                console_window.append_line(success_msg, "success")
            sys.stdout.flush()
            return True
        else:
            stderr = process.stderr.read()
            error_msg = f"✗ Error executing {script_path.name}:\nReturn code: {return_code}"
            print(f"\n{error_msg}")
            if console_window:
                console_window.append_line(error_msg, "error")
            if stderr:
                print(f"Error: {stderr}")
                if console_window:
                    console_window.append_line(f"Error: {stderr}", "error")
            sys.stdout.flush()
            return False
        
    except Exception as e:
        error_msg = f"✗ Unexpected error executing {script_path.name}: {str(e)}"
        print(f"\n{error_msg}")
        if console_window:
            console_window.append_line(error_msg, "error")
        sys.stdout.flush()
        return False


def run_scripts(selected_scripts, selected_period, progress_window, console_window):
    """
    Execute selected scripts with progress updates
    """
    header = f"{'='*62}\nRPA Automation - Batch Executor\n{'='*62}\n"
    print(header)
    if console_window:
        console_window.append_line(f"{'='*62}", "info")
        console_window.append_line("RPA Automation - Batch Executor", "info")
        console_window.append_line(f"{'='*62}", "info")
    
    if not selected_scripts:
        msg = "No scripts selected to execute."
        print(msg)
        if console_window:
            console_window.append_line(msg, "error")
        progress_window.update_status("No scripts selected")
        return
    
    info_msg = f"Found {len(selected_scripts)} script(s) to execute:"
    print(f"{info_msg}\n")
    if console_window:
        console_window.append_line(info_msg, "info")
    
    for i, script in enumerate(selected_scripts, 1):
        script_line = f"  {i}. {script.name}"
        print(script_line)
        if console_window:
            console_window.append_line(script_line, "normal")
    
    period_msg = f"\nSelected period: {selected_period.strftime('%B/%Y')}"

    print(period_msg)
    if console_window:
        console_window.append_line(period_msg, "info")
    print()
    
    # Calculate the reference date based on selected period
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
    
    date_info = f"Date range: {initial_date} to {final_date}"
    path_info = f"Output path: {path}"
    print(f"{date_info}\n{path_info}\n")
    if console_window:
        console_window.append_line(date_info, "info")
        console_window.append_line(path_info, "info")
    
    # Execute each script with date parameters
    results = {}
    for i, script in enumerate(selected_scripts, 1):
        subfolder_name = script.parent.name
        folder_msg = f"{subfolder_name}"
        print(folder_msg)
        if console_window:
            console_window.append_line(folder_msg, "info")

        progress_window.update_progress(i-1, len(selected_scripts), script.name)
        success = execute_script(
            script, 
            progress_window,
            console_window,
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
    summary_header = f"\n{'='*62}\nExecution Summary\n{'='*62}\n"
    print(summary_header)
    if console_window:
        console_window.append_line(f"{'='*62}", "info")
        console_window.append_line("Execution Summary", "info")
        console_window.append_line(f"{'='*62}", "info")
    
    successful = sum(1 for success in results.values() if success)
    failed = len(results) - successful
    
    summary = f"Total scripts: {len(results)}\nSuccessful: {successful}\nFailed: {failed}"
    print(f"{summary}\n")
    if console_window:
        console_window.append_line(f"Total scripts: {len(results)}", "info")
        console_window.append_line(f"Successful: {successful}", "success")
        console_window.append_line(f"Failed: {failed}", "error" if failed > 0 else "info")
    
    if failed > 0:
        fail_header = "Failed scripts:"
        print(fail_header)
        if console_window:
            console_window.append_line(fail_header, "error")
        for script_name, success in results.items():
            if not success:
                fail_line = f"  ✗ {script_name}"
                print(fail_line)
                if console_window:
                    console_window.append_line(fail_line, "error")
        progress_window.update_status(f"Done: {failed} failed")
    else:
        success_msg = "All scripts completed!"
        progress_window.update_status(success_msg)
        if console_window:
            console_window.append_line(success_msg, "success")
    
    footer = f"\n{'='*62}\nBatch execution completed!\n{'='*62}"
    print(footer)
    if console_window:
        console_window.append_line(f"{'='*62}", "info")
        console_window.append_line("Batch execution completed!", "success")
        console_window.append_line(f"{'='*62}", "info")


def main():
    """
    Main function to show script selector and execute selected scripts
    """
    base_folder = "relatorios"

    # Show script selector window
    selector = ScriptSelectorWindow(base_folder)
    selected_scripts, selected_period, show_console = selector.show()
    
    # If no scripts selected, exit
    if not selected_scripts:
        print("No scripts selected. Exiting.")
        return
    
    # Create progress window (main Tk window)
    progress_window = FloatingProgressWindow()
    
    # Create console window only if checkbox is checked (as Toplevel of progress window)
    console_window = None
    if show_console:
        console_window = FloatingConsoleWindow(progress_window.root)
    
    # Run scripts in a separate thread to keep GUI responsive
    script_thread = Thread(target=run_scripts, args=(selected_scripts, selected_period, progress_window, console_window))
    script_thread.daemon = True
    script_thread.start()
    
    # Start the GUI event loop
    progress_window.root.mainloop()
    
    # Close console window when done (if it exists)
    if console_window:
        console_window.close()


if __name__ == "__main__":
    main()