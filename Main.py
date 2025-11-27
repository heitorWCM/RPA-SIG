import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from threading import Thread

from modules.DateFolder import DeterminaDataECaminho

class FloatingProgressWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("RPA Progress")
        
        # Set window size and position (bottom right corner)
        window_width = 250
        window_height = 70  # Increased to fit progress bar
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        x = screen_width - window_width - 10
        y = screen_height - window_height - 80
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.resizable(False, False)
        
        self.root.iconbitmap("assets/icon.ico")

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
        
    def update_progress(self, current, total, script_name, substep=0, total_substeps=0):
        """Update the progress display"""
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
        self.root.update()
    
    def update_status(self, message):
        """Update status message"""
        self.status_label.config(text=message)
        self.root.update()
    
    def close(self):
        """Close the window"""
        self.root.destroy()

def find_python_scripts(base_folder):
    """
    Find all Python scripts in subfolders of the base folder,
    excluding __init__.py and Main.py
    """
    scripts = []
    base_path = Path(base_folder)
    
    # Walk through all subdirectories
    for subfolder in base_path.iterdir():
        if subfolder.is_dir() and not subfolder.name.startswith('__'):
            # Look for Python files in each subfolder
            for file in subfolder.rglob('*.py'):
                # Exclude __init__.py files
                if file.name != '__init__.py' and file.name != 'Main.py':
                    scripts.append(file)
    
    return sorted(scripts)

def execute_script(script_path, progress_window=None, script_index=0, total_scripts=0):
    """
    Execute a Python script and return its status
    Can read progress markers from the script output
    """
    print(f"\n{'='*80}")
    print(f"Executing: {script_path}")
    print(f"{'='*80}\n")
    
    try:
        # Execute the script using subprocess with real-time output
        process = subprocess.Popen(
            [sys.executable, str(script_path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=1
        )
        
        current_substep = 0
        total_substeps = 0
        
        # Read output line by line
        for line in process.stdout:
            print(line, end='')
            
            # Check for progress markers in the output
            # Format: PROGRESS:current/total
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
                except:
                    pass
        
        # Wait for process to complete
        return_code = process.wait()
        
        if return_code == 0:
            print(f"\n✓ Successfully completed: {script_path.name}")
            return True
        else:
            stderr = process.stderr.read()
            print(f"\n✗ Error executing {script_path.name}:")
            print(f"Return code: {return_code}")
            if stderr:
                print(f"Error: {stderr}")
            return False
        
    except Exception as e:
        print(f"\n✗ Unexpected error executing {script_path.name}: {str(e)}")
        return False

def run_scripts(base_folder, progress_window):
    """
    Find and execute all scripts with progress updates
    """
    print(f"{'='*80}")
    print(f"RPA Automation - Batch Executor")
    print(f"{'='*80}\n")
    
    # Check if relatorios folder exists
    if not os.path.exists(base_folder):
        print(f"Error: '{base_folder}' folder not found!")
        progress_window.update_status("Error: Folder not found!")
        return
    
    # Find all Python scripts
    scripts = find_python_scripts(base_folder)
    
    if not scripts:
        print(f"No Python scripts found in '{base_folder}' subfolders.")
        progress_window.update_status("No scripts found")
        return
    
    print(f"Found {len(scripts)} script(s) to execute:\n")
    for i, script in enumerate(scripts, 1):
        print(f"  {i}. {script.relative_to(base_folder)}")
    
    print()
    
    # Execute each script
    results = {}
    for i, script in enumerate(scripts, 1):
        progress_window.update_progress(i-1, len(scripts), script.name)
        success = execute_script(script, progress_window, i-1, len(scripts))
        results[script.name] = success
    
    # Update to show completion
    progress_window.update_progress(len(scripts), len(scripts), "Complete")
    
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
    Main function to execute all scripts with GUI progress window
    """
    base_folder = "relatorios"
    
    # Cria parâmetros de data
    datesFilter = DeterminaDataECaminho(r"C:\temp", "TesteRPA", start_day=3)
    datesFilter.create_folder()

    # Create progress window
    progress_window = FloatingProgressWindow()
    
    # Run scripts in a separate thread to keep GUI responsive
    script_thread = Thread(target=run_scripts, args=(base_folder, progress_window))
    script_thread.daemon = True
    script_thread.start()
    
    # Start the GUI event loop
    progress_window.root.mainloop()

if __name__ == "__main__":
    main()