# RPA SIG - Setup Guide

This guide will help you set up the RPA SIG project from scratch.

## Prerequisites Checklist

Before starting, ensure you have:

- [ ] Windows 10 or higher
- [ ] Python 3.8 or higher installed
- [ ] Prosyst ERP installed and configured
- [ ] Admin rights (for first-time installation)
- [ ] Internet connection (for downloading dependencies)

## Step-by-Step Setup

### 1. Install Python

If you don't have Python installed:

1. Download Python from [python.org](https://www.python.org/downloads/)
2. Run the installer
3. **Important**: Check "Add Python to PATH" during installation
4. Verify installation by opening Command Prompt and typing:
   ```bash
   python --version
   ```

### 2. Clone or Download the Project

**Option A: Using Git**
```bash
git clone https://github.com/yourusername/rpa-sig.git
cd rpa-sig
```

**Option B: Download ZIP**
1. Download the ZIP file from GitHub
2. Extract to your desired location
3. Open Command Prompt in the extracted folder

### 3. Install Dependencies

Run the installation script:
```bash
install.bat
```

Or manually install:
```bash
pip install -r requirements.txt
```

### 4. Verify Installation

Test that all libraries are installed:
```bash
python -c "import customtkinter, pyautogui, pygetwindow; print('All libraries installed successfully!')"
```

### 5. Configure the Application

#### Set Output Directory

By default, reports are saved to: `C:\temp\TesteRPA\`

To change this:
1. Open `Main.py`
2. Find the `run_scripts()` function
3. Modify the `path` variable (around line 526):
   ```python
   path = os.path.join(
       r"C:\your\desired\path",  # Change this
       "TesteRPA",
       reference_date.strftime("%Y"),
       f"{reference_date.strftime('%m')}.{reference_date.strftime('%b').upper()}"
   )
   ```

#### Add Custom Icon (Optional)

1. Place your `.ico` file in the `assets/` folder
2. Name it `icon.ico` or update the reference in `Main.py`

### 6. Prepare Image Assets

The automation relies on image recognition. Ensure you have:

1. **Base Images**: Screenshots of UI elements in the `base/` folder
2. **Module Images**: Screenshots for specific modules in their respective folders

#### Capturing Images

To capture images for recognition:

1. Open Prosyst ERP
2. Navigate to the screen you want to automate
3. Use Windows Snipping Tool (Win + Shift + S) to capture:
   - Buttons
   - Input fields
   - Checkboxes
   - Labels
4. Save as PNG with descriptive names
5. Place in appropriate folders

**Naming Convention**:
```
{PR_NAME}-{ELEMENT_TYPE}-{DESCRIPTION}.png

Examples:
PR07416-Filter.png
PRX012016-RightClickExport.png
PRX034504-DateFilter-1.png
```

### 7. Test the Application

1. Open Prosyst ERP and log in
2. Run the application:
   ```bash
   python Main.py
   ```
3. Try running a single simple report to verify everything works

### 8. Create a Shortcut (Optional)

**For Windows Desktop Shortcut:**

1. Right-click on `Main.py`
2. Select "Create shortcut"
3. Right-click the shortcut → Properties
4. In "Target" field, add Python path:
   ```
   C:\Path\To\Python\python.exe "C:\Path\To\rpa-sig\Main.py"
   ```
5. Change icon by clicking "Change Icon" and selecting `assets/icon.ico`

## Troubleshooting Setup Issues

### Python Not Found

**Error**: `'python' is not recognized as an internal or external command`

**Solution**:
1. Reinstall Python with "Add to PATH" checked
2. Or manually add Python to PATH:
   - Open System Properties → Environment Variables
   - Add Python installation directory to PATH

### Permission Denied

**Error**: `PermissionError: [Errno 13] Permission denied`

**Solution**:
1. Run Command Prompt as Administrator
2. Or change installation directory to user folder

### Module Not Found

**Error**: `ModuleNotFoundError: No module named 'customtkinter'`

**Solution**:
1. Ensure you're using the correct Python installation
2. Reinstall requirements:
   ```bash
   pip install -r requirements.txt --force-reinstall
   ```

### PyWin32 Installation Issues

**Error**: Issues installing `pywin32`

**Solution**:
1. Download the appropriate wheel file from [PyWin32 GitHub](https://github.com/mhammond/pywin32/releases)
2. Install manually:
   ```bash
   pip install pywin32-306-cp310-cp310-win_amd64.whl
   ```
   (Replace with your Python version)

### Screen Resolution Issues

**Issue**: Images not being recognized

**Solution**:
1. Set screen resolution to 1920x1080
2. Set DPI scaling to 100%
3. Recapture images at this resolution

## Performance Optimization

### Improve Execution Speed

1. **Close unnecessary applications** - Free up system resources
2. **Disable screen saver** - Prevents interruption during automation
3. **Set power plan to High Performance**
4. **Increase timeout values** if system is slow

### Reduce Memory Usage

1. Run fewer scripts simultaneously
2. Close console window if not needed
3. Limit number of loading screen images in `CarregandoDados-IMG/`

## Next Steps

Once setup is complete:

1. Read the main [README.md](README.md) for usage instructions
2. Review existing report scripts in `relatorios/` to understand the structure
3. Test with a single report before running batch operations
4. Create backups of your image assets
5. Document any custom configurations

## Getting Help

If you encounter issues during setup:

1. Check the [Troubleshooting](#troubleshooting-setup-issues) section above
2. Review Python and library documentation
3. Open an issue on GitHub with:
   - Your Python version
   - Windows version
   - Error messages
   - Steps to reproduce

## Maintenance

### Updating Dependencies

Periodically update libraries:
```bash
pip install -r requirements.txt --upgrade
```

### Backing Up Configurations

Before major changes:
1. Backup the entire project folder
2. Export your image assets
3. Document custom configurations

---

**Congratulations!** Your RPA SIG environment is now set up and ready to use. Proceed to the main README for usage instructions.
