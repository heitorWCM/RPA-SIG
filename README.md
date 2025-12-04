# RPA SIG - Automated Report Generator

![Python](https://img.shields.io/badge/python-3.8+-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

A Robotic Process Automation (RPA) system designed to automate the extraction of reports from Prosyst ERP. This tool provides a user-friendly GUI to select and execute multiple automated report generation scripts.

## 📋 Features

- **GUI-Based Script Selection**: Easy-to-use interface for selecting which reports to generate
- **Batch Processing**: Run multiple reports sequentially with a single click
- **Period Selection**: Choose reporting periods from the last 6 months to 2 months ahead
- **Real-time Progress Tracking**: Visual progress bars and console output
- **Floating Console Window**: Optional transparent console for monitoring execution
- **Modular Architecture**: Organized into reusable modules and report scripts
- **Automatic Excel Export**: Reports are automatically saved as Excel files
- **Error Handling**: Robust error handling and timeout management

## 🖥️ System Requirements

- **Operating System**: Windows 10 or higher (required for Windows API features)
- **Python**: Version 3.8 or higher
- **Prosyst ERP**: Must be installed and configured
- **Screen Resolution**: 1920x1080 or higher recommended
- **RAM**: 4GB minimum, 8GB recommended

## 📦 Installation

### Option 1: Automated Installation (Recommended)

1. Clone the repository:
```bash
git clone https://github.com/yourusername/rpa-sig.git
cd rpa-sig
```

2. Run the installation script:
```bash
install.bat
```

### Option 2: Manual Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/rpa-sig.git
cd rpa-sig
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## 🚀 Usage

### Basic Usage

1. Launch the application:
```bash
python Main.py
```

2. Select the desired reporting period from the dropdown menu

3. Choose which reports to generate by checking the boxes in the tabs

4. (Optional) Check "Show Console Window" to see detailed execution logs

5. Click "RUN SCRIPTS" to start the automation

### Report Categories

The application organizes reports into two main categories:

#### SIG_Suprimentos (Supply Chain)
- Estoque de MP (Raw Material Inventory)
- Estoque de Produto Intermediário (Intermediate Product Inventory)
- Movimentação de MP (Raw Material Movement)
- Movimentação de Intermediário (Intermediate Product Movement)
- Itens Gerais (General Items)
- Aquisição de MP (Raw Material Acquisition)
- Aquisição de Embalagem (Packaging Acquisition)
- Uso e Consumo (Use and Consumption)
- Fretes (Freight Reports)
- Inflação IBGE (IBGE Inflation Data)
- Cotação Dollar (Dollar Exchange Rate)

#### SIG_Producao (Production)
- Produção Bruta (Gross Production)
- Reprocesso (Reprocessing)
- Número de OPs (Number of Production Orders)
- Adição (Addition)
- Sucata (Scrap)
- NC (Non-Conformity)
- LP (Loss Prevention)
- Fuligem (Soot)

## 📁 Project Structure

```
rpa-sig/
├── Main.py                          # Main application entry point
├── requirements.txt                 # Python dependencies
├── install.bat                      # Windows installation script
├── install.sh                       # Linux/Mac installation script
├── README.md                        # This file
│
├── assets/                          # Application assets
│   ├── icon.ico                     # Application icon
│   └── Nouveau_IBM_Stretch.ttf      # Custom font for console
│
├── modules/                         # Reusable automation modules
│   ├── __init__.py
│   ├── AbrePR.py                    # Opens PR reports in Prosyst
│   ├── CarregandoDados.py           # Waits for data loading screens
│   ├── CheckBoxCheck.py             # Checks/unchecks checkboxes
│   ├── ClickOnExcel.py              # Exports and saves Excel files
│   ├── ClipToExcel.py               # Clipboard to Excel conversion
│   ├── DateFolder.py                # Date and folder management
│   ├── Layout.py                    # Report layout selection
│   ├── LocateImageOnScreen.py       # Image recognition for UI elements
│   ├── MouseBusy.py                 # Detects busy cursor state
│   ├── WaitOnWindow.py              # Waits for specific windows
│   ├── WaitWhileImageExists.py      # Waits while image is visible
│   │
│   ├── AbrePR-IMG/                  # Images for opening reports
│   ├── CarregandoDados-IMG/         # Loading screen images
│   ├── CheckBoxCheck-IMG/           # Checkbox state images
│   ├── ClickOnExcel-IMG/            # Excel export UI images
│   └── Layout-IMG/                  # Layout selection images
│
├── relatorios/                      # Report scripts
│   ├── SIG_Producao/                # Production reports
│   │   ├── 01-PRODUCAO_BRUTA.py
│   │   ├── 02-REPROCESSO.py
│   │   ├── 03-NUMERO_OPs.py
│   │   ├── 04-ADICAO.py
│   │   ├── 05-SUCATA.py
│   │   ├── 06-NC.py
│   │   ├── 07-LP.py
│   │   └── 08-FULIGEM.py
│   │
│   └── SIG_Suprimentos/             # Supply chain reports
│       ├── 01-ESTOQUE_DE_MP.PY
│       ├── 02-ESTOQUE_DE_PRODUTO_INTERMEDIARIO.PY
│       ├── 03-MOVIMENTACAO_MP.py
│       ├── 04-MOVIMENTACAO_INTERMEDIARIO.py
│       ├── 05-ITENS_GERAIS.PY
│       ├── 06-AQUISICAO_MP.py
│       ├── 07-AQUISICAO_EMBALAGEM.py
│       ├── 08-USO_E_CONSUMO.PY
│       ├── 09-FRETES_CONTA_CONTABIL_14101.py
│       ├── 10-FRETES_CONTA_CONTABIL_42205.py
│       ├── 11-FRETES_CONTA_CONTABIL_42443.py
│       ├── 12-INFLAÇÃO_IBGE.py
│       └── 13-COTACAO_DOLLAR.py
│
└── base/                            # Base images for UI element recognition
    ├── PR07416-Filter.png
    ├── PRX012016-Filter.png
    ├── PRX004317-Filter.png
    ├── PRX034504-Filter.png
    └── ...
```

## 🔧 Configuration

### Adding New Reports

1. Create a new Python script in the appropriate category folder under `relatorios/`
2. Use the standard header template from existing scripts
3. Implement the `ParametrosDados` class with PR name and file name
4. Add necessary UI element images to the `base/` folder
5. The script will automatically appear in the GUI on next launch

### Customizing Date Ranges

The application calculates date ranges based on the selected period:
- **Initial Date**: First day of the previous month
- **Final Date**: Last day of the previous month
- **Output Path**: `C:\temp\TesteRPA\YYYY\MM.MON\`

Modify the date calculation logic in `Main.py` if different ranges are needed.

## 🐛 Troubleshooting

### Common Issues

**Issue**: Scripts fail to find UI elements
- **Solution**: Ensure Prosyst ERP is running at the correct screen resolution (1920x1080 recommended)
- **Solution**: Update image references in the `base/` folder if UI has changed

**Issue**: Window not found errors
- **Solution**: Check that Prosyst ERP is open and logged in before running scripts
- **Solution**: Increase timeout values in `WaitOnWindow` calls if system is slow

**Issue**: Excel export fails
- **Solution**: Ensure target directory exists and has write permissions
- **Solution**: Close any Excel files that might be locked

**Issue**: Mouse cursor detected as busy indefinitely
- **Solution**: Increase timeout values in `MouseBusy()` calls
- **Solution**: Check for system dialogs or popups blocking execution

### Debug Mode

Enable the console window to see detailed execution logs:
1. Check "Show Console Window" before running scripts
2. Monitor progress and error messages in real-time

## 📝 Adding Custom Fonts

The console window uses a custom retro font (`Nouveau_IBM_Stretch.ttf`). To use a different font:

1. Place your `.ttf` font file in the `assets/` folder
2. Update the font path in `FloatingConsoleWindow.__init__()` in `Main.py`
3. Modify the font family name to match your font

## 🤝 Contributing

Contributions are welcome! Please follow these guidelines:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## ⚠️ Disclaimer

This automation tool is designed for internal use with Prosyst ERP. Users are responsible for:
- Ensuring compliance with their organization's automation policies
- Maintaining proper security and access controls
- Validating the accuracy of generated reports
- Backing up data before running automation scripts

## 📧 Support

For issues, questions, or contributions, please open an issue on GitHub.

## 🙏 Acknowledgments

- Built with [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) for modern UI
- Uses [PyAutoGUI](https://github.com/asweigart/pyautogui) for automation
- Selenium for web-based data extraction

---

**Note**: This tool requires Windows OS and access to Prosyst ERP. It performs automated mouse and keyboard actions, so the computer should not be used for other tasks during execution.
