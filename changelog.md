# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Initial project setup
- Installation scripts for automated dependency installation
- Comprehensive documentation (README, SETUP_GUIDE, CONTRIBUTING)

## [1.0.0] - 2024-12-04

### Added
- GUI-based script selection interface with tabbed organization
- Period selection dropdown with last 6 months and next 2 months
- Real-time progress tracking with visual progress bars
- Optional floating console window with retro styling
- Batch script execution capability
- Automatic Excel export functionality
- 21 automated report scripts across two categories:
  - 8 Production reports (SIG_Producao)
  - 13 Supply Chain reports (SIG_Suprimentos)

### Modules
- `AbrePR`: Opens PR reports in Prosyst ERP
- `CarregandoDados`: Intelligent loading screen detection
- `CheckBoxCheck`: Automated checkbox state management
- `ClickOnExcel`: Multi-format Excel export (PRX and standard reports)
- `ClipToExcel`: Clipboard to Excel conversion utility
- `DateFolder`: Date range and folder path management
- `Layout`: Report layout selection automation
- `LocateImageOnScreen`: Robust image recognition for UI elements
- `MouseBusy`: Windows cursor state detection
- `WaitOnWindow`: Window detection with timeout handling
- `WaitWhileImageExists`: Continuous image monitoring

### Features
- Thread-safe GUI updates with queue-based messaging
- Customizable date ranges based on selected period
- Automatic output folder creation with year/month structure
- Window management (centering, maximizing, positioning)
- Multiple timeout configurations for different operations
- Support for both PRX and standard PR report types
- Chrome WebDriver integration for web-based reports
- PDF generation for external data sources (IBGE, BCB)

### UI Components
- Main script selector window with category tabs
- "Select All" checkboxes per category
- Progress window with step-by-step tracking
- Semi-transparent console window with colored output
- Custom font support for retro terminal aesthetic

### Technical
- Python 3.8+ compatibility
- Windows API integration (pywin32)
- Image recognition with PyAutoGUI
- Web automation with Selenium
- Excel generation with xlsxwriter
- Date manipulation with python-dateutil

## [0.9.0] - 2024-11-XX (Beta)

### Added
- Initial beta release
- Core automation framework
- Basic GUI prototype
- First set of report scripts

### Changed
- Refactored module structure for better reusability
- Improved error handling across all modules

### Fixed
- Window detection timeout issues
- Image recognition false positives
- Excel export path handling

## Version History Legend

### Types of Changes
- `Added` for new features
- `Changed` for changes in existing functionality
- `Deprecated` for soon-to-be removed features
- `Removed` for now removed features
- `Fixed` for any bug fixes
- `Security` in case of vulnerabilities

---

## Release Notes

### Version 1.0.0 Highlights

This is the first stable release of RPA SIG, bringing together months of development and testing. The system now provides a complete, production-ready solution for automating report generation from Prosyst ERP.

**Key improvements from beta**:
- Redesigned GUI with better organization
- Enhanced error handling and recovery
- Comprehensive documentation
- Improved performance and reliability
- Thread-safe operations throughout

**Known Issues**:
- Requires specific screen resolution (1920x1080) for optimal image recognition
- Windows-only compatibility due to Windows API dependencies
- Chrome must be installed for web-based reports

**Upgrade Notes**:
- No migration needed for fresh installations
- Beta users should backup existing configuration and image assets before upgrading

---

[Unreleased]: https://github.com/yourusername/rpa-sig/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/yourusername/rpa-sig/releases/tag/v1.0.0
[0.9.0]: https://github.com/yourusername/rpa-sig/releases/tag/v0.9.0
