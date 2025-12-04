# Contributing to RPA SIG

Thank you for your interest in contributing to RPA SIG! This document provides guidelines and instructions for contributing to the project.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [How Can I Contribute?](#how-can-i-contribute)
- [Development Setup](#development-setup)
- [Coding Standards](#coding-standards)
- [Commit Guidelines](#commit-guidelines)
- [Pull Request Process](#pull-request-process)

## Code of Conduct

By participating in this project, you agree to maintain a respectful and inclusive environment for all contributors.

## How Can I Contribute?

### Reporting Bugs

Before creating a bug report:
- Check the [Issues](https://github.com/yourusername/rpa-sig/issues) to see if it's already been reported
- Collect information about the bug:
  - Python version
  - Windows version
  - Steps to reproduce
  - Expected vs actual behavior
  - Screenshots if applicable

Create a bug report with the **Bug Report** template including:
- A clear, descriptive title
- Detailed steps to reproduce
- What you expected to happen
- What actually happened
- Any relevant logs or screenshots

### Suggesting Enhancements

Enhancement suggestions are tracked as GitHub issues. When creating an enhancement suggestion:
- Use a clear, descriptive title
- Provide a detailed description of the suggested enhancement
- Explain why this enhancement would be useful
- Include mockups or examples if applicable

### Adding New Reports

To add a new report automation script:

1. Create a new Python file in the appropriate category folder:
   - `relatorios/SIG_Suprimentos/` for supply chain reports
   - `relatorios/SIG_Producao/` for production reports

2. Use the standard template:
```python
class ParametrosDados:
    def __init__(self, nomePR="PR12345", nomeArquivo="XX - Report Name"):
        self.nomePR = nomePR
        self.nomeArquivo = nomeArquivo

### Standard header ###
import pyautogui
import time
import sys
from pathlib import Path
import argparse

# Argument parser for dates and path
parser = argparse.ArgumentParser()
parser.add_argument('--initial_date', type=str, required=False)
parser.add_argument('--final_date', type=str, required=False)
parser.add_argument('--path', type=str, required=False)
args = parser.parse_args()

if args.initial_date and args.final_date and args.path:
    class DateFilter:
        def __init__(self, initial, final, path):
            self.initial_date = initial
            self.final_date = final
            self.path = path
    
    datesFilter = DateFilter(args.initial_date, args.final_date, args.path)
else:
    print("Error: required arguments missing.")
    sys.exit()

# Path configuration
current_file = Path(__file__).resolve()
project_root = current_file.parent.parent.parent
sys.path.insert(0, str(project_root))

# Module imports
from modules import (
    locate_image_on_screen,
    AbrePR,
    LimpaPR,
    ClickOnExcel,
    CarregandoDados,
    WaitOnWindow,
    MouseBusy
)

# Progress tracking
total_steps = 5  # Update based on your script
current_step = 0

# Your automation logic here
```

3. Capture necessary UI element images
4. Test thoroughly
5. Document any special requirements

### Improving Documentation

Documentation improvements are always welcome:
- Fix typos or unclear instructions
- Add missing documentation
- Improve examples
- Translate documentation

## Development Setup

1. Fork the repository
2. Clone your fork:
   ```bash
   git clone https://github.com/yourusername/rpa-sig.git
   cd rpa-sig
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Create a branch for your changes:
   ```bash
   git checkout -b feature/your-feature-name
   ```

## Coding Standards

### Python Style Guide

Follow [PEP 8](https://pep8.org/) with these specifications:

- **Indentation**: 4 spaces (no tabs)
- **Line length**: Maximum 100 characters for code, 80 for docstrings/comments
- **Naming conventions**:
  - Functions: `snake_case`
  - Classes: `PascalCase`
  - Constants: `UPPER_SNAKE_CASE`
  - Variables: `snake_case`

### Code Quality

- Write clear, self-documenting code
- Add comments for complex logic
- Include docstrings for functions and classes:
```python
def exemplo_funcao(param1, param2):
    """
    Brief description of what the function does.
    
    Args:
        param1 (type): Description of param1
        param2 (type): Description of param2
    
    Returns:
        type: Description of return value
    """
    pass
```

### Error Handling

Always include proper error handling:
```python
try:
    # Your code
    pass
except SpecificException as e:
    print(f"Error: {e}")
    # Handle error appropriately
```

### Image Recognition

When adding new image assets:
- Use descriptive names: `{PR_NAME}-{ELEMENT}-{DESCRIPTION}.png`
- Capture at 1920x1080 resolution
- Save as PNG format
- Place in appropriate folders:
  - `base/` for report-specific elements
  - `modules/{module}-IMG/` for module-specific elements

## Commit Guidelines

### Commit Message Format

Use the following format:
```
<type>(<scope>): <subject>

<body>

<footer>
```

**Types**:
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation changes
- `style`: Code style changes (formatting, no logic change)
- `refactor`: Code refactoring
- `test`: Adding or updating tests
- `chore`: Maintenance tasks

**Examples**:
```
feat(reports): add new production efficiency report

Implements PRX123456 report for tracking production efficiency
metrics. Includes filters for date range and department.

Closes #42
```

```
fix(modules): correct image recognition timeout

Increased timeout in CarregandoDados from 60s to 120s to prevent
premature script termination on slower systems.
```

## Pull Request Process

1. **Update documentation**: Ensure README and other docs reflect your changes

2. **Test your changes**: Run the application and test affected functionality

3. **Update the CHANGELOG**: Add your changes under "Unreleased" section

4. **Create Pull Request**: Use the PR template and provide:
   - Clear description of changes
   - Related issue numbers
   - Screenshots/videos if UI changes
   - Test results

5. **Code Review**: Address any feedback from reviewers

6. **Merge**: Once approved, your PR will be merged

### Pull Request Checklist

Before submitting:
- [ ] Code follows project style guidelines
- [ ] Self-review of code completed
- [ ] Comments added to complex code
- [ ] Documentation updated
- [ ] No new warnings generated
- [ ] Tested on Windows 10/11
- [ ] Related issues referenced

## Testing

### Manual Testing

1. Test the specific report/feature you modified
2. Test the main GUI workflow
3. Verify progress tracking works correctly
4. Check error handling with invalid inputs
5. Ensure output files are generated correctly

### Test Checklist

- [ ] Script completes without errors
- [ ] Progress indicators update correctly
- [ ] Excel files are saved in correct location
- [ ] Error messages are clear and helpful
- [ ] UI remains responsive during execution
- [ ] Console output is informative

## Getting Help

- **Questions**: Open a GitHub issue with the "question" label
- **Discussion**: Use GitHub Discussions for broader topics
- **Real-time chat**: [Add your preferred chat platform]

## Recognition

Contributors will be recognized in:
- GitHub contributors page
- CHANGELOG.md
- README.md acknowledgments section

## License

By contributing, you agree that your contributions will be licensed under the MIT License.

---

Thank you for contributing to RPA SIG! Your efforts help make automation accessible and reliable for everyone.
