# Specification: Create .venv and 1-Click Setup and Launcher

## Overview

This task creates automated setup and launch scripts for the docs-to-markdown application. The goal is to provide a frictionless onboarding experience where users can set up the development environment and run the application with a single click, rather than manually executing multiple PowerShell commands as currently documented in the README.

## Workflow Type

**Type**: feature

**Rationale**: This is a feature addition that enhances the user experience by automating setup and launch processes. It does not modify existing functionality but adds new scripts to make the application more accessible to non-technical users.

## Task Scope

### Services Involved
- **main** (primary) - Python application that converts documents to Markdown

### This Task Will:
- [ ] Create a Windows batch file (`setup.bat`) for one-click environment setup
- [ ] Create a Windows batch file (`run.bat`) for one-click application launch
- [ ] Create a PowerShell version (`setup.ps1`) for users who prefer PowerShell
- [ ] Create a PowerShell version (`run.ps1`) for users who prefer PowerShell
- [ ] Ensure scripts handle virtual environment creation and activation
- [ ] Ensure scripts install dependencies from requirements.txt
- [ ] Ensure scripts install the package in editable mode
- [ ] Add error handling and user-friendly messages
- [ ] Update README.md with instructions for using the new scripts

### Out of Scope:
- Modifying the core application logic
- Creating GUI launchers
- Supporting Linux/macOS (could be added in future iterations)
- Automated testing infrastructure (that's a separate task)

## Service Context

### Main Service

**Tech Stack:**
- Language: Python 3.9+
- Framework: None (CLI application)
- Key directories: 
  - `docs_to_markdown/` - Source code
  - `tests/` - Test suite

**Entry Point:** `docs_to_markdown/__main__.py` (exposed as `docs-to-markdown` console script)

**How to Run (Manual):**
```bash
python -m venv .venv
.\\.venv\\Scripts\\Activate.ps1
python -m pip install -r requirements.txt
python -m pip install -e .
docs-to-markdown "input_path" "output_path"
```

**Dependencies:**
- mammoth==1.10.0
- markdownify==0.14.1
- beautifulsoup4==4.12.3
- lxml==5.3.0

**Testing:** pytest

## Files to Modify

| File | Service | What to Change |
|------|---------|---------------|
| `README.md` | main | Add section about 1-click setup and launcher scripts |

## Files to Create

| File | Purpose | Key Features |
|------|---------|--------------|
| `setup.bat` | Windows batch setup script | Create venv, install dependencies, install package |
| `run.bat` | Windows batch launcher | Activate venv, launch app with args |
| `setup.ps1` | PowerShell setup script | Same as setup.bat but with PowerShell features |
| `run.ps1` | PowerShell launcher | Same as run.bat but with PowerShell features |

## Files to Reference

These files show patterns to follow:

| File | Pattern to Copy |
|------|----------------|
| `requirements.txt` | Dependencies to install |
| `pyproject.toml` | Package installation command (`pip install -e .`) |
| `README.md` | Current manual setup process to automate |

## Patterns to Follow

### Virtual Environment Setup Pattern

From `README.md` manual instructions:

```powershell
python -m venv .venv
.\\.venv\\Scripts\\Activate.ps1
python -m pip install -r requirements.txt
python -m pip install -e .
```

**Key Points:**
- Use `python -m venv` for cross-platform compatibility
- Activate using Scripts directory on Windows
- Install from requirements.txt for dependencies
- Install in editable mode with `-e .` flag for development

### Package Launch Pattern

From `pyproject.toml` console script configuration:

```toml
[project.scripts]
docs-to-markdown = "docs_to_markdown.__main__:main"
```

**Key Points:**
- The package exposes a `docs-to-markdown` console script
- After installation, can be run directly from command line
- Accepts input and output path arguments

## Requirements

### Functional Requirements

1. **One-Click Setup Script**
   - Description: A script that creates the virtual environment and installs all dependencies with a single double-click
   - Acceptance: Running setup.bat/setup.ps1 creates .venv, installs all packages, and displays success message
   - Must handle: Existing .venv (skip or recreate), missing Python (error with helpful message)

2. **One-Click Launcher Script**
   - Description: A script that launches the application with the virtual environment activated
   - Acceptance: Running run.bat/run.ps1 activates venv and launches docs-to-markdown CLI
   - Must handle: Missing .venv (error message telling user to run setup first), missing arguments (show usage)

3. **Error Handling**
   - Description: Scripts should provide clear error messages for common failure scenarios
   - Acceptance: Scripts catch errors and display user-friendly messages before exiting
   - Must handle: Python not installed, network errors during pip install, permission errors

4. **Documentation Update**
   - Description: Update README.md to document the new 1-click setup process
   - Acceptance: README includes new section with clear instructions for using setup.bat/run.bat
   - Must include: Quick start section, alternative manual setup reference, troubleshooting tips

### Edge Cases

1. **Existing .venv** - Setup script should detect existing virtual environment and ask user whether to recreate or skip
2. **Python not in PATH** - Script should detect this and provide clear instructions on how to fix
3. **Network connectivity issues** - pip install may fail; script should retry and provide helpful error message
4. **Permissions issues** - Script should detect and inform user if they need admin privileges
5. **Different Python versions** - Script should verify Python 3.9+ is installed
6. **Arguments with spaces** - Launcher should properly handle paths with spaces when passing to CLI

## Implementation Notes

### DO
- Follow the manual setup process from README.md for setup script
- Use standard Python venv module (not virtualenv)
- Include pause at end of batch files so users can see output
- Add echo statements to show progress to user
- Check for Python availability before proceeding
- Use `%~dp0` in batch files for script directory reference
- Use `$PSScriptRoot` in PowerShell for script directory reference
- Update README with both batch and PowerShell instructions
- Include --help flag usage example in launcher

### DON'T
- Hardcode absolute paths (use relative paths from script location)
- Assume Python is always called "python" (check for python3 too)
- Skip error checking for critical operations
- Create overly complex scripts (keep them simple and maintainable)
- Use non-standard virtual environment locations
- Modify the core application code
- Create scripts that require user to be administrator

## Development Environment

### Current Project State

```bash
# The project already has:
- .venv directory exists
- Dependencies installed globally
- Package structure with console script entry point
- Manual setup documented in README
```

### Required Environment Variables
- None (scripts should work without environment variables)

### Python Version
- Python 3.9 or higher required

## Success Criteria

The task is complete when:

1. [ ] setup.bat successfully creates .venv and installs all dependencies
2. [ ] run.bat successfully launches the application with proper arguments
3. [ ] setup.ps1 provides equivalent functionality to setup.bat
4. [ ] run.ps1 provides equivalent functionality to run.bat
5. [ ] All scripts include helpful error messages and user feedback
6. [ ] README.md is updated with 1-click setup instructions
7. [ ] Scripts handle common edge cases (existing .venv, missing Python, etc.)
8. [ ] No console errors during normal operation
9. [ ] Scripts work on fresh Windows installation (after Python is installed)

## QA Acceptance Criteria

**CRITICAL**: These criteria must be verified by the QA Agent before sign-off.

### Unit Tests
| Test | File | What to Verify |
|------|------|----------------|
| Setup script creates .venv | setup.bat/setup.ps1 | Virtual environment directory exists after running |
| Setup installs dependencies | setup.bat/setup.ps1 | All packages from requirements.txt are installed in .venv |
| Setup installs package | setup.bat/setup.ps1 | docs-to-markdown command available in .venv |
| Launcher activates venv | run.bat/run.ps1 | Application runs using venv Python |

### Integration Tests
| Test | Services | What to Verify |
|------|----------|----------------|
| Full workflow | setup → run | User can run setup then immediately run app successfully |
| Fresh install | New environment | Scripts work on system without existing .venv |
| Existing venv | setup with .venv present | Script handles existing virtual environment appropriately |

### End-to-End Tests
| Flow | Steps | Expected Outcome |
|------|-------|------------------|
| New user onboarding | 1. Clone repo 2. Double-click setup.bat 3. Wait for completion 4. Double-click run.bat | Application launches successfully and shows usage/help |
| Convert documents | 1. Run setup 2. Create test docx 3. run.bat input output | Documents converted successfully to markdown |

### Browser Verification (if frontend)
| Page/Component | URL | Checks |
|----------------|-----|--------|
| N/A | N/A | This is a CLI application, no browser verification needed |

### Database Verification (if applicable)
| Check | Query/Command | Expected |
|-------|---------------|----------|
| N/A | N/A | No database used |

### Script Verification
| Script | Command | Expected Behavior |
|--------|---------|-------------------|
| setup.bat | Double-click or `setup.bat` | Creates .venv, installs deps, shows success message |
| run.bat | `run.bat "input" "output"` | Activates venv, runs conversion, shows results |
| setup.ps1 | `.\setup.ps1` | Same as setup.bat with PowerShell features |
| run.ps1 | `.\run.ps1 "input" "output"` | Same as run.bat with PowerShell features |

### QA Sign-off Requirements
- [ ] All unit tests pass
- [ ] All integration tests pass
- [ ] All E2E tests pass
- [ ] Scripts work on Windows 10/11
- [ ] Scripts work with Python 3.9, 3.10, 3.11, 3.12
- [ ] Error messages are clear and helpful
- [ ] README documentation is accurate and complete
- [ ] Scripts handle edge cases gracefully
- [ ] No regressions in existing functionality
- [ ] Code follows established patterns (simple, maintainable)
- [ ] No security vulnerabilities introduced (scripts don't expose credentials)
