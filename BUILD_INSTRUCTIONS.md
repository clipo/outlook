# Building Windows Executable from Gmail Autocomplete Builder

## Quick Build (Automated)

On a Windows machine with Python installed:

```bash
# Install PyInstaller
pip install pyinstaller

# Run the build script
python build_exe.py
```

This creates a `GmailAutocomplete_Windows` folder with:
- `GmailAutocomplete.exe` - GUI version (recommended)
- `GmailAutocompleteCLI.exe` - Command-line version
- `run_cli.bat` - Helper for CLI version

## Manual Build Instructions

### 1. Install PyInstaller and Windows Dependencies

```bash
# Install PyInstaller
pip install pyinstaller

# Install Windows-specific dependencies
pip install pywin32-ctypes pywin32

# Or install everything at once
pip install pyinstaller pywin32-ctypes pywin32
```

**Note:** Windows requires additional packages that may not be installed by default in Anaconda environments.

### 2. Build GUI Version (Recommended)

```bash
pyinstaller --onefile --windowed --name=GmailAutocomplete gmail_autocomplete_gui.py
```

- `--onefile`: Creates a single executable
- `--windowed`: No console window (GUI only)
- `--name`: Sets the executable name

### 3. Build CLI Version (Optional)

```bash
pyinstaller --onefile --console --name=GmailAutocompleteCLI gmail_autocomplete_builder.py
```

- `--console`: Keeps console window for CLI

### 4. Find Your Executables

After building, find the .exe files in the `dist/` folder.

## Adding a Custom Icon

To add an icon to your executable:

1. Create or download an `.ico` file
2. Add `--icon=youricon.ico` to the PyInstaller command:

```bash
pyinstaller --onefile --windowed --icon=mail.ico --name=GmailAutocomplete gmail_autocomplete_gui.py
```

## Cross-Platform Building

**Important:** You must build on the target platform:
- Build on Windows to create Windows .exe files
- Build on Mac to create Mac apps
- Build on Linux to create Linux executables

## Distributing the Executable

The resulting .exe file is completely standalone:
- No Python installation required on target machine
- All dependencies are included
- Just copy the .exe file to any Windows computer

## File Size

The executable will be ~10-15 MB because it includes:
- Python interpreter
- All required libraries
- Your script

## Troubleshooting

### "Could not import `pywintypes` or `win32api`"
```bash
pip install pywin32-ctypes pywin32
```

### "The 'pathlib' package is an obsolete backport"
```bash
conda remove pathlib
# or
pip uninstall pathlib
```

### "Failed to execute script"
- Make sure all import statements work when running the Python script normally
- Check that the script runs without errors: `python gmail_autocomplete_gui.py`

### Antivirus Warnings
- Some antivirus software may flag PyInstaller executables
- This is a false positive - you can whitelist the file
- Consider signing your executable with a code signing certificate for distribution

### Missing DLL errors
- Run on the same Windows version you built on
- Or use `--onefile` flag to include everything

### Build fails in Anaconda environment
Create a clean virtual environment:
```bash
python -m venv build_env
build_env\Scripts\activate
pip install pyinstaller pywin32-ctypes pywin32
python build_exe.py
```

### "UnicodeEncodeError: 'charmap' codec can't encode character"
This happens when Windows can't encode Unicode characters. The build script has been updated to handle this, but if you see this error:
- Make sure you're using the latest version of the build script from GitHub
- Or manually edit any text files to use ASCII characters instead of Unicode symbols

### Build succeeds but shows "Pillow not installed"
This is just a warning - the build will complete successfully. To add custom icons:
```bash
pip install Pillow
```
Then run the build script again.

## Alternative: py2exe

If PyInstaller doesn't work, try py2exe (Windows only):

```python
# setup.py for py2exe
from distutils.core import setup
import py2exe

setup(
    windows=['gmail_autocomplete_gui.py'],
    options={
        'py2exe': {
            'bundle_files': 1,
            'compressed': True,
        }
    },
    zipfile=None,
)
```

Then run:
```bash
python setup.py py2exe
```

## Testing the Executable

1. Copy the .exe to a different folder
2. Run it without Python installed
3. Test all features:
   - Connect to Gmail
   - Scan messages
   - Export CSV
   - Check the output file

## Creating an Installer (Optional)

For professional distribution, create an installer using:
- **NSIS** (Nullsoft Scriptable Install System)
- **Inno Setup**
- **WiX Toolset**

These create a setup.exe that:
- Installs the program
- Creates Start Menu shortcuts
- Adds uninstall option
- Can be signed with certificates