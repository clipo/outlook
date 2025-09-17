#!/usr/bin/env python3
"""
Build script to create Windows executable from Gmail Autocomplete Builder
Creates both GUI and command-line versions
"""

import os
import sys
import subprocess
import shutil

def check_pyinstaller():
    """Check if PyInstaller and Windows dependencies are installed"""
    missing_packages = []
    
    try:
        import PyInstaller
    except ImportError:
        missing_packages.append("pyinstaller")
    
    try:
        import win32api
    except ImportError:
        missing_packages.append("pywin32")
    
    try:
        import win32ctypes
    except ImportError:
        missing_packages.append("pywin32-ctypes")
    
    if missing_packages:
        print(f"Missing packages: {', '.join(missing_packages)}")
        print("Installing required packages...")
        for package in missing_packages:
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            except subprocess.CalledProcessError:
                print(f"Failed to install {package}. Please install manually:")
                print(f"pip install {package}")
                return False
    
    return True

def create_icon():
    """Create a simple icon file if Pillow is available"""
    icon_created = False
    try:
        from PIL import Image, ImageDraw, ImageFont
        
        # Create a simple icon
        img = Image.new('RGBA', (256, 256), color=(255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        
        # Draw envelope shape
        draw.rectangle([40, 80, 216, 176], fill=(52, 152, 219), outline=(41, 128, 185), width=3)
        draw.polygon([(40, 80), (128, 140), (216, 80)], fill=(41, 128, 185))
        
        # Save as ICO
        img.save('icon.ico', format='ICO', sizes=[(256, 256)])
        icon_created = True
        print("✓ Created icon.ico")
        
    except ImportError:
        print("! Pillow not installed, skipping icon creation")
        print("  To add an icon, install Pillow: pip install Pillow")
    except Exception as e:
        print(f"! Could not create icon: {e}")
    
    return icon_created

def build_gui_exe():
    """Build the GUI version executable"""
    print("\nBuilding GUI executable...")
    
    # Check if icon exists
    icon_param = []
    if os.path.exists('icon.ico'):
        icon_param = ['--icon=icon.ico', '--add-data=icon.ico;.']
    
    # PyInstaller command for GUI
    cmd = [
        'pyinstaller',
        '--onefile',
        '--windowed',  # No console window
        '--name=GmailAutocomplete',
        '--clean',
        '--noconfirm',
    ] + icon_param + [
        'gmail_autocomplete_gui.py'
    ]
    
    try:
        subprocess.check_call(cmd)
        print("✓ GUI executable built successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ Failed to build GUI executable: {e}")
        return False

def build_cli_exe():
    """Build the command-line version executable"""
    print("\nBuilding command-line executable...")
    
    # Check if icon exists
    icon_param = []
    if os.path.exists('icon.ico'):
        icon_param = ['--icon=icon.ico']
    
    # PyInstaller command for CLI
    cmd = [
        'pyinstaller',
        '--onefile',
        '--console',  # Keep console window
        '--name=GmailAutocompleteCLI',
        '--clean',
        '--noconfirm',
    ] + icon_param + [
        'gmail_autocomplete_builder.py'
    ]
    
    try:
        subprocess.check_call(cmd)
        print("✓ CLI executable built successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ Failed to build CLI executable: {e}")
        return False

def create_batch_file():
    """Create a batch file for easy CLI usage"""
    batch_content = """@echo off
echo Gmail to Outlook Autocomplete Builder (CLI)
echo ==========================================
echo.
echo This will scan your Gmail sent messages and create
echo a CSV file for importing into Outlook.
echo.
echo You'll need your Gmail address and an App Password.
echo Get an App Password at: https://myaccount.google.com/apppasswords
echo.
set /p EMAIL="Enter your Gmail address: "
set /p PASSWORD="Enter your App Password: "
echo.
GmailAutocompleteCLI.exe %EMAIL% --password %PASSWORD%
echo.
pause
"""
    
    with open('dist/run_cli.bat', 'w') as f:
        f.write(batch_content)
    
    print("✓ Created run_cli.bat helper script")

def organize_dist():
    """Organize the distribution folder"""
    print("\nOrganizing distribution folder...")
    
    # Create output directory
    output_dir = 'GmailAutocomplete_Windows'
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)
    
    # Copy executables
    files_to_copy = []
    
    if os.path.exists('dist/GmailAutocomplete.exe'):
        files_to_copy.append(('dist/GmailAutocomplete.exe', 'GmailAutocomplete.exe'))
    
    if os.path.exists('dist/GmailAutocompleteCLI.exe'):
        files_to_copy.append(('dist/GmailAutocompleteCLI.exe', 'GmailAutocompleteCLI.exe'))
    
    if os.path.exists('dist/run_cli.bat'):
        files_to_copy.append(('dist/run_cli.bat', 'run_cli.bat'))
    
    # Copy README
    if os.path.exists('README.md'):
        files_to_copy.append(('README.md', 'README.txt'))
    
    for src, dst in files_to_copy:
        shutil.copy2(src, os.path.join(output_dir, dst))
        print(f"  Copied {dst}")
    
    # Create quick start guide
    quickstart = """Gmail to Outlook Autocomplete Builder
======================================

QUICK START:
1. Double-click 'GmailAutocomplete.exe' for the GUI version
   OR
2. Double-click 'run_cli.bat' for the command-line version

REQUIREMENTS:
- Gmail account with IMAP enabled
- Gmail App Password (get from https://myaccount.google.com/apppasswords)

The tool will:
1. Connect to your Gmail account
2. Scan your sent messages
3. Create a CSV file with all email addresses
4. Show instructions for importing into Outlook

After running, import the CSV file into Outlook:
- File -> Open & Export -> Import/Export
- Choose "Import from another program or file"
- Select "Comma Separated Values"
- Browse to the CSV file created
- Import to Contacts folder

Questions? See README.txt for detailed instructions.
"""
    
    with open(os.path.join(output_dir, 'QUICK_START.txt'), 'w', encoding='utf-8') as f:
        f.write(quickstart)
    
    print(f"✓ Distribution package created in '{output_dir}' folder")
    
    return output_dir

def main():
    print("Gmail Autocomplete Builder - Windows Executable Creator")
    print("=" * 50)
    
    # Check Python version
    if sys.version_info < (3, 6):
        print("✗ Python 3.6 or higher required")
        sys.exit(1)
    
    # Check/install PyInstaller
    if not check_pyinstaller():
        print("✗ Could not install PyInstaller")
        sys.exit(1)
    
    # Create icon (optional)
    create_icon()
    
    # Build executables
    gui_success = build_gui_exe()
    cli_success = build_cli_exe()
    
    if not gui_success and not cli_success:
        print("\n✗ Build failed!")
        sys.exit(1)
    
    # Create batch file helper
    if cli_success:
        create_batch_file()
    
    # Organize output
    output_folder = organize_dist()
    
    print("\n" + "=" * 50)
    print("BUILD COMPLETE!")
    print("=" * 50)
    print(f"\nYour Windows executables are in: {output_folder}/")
    print("\nFiles created:")
    print("  - GmailAutocomplete.exe (GUI version - recommended)")
    if cli_success:
        print("  - GmailAutocompleteCLI.exe (Command-line version)")
        print("  - run_cli.bat (Helper script for CLI)")
    print("  - QUICK_START.txt (Instructions)")
    print("  - README.txt (Detailed documentation)")
    print("\nYou can now copy this folder to any Windows machine!")
    print("No Python installation required on the target machine.")
    
    # Cleanup build files (optional)
    answer = input("\nClean up build files? (y/n): ").lower()
    if answer == 'y':
        for folder in ['build', '__pycache__']:
            if os.path.exists(folder):
                shutil.rmtree(folder)
                print(f"  Removed {folder}/")
        for file in ['*.spec']:
            import glob
            for f in glob.glob(file):
                os.remove(f)
                print(f"  Removed {f}")

if __name__ == '__main__':
    main()