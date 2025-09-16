#!/usr/bin/env python3
"""
Build script to create macOS app bundle and command-line binary
Creates native Mac applications from Gmail Autocomplete Builder
"""

import os
import sys
import subprocess
import shutil
import plistlib

def check_pyinstaller():
    """Check if PyInstaller is installed"""
    try:
        import PyInstaller
        return True
    except ImportError:
        print("PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        return True

def create_icns():
    """Create macOS icon file (.icns) if Pillow is available"""
    icon_created = False
    try:
        from PIL import Image, ImageDraw
        
        # Create icon at multiple sizes for macOS
        sizes = [16, 32, 64, 128, 256, 512, 1024]
        
        for size in sizes:
            img = Image.new('RGBA', (size, size), color=(255, 255, 255, 0))
            draw = ImageDraw.Draw(img)
            
            # Scale envelope shape to icon size
            scale = size / 256
            
            # Draw envelope shape
            draw.rectangle(
                [int(40*scale), int(80*scale), int(216*scale), int(176*scale)],
                fill=(52, 152, 219),
                outline=(41, 128, 185),
                width=max(1, int(3*scale))
            )
            draw.polygon([
                (int(40*scale), int(80*scale)),
                (int(128*scale), int(140*scale)),
                (int(216*scale), int(80*scale))
            ], fill=(41, 128, 185))
            
            # Save each size
            img.save(f'icon_{size}.png')
        
        # Create ICNS using iconutil (macOS command)
        # First create iconset directory
        os.makedirs('GmailAutocomplete.iconset', exist_ok=True)
        
        # Move/rename files to iconset format
        icon_files = [
            (16, 'icon_16x16.png'),
            (32, 'icon_16x16@2x.png'),
            (32, 'icon_32x32.png'),
            (64, 'icon_32x32@2x.png'),
            (128, 'icon_128x128.png'),
            (256, 'icon_128x128@2x.png'),
            (256, 'icon_256x256.png'),
            (512, 'icon_256x256@2x.png'),
            (512, 'icon_512x512.png'),
            (1024, 'icon_512x512@2x.png'),
        ]
        
        for size, filename in icon_files:
            if os.path.exists(f'icon_{size}.png'):
                shutil.copy(f'icon_{size}.png', f'GmailAutocomplete.iconset/{filename}')
        
        # Convert to ICNS
        subprocess.run(['iconutil', '-c', 'icns', 'GmailAutocomplete.iconset'])
        
        # Clean up
        shutil.rmtree('GmailAutocomplete.iconset')
        for size in sizes:
            if os.path.exists(f'icon_{size}.png'):
                os.remove(f'icon_{size}.png')
        
        icon_created = os.path.exists('GmailAutocomplete.icns')
        if icon_created:
            print("✓ Created GmailAutocomplete.icns")
        
    except ImportError:
        print("! Pillow not installed, skipping icon creation")
        print("  To add an icon: pip install Pillow")
    except Exception as e:
        print(f"! Could not create icon: {e}")
    
    return icon_created

def build_mac_app():
    """Build the macOS app bundle"""
    print("\nBuilding macOS app bundle...")
    
    # Check for icon
    icon_param = []
    if os.path.exists('GmailAutocomplete.icns'):
        icon_param = ['--icon=GmailAutocomplete.icns']
    
    # PyInstaller command for Mac app
    cmd = [
        'pyinstaller',
        '--onefile',
        '--windowed',  # Creates .app bundle
        '--name=GmailAutocomplete',
        '--osx-bundle-identifier=com.gmail.autocomplete',
        '--clean',
        '--noconfirm',
    ] + icon_param + [
        '--add-data=GmailAutocomplete.icns:.',  # Include icon in bundle if it exists
        'gmail_autocomplete_gui.py'
    ] if os.path.exists('GmailAutocomplete.icns') else [
        'pyinstaller',
        '--onefile',
        '--windowed',
        '--name=GmailAutocomplete',
        '--osx-bundle-identifier=com.gmail.autocomplete',
        '--clean',
        '--noconfirm',
        'gmail_autocomplete_gui.py'
    ]
    
    try:
        subprocess.check_call(cmd)
        print("✓ Mac app bundle built successfully")
        
        # Customize Info.plist
        customize_app_bundle()
        
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ Failed to build Mac app: {e}")
        return False

def customize_app_bundle():
    """Customize the app bundle's Info.plist"""
    plist_path = 'dist/GmailAutocomplete.app/Contents/Info.plist'
    
    if not os.path.exists(plist_path):
        return
    
    try:
        # Read existing plist
        with open(plist_path, 'rb') as f:
            plist = plistlib.load(f)
        
        # Update values
        plist['CFBundleDisplayName'] = 'Gmail Autocomplete'
        plist['CFBundleName'] = 'Gmail Autocomplete'
        plist['NSHumanReadableCopyright'] = 'Gmail to Outlook Autocomplete Builder'
        plist['CFBundleShortVersionString'] = '1.0.0'
        plist['CFBundleVersion'] = '1.0.0'
        
        # Add high resolution capable flag
        plist['NSHighResolutionCapable'] = True
        
        # Write updated plist
        with open(plist_path, 'wb') as f:
            plistlib.dump(plist, f)
        
        print("✓ Customized app bundle Info.plist")
        
    except Exception as e:
        print(f"! Could not customize Info.plist: {e}")

def build_cli_binary():
    """Build the command-line binary"""
    print("\nBuilding command-line binary...")
    
    # PyInstaller command for CLI
    cmd = [
        'pyinstaller',
        '--onefile',
        '--console',
        '--name=gmail-autocomplete',
        '--clean',
        '--noconfirm',
        'gmail_autocomplete_builder.py'
    ]
    
    try:
        subprocess.check_call(cmd)
        print("✓ CLI binary built successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ Failed to build CLI binary: {e}")
        return False

def create_installer_dmg():
    """Create a DMG installer for easy distribution"""
    print("\nCreating DMG installer...")
    
    # Create temporary directory for DMG contents
    dmg_dir = 'GmailAutocomplete_macOS'
    if os.path.exists(dmg_dir):
        shutil.rmtree(dmg_dir)
    os.makedirs(dmg_dir)
    
    # Copy app bundle
    if os.path.exists('dist/GmailAutocomplete.app'):
        shutil.copytree('dist/GmailAutocomplete.app', 
                       f'{dmg_dir}/GmailAutocomplete.app')
    
    # Copy CLI binary
    if os.path.exists('dist/gmail-autocomplete'):
        shutil.copy('dist/gmail-autocomplete', dmg_dir)
    
    # Create Applications symlink
    os.symlink('/Applications', f'{dmg_dir}/Applications')
    
    # Create README
    readme = """Gmail to Outlook Autocomplete Builder for macOS
===============================================

INSTALLATION:
1. Drag GmailAutocomplete.app to Applications folder
2. For command-line tool, copy gmail-autocomplete to /usr/local/bin/

FIRST RUN:
• Right-click the app and select "Open" (required for unsigned apps)
• Enter your Gmail credentials and app password
• The tool will scan your sent messages and create a CSV file

COMMAND LINE USAGE:
./gmail-autocomplete your.email@gmail.com

REQUIREMENTS:
• macOS 10.15 or later
• Gmail account with IMAP enabled
• Gmail App Password (get from https://myaccount.google.com/apppasswords)

The app will create a CSV file that you can import into Outlook to restore
email address autocomplete functionality.
"""
    
    with open(f'{dmg_dir}/README.txt', 'w') as f:
        f.write(readme)
    
    # Create DMG
    dmg_name = 'GmailAutocomplete-macOS.dmg'
    
    try:
        # Remove old DMG if exists
        if os.path.exists(dmg_name):
            os.remove(dmg_name)
        
        # Create DMG using hdiutil
        subprocess.check_call([
            'hdiutil', 'create', '-volname', 'Gmail Autocomplete',
            '-srcfolder', dmg_dir, '-ov', '-format', 'UDZO',
            dmg_name
        ])
        
        print(f"✓ Created {dmg_name}")
        
        # Clean up temp directory
        shutil.rmtree(dmg_dir)
        
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"! Could not create DMG: {e}")
        return False

def create_homebrew_formula():
    """Create a Homebrew formula for easy installation"""
    formula = """class GmailAutocomplete < Formula
  desc "Gmail to Outlook Autocomplete Builder"
  homepage "https://github.com/yourusername/gmail-autocomplete"
  version "1.0.0"
  
  # For actual deployment, you'd host the binary and update this URL
  url "file://#{Dir.pwd}/dist/gmail-autocomplete"
  sha256 "PLACEHOLDER_SHA256"
  
  def install
    bin.install "gmail-autocomplete"
  end
  
  test do
    system "#{bin}/gmail-autocomplete", "--help"
  end
end
"""
    
    with open('gmail-autocomplete.rb', 'w') as f:
        f.write(formula)
    
    print("✓ Created Homebrew formula template (gmail-autocomplete.rb)")

def main():
    print("Gmail Autocomplete Builder - macOS Binary Creator")
    print("=" * 50)
    
    # Check Python version
    if sys.version_info < (3, 6):
        print("✗ Python 3.6 or higher required")
        sys.exit(1)
    
    # Check if running on macOS
    if sys.platform != 'darwin':
        print("✗ This script must be run on macOS to create Mac binaries")
        sys.exit(1)
    
    # Check/install PyInstaller
    if not check_pyinstaller():
        print("✗ Could not install PyInstaller")
        sys.exit(1)
    
    # Create icon
    create_icns()
    
    # Build binaries
    app_success = build_mac_app()
    cli_success = build_cli_binary()
    
    if not app_success and not cli_success:
        print("\n✗ Build failed!")
        sys.exit(1)
    
    # Create DMG installer
    if app_success:
        create_installer_dmg()
    
    # Create Homebrew formula template
    if cli_success:
        create_homebrew_formula()
    
    print("\n" + "=" * 50)
    print("BUILD COMPLETE!")
    print("=" * 50)
    print("\nCreated:")
    
    if app_success:
        print("  ✓ dist/GmailAutocomplete.app - Mac application")
    
    if cli_success:
        print("  ✓ dist/gmail-autocomplete - Command-line tool")
    
    if os.path.exists('GmailAutocomplete-macOS.dmg'):
        print("  ✓ GmailAutocomplete-macOS.dmg - Installer package")
    
    print("\nINSTALLATION:")
    print("  • GUI App: Copy GmailAutocomplete.app to /Applications")
    print("  • CLI Tool: Copy gmail-autocomplete to /usr/local/bin/")
    print("  • Or: Double-click the DMG and drag to Applications")
    
    print("\nUSAGE:")
    print("  • GUI: Double-click GmailAutocomplete.app")
    print("  • CLI: gmail-autocomplete your.email@gmail.com")
    
    print("\nNOTE: First run requires right-click → Open (unsigned app)")

if __name__ == '__main__':
    main()