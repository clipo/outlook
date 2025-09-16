# Gmail Autocomplete Builder - macOS Installation Guide

## Quick Install

### Option 1: DMG Installer (Easiest)
1. Double-click `GmailAutocomplete-macOS.dmg`
2. Drag `GmailAutocomplete.app` to your Applications folder
3. First time: Right-click the app and select "Open" (required for unsigned apps)

### Option 2: Direct Copy
- **GUI App**: Copy `dist/GmailAutocomplete.app` to `/Applications`
- **CLI Tool**: Copy `dist/gmail-autocomplete` to `/usr/local/bin/`

## First Run Security

macOS requires verification for unsigned apps:
1. Right-click `GmailAutocomplete.app`
2. Select "Open" from the context menu
3. Click "Open" in the security dialog
4. This only needs to be done once

## Usage

### GUI Application
1. Open `GmailAutocomplete.app` from Applications
2. Enter your Gmail address
3. Enter your [App Password](https://myaccount.google.com/apppasswords)
4. Click "Scan Gmail & Create CSV"
5. The app creates a CSV file on your Desktop
6. Follow the import instructions shown

### Command Line Tool
```bash
# Basic usage (prompts for password)
gmail-autocomplete your.email@gmail.com

# With password
gmail-autocomplete your.email@gmail.com --password YOUR_APP_PASSWORD

# Scan more messages
gmail-autocomplete your.email@gmail.com --max-messages 1000

# Custom output location
gmail-autocomplete your.email@gmail.com --output ~/Documents/contacts.csv
```

## Installing CLI Tool System-Wide

To use `gmail-autocomplete` from anywhere:

```bash
# Copy to system bin
sudo cp dist/gmail-autocomplete /usr/local/bin/
sudo chmod +x /usr/local/bin/gmail-autocomplete

# Test it works
gmail-autocomplete --help
```

## Requirements

- macOS 10.15 (Catalina) or later
- Gmail account with IMAP enabled
- Gmail App Password (not your regular password)

## Getting a Gmail App Password

1. Go to https://myaccount.google.com/apppasswords
2. Sign in to your Google account
3. Enable 2-factor authentication if needed
4. Select "Mail" as the app type
5. Generate and copy the 16-character password
6. Use this password in the app

## Import to Outlook

After creating the CSV file:

1. Open Microsoft Outlook
2. Go to **File → Open & Export → Import/Export**
3. Choose **"Import from another program or file"**
4. Select **"Comma Separated Values"**
5. Browse to the CSV file (default: Desktop/outlook_contacts.csv)
6. Select your **Contacts** folder as destination
7. Map fields if prompted
8. Click **Finish**

## Troubleshooting

### "Cannot be opened because it is from an unidentified developer"
- Right-click the app and select "Open" instead of double-clicking
- Or: System Preferences → Security & Privacy → Allow app

### "Gmail login failed"
- Make sure you're using an App Password, not your regular password
- Verify IMAP is enabled in Gmail settings
- Check your internet connection

### CLI tool "command not found"
- Make sure you copied it to `/usr/local/bin/`
- Or run directly: `./dist/gmail-autocomplete`

### No autocomplete after import
- Restart Outlook after importing
- Check contacts were imported to the main Contacts folder
- Verify Outlook autocomplete is enabled in settings

## Features

- **Smart Sorting**: Most-contacted addresses appear first
- **Name Recognition**: Extracts names from email headers
- **Batch Processing**: Efficiently scans hundreds of messages
- **Privacy**: Runs entirely locally, no data sent externally
- **Native macOS**: Built specifically for Mac with native UI

## Uninstall

To remove the application:
1. Delete `GmailAutocomplete.app` from `/Applications`
2. Remove CLI tool: `sudo rm /usr/local/bin/gmail-autocomplete`
3. Delete any created CSV files

## Building from Source

If you want to build the binaries yourself:

```bash
# Install PyInstaller
pip install pyinstaller

# Run the build script
python build_macos.py

# Creates:
# - dist/GmailAutocomplete.app (GUI)
# - dist/gmail-autocomplete (CLI)
# - GmailAutocomplete-macOS.dmg (Installer)
```

## Support

- The app requires macOS 10.15 or later
- Works with Gmail, G Suite, and Google Workspace accounts
- Compatible with Outlook 2016, 2019, 2021, and Microsoft 365

## Privacy & Security

- All processing happens locally on your Mac
- No data is transmitted to external servers
- Your password is only used for Gmail IMAP connection
- App passwords can be revoked anytime from Google Account settings