# Gmail to Outlook Autocomplete Builder

Restore Outlook's email address autocomplete functionality when using Gmail by scanning your sent messages and creating an importable contact list.

## 🎯 Problem Solved

When using Outlook with Gmail (IMAP), the autocomplete cache doesn't work properly:
- Modern Outlook (2010+) doesn't maintain autocomplete with Gmail IMAP
- The old .NK2 file format is obsolete
- Gmail doesn't sync autocomplete data with Outlook

**Solution:** This tool scans your Gmail sent messages, extracts all email addresses you've corresponded with, and creates a CSV file that Outlook can import as contacts - restoring autocomplete functionality.

## 🚀 Quick Start

### For Windows Users

1. **Get the tool**
   - Download the Windows executable from releases
   - Or build from source using `python build_exe.py`

2. **Run it**
   - Double-click `GmailAutocomplete.exe` for the GUI version
   - Or use `GmailAutocompleteCLI.exe` for command line

3. **Get Gmail App Password**
   - Go to https://myaccount.google.com/apppasswords
   - Generate an app-specific password
   - Use this instead of your regular Gmail password

4. **Scan & Export**
   - Enter your Gmail address and app password
   - Click "Start Processing"
   - Tool creates `outlook_contacts.csv`

5. **Import to Outlook**
   - Open Outlook
   - File → Open & Export → Import/Export
   - Choose "Import from another program or file"
   - Select "Comma Separated Values"
   - Browse to the CSV file
   - Import to Contacts folder

### For Mac Users

1. **Install**
   - Download `GmailAutocomplete-macOS.dmg`
   - Drag app to Applications
   - First run: Right-click → Open (unsigned app)

2. **Use GUI or CLI**
   ```bash
   # GUI: Open GmailAutocomplete.app from Applications
   
   # CLI: Install command line tool
   sudo cp gmail-autocomplete /usr/local/bin/
   gmail-autocomplete your.email@gmail.com
   ```

## 📋 Prerequisites

- **Gmail App Password** (required for security)
  - Enable 2-factor authentication
  - Generate app password at https://myaccount.google.com/apppasswords
- **IMAP enabled** in Gmail settings
- **Python 3.6+** (only if running from source)

## 💻 Running from Source

### Basic Usage

```bash
# Clone the repository
git clone [repository-url]
cd outlook

# No dependencies needed - uses Python standard library
python gmail_autocomplete_builder.py your.email@gmail.com

# Or use the GUI version
python gmail_autocomplete_gui.py
```

### Command Line Options

```bash
# Scan more messages (default: 500)
python gmail_autocomplete_builder.py your.email@gmail.com --max-messages 1000

# Specify output file
python gmail_autocomplete_builder.py your.email@gmail.com --output my_contacts.csv

# Provide password directly (not recommended)
python gmail_autocomplete_builder.py your.email@gmail.com --password APP_PASSWORD
```

## 🔨 Building Executables

### Windows (.exe)

```bash
# Install dependencies
pip install pyinstaller pywin32-ctypes pywin32

# Build executables
python build_exe.py
```

**Common Windows Issues:**
- If you get `pathlib` errors: `conda remove pathlib`
- If you get `pywintypes` errors: `pip install pywin32-ctypes pywin32`
- If you get Unicode encoding errors: Use latest build script from GitHub
- "Pillow not installed" warning: Safe to ignore (icons are optional)

Creates:
- `GmailAutocomplete.exe` - Windows GUI application
- `GmailAutocompleteCLI.exe` - Command line tool

### macOS (.app)

```bash
pip install pyinstaller
python build_macos.py
```

Creates:
- `GmailAutocomplete.app` - Mac application
- `gmail-autocomplete` - Command line binary
- `GmailAutocomplete-macOS.dmg` - Installer package

## 📁 Project Structure

```
outlook/
├── gmail_autocomplete_builder.py    # Core CLI script
├── gmail_autocomplete_gui.py        # GUI version (cross-platform)
├── gmail_autocomplete_mac.py        # macOS-optimized GUI
├── build_exe.py                     # Windows build script
├── build_macos.py                   # macOS build script
├── requirements.txt                 # Dependencies (just PyInstaller for building)
└── README.md                        # This file
```

## ✨ Features

- **Frequency Sorting**: Most-contacted addresses appear first in autocomplete
- **Name Extraction**: Automatically extracts names from email headers
- **Batch Processing**: Efficiently scans hundreds/thousands of messages
- **Privacy Focused**: Runs entirely locally, no data sent to external servers
- **Cross-Platform**: Works on Windows, macOS, and Linux
- **No Dependencies**: Core script uses only Python standard library

## 🔒 Security & Privacy

- **Local Processing**: All operations happen on your computer
- **No Data Transmission**: Nothing is sent to external servers
- **App Passwords**: Uses Google's secure app-specific passwords
- **Revocable Access**: App passwords can be revoked anytime
- **Read-Only**: Only reads your sent messages, makes no changes

## ⚙️ How It Works

1. **Connects** to Gmail via IMAP using your app password
2. **Scans** your Sent folder to find all recipient email addresses
3. **Counts** frequency of communication with each address
4. **Extracts** names when available from email headers
5. **Exports** to Outlook-compatible CSV format
6. **Sorts** by frequency so most-used addresses appear first

## 📊 Output Files

- **`outlook_contacts.csv`** - Main file for Outlook import
  - Contains: First Name, Last Name, Email Address, Display Name
  - Sorted by frequency of use
  
- **`outlook_contacts_report.txt`** - Frequency report
  - Shows top 50 most contacted addresses
  - Useful for reviewing before import

## 🐛 Troubleshooting

### "Authentication Failed"
- Use an App Password, not your regular Gmail password
- Enable IMAP in Gmail settings
- Check 2-factor authentication is enabled

### "Cannot find sent folder"
- The tool tries multiple folder names automatically
- Ensure you have sent messages in Gmail

### No Autocomplete After Import
- Restart Outlook after importing
- Verify contacts imported to main Contacts folder
- Check Outlook autocomplete settings are enabled

### Mac: "Cannot be opened - unidentified developer"
- Right-click the app and select "Open"
- Or: System Preferences → Security & Privacy → Allow

## 🤝 Contributing

Contributions are welcome! Please feel free to submit pull requests.

## 📝 License

This project is provided as-is for personal use. Feel free to modify and distribute.

## 🆘 Support

### Common Issues

**Q: Why do I need an App Password?**
A: Google requires app-specific passwords for third-party applications for security. Your regular password won't work.

**Q: How many messages should I scan?**
A: 500-1000 is usually sufficient. More messages = more complete autocomplete but longer processing time.

**Q: Will this work with Outlook.com/Hotmail?**
A: This tool is designed for Gmail accounts used with Outlook desktop application.

**Q: Can I run this regularly to update autocomplete?**
A: Yes, you can run it periodically and re-import to keep autocomplete current.

## 🔗 Links

- [Get Gmail App Password](https://myaccount.google.com/apppasswords)
- [Enable Gmail IMAP](https://support.google.com/mail/answer/7126229)
- [Outlook Import Help](https://support.microsoft.com/en-us/office/import-contacts-to-outlook-bb796340-b58a-46c1-90c7-b549b8f3c5f8)

---

**Note:** This tool was created to solve the specific problem of Outlook losing autocomplete functionality when used with Gmail. It's a workaround for a limitation in how Outlook handles IMAP accounts.