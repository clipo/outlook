#!/usr/bin/env python3
"""
Gmail Autocomplete Builder for Outlook - macOS Native Version
Optimized for macOS with native look and feel
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import threading
import imaplib
import email
from email.header import decode_header
import re
import csv
from collections import defaultdict
import ssl
import sys
import os
import subprocess
from datetime import datetime
import platform

class GmailAutocompleteMac:
    def __init__(self, root):
        self.root = root
        self.root.title("Gmail to Outlook Autocomplete Builder")
        
        # Set window size and center it
        window_width = 650
        window_height = 600
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # macOS-specific styling
        self.setup_mac_style()
        
        self.email_addresses = defaultdict(lambda: {'count': 0, 'name': '', 'last_used': None})
        self.imap = None
        self.processing = False
        
        self.setup_ui()
        self.setup_mac_menu()
        
    def setup_mac_style(self):
        """Configure macOS-specific styling"""
        # Make window look native on macOS
        if platform.system() == 'Darwin':
            try:
                # Set macOS-specific window attributes
                self.root.tk.call('tk', 'windowingsystem')  # Returns 'aqua' on macOS
                
                # Configure native macOS button style
                style = ttk.Style()
                style.theme_use('aqua')
                
            except:
                pass
    
    def setup_mac_menu(self):
        """Create native macOS menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Application menu (shows as "Gmail Autocomplete" on macOS)
        if platform.system() == 'Darwin':
            app_menu = tk.Menu(menubar, tearoff=0)
            menubar.add_cascade(label='Gmail Autocomplete', menu=app_menu)
            app_menu.add_command(label='About Gmail Autocomplete', command=self.show_about)
            app_menu.add_separator()
            app_menu.add_command(label='Preferences...', command=self.show_preferences, accelerator='⌘,')
            app_menu.add_separator()
            app_menu.add_command(label='Quit', command=self.root.quit, accelerator='⌘Q')
            
            # Bind Command+Q for quit
            self.root.bind('<Command-q>', lambda e: self.root.quit())
            self.root.bind('<Command-comma>', lambda e: self.show_preferences())
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='File', menu=file_menu)
        file_menu.add_command(label='Export CSV...', command=self.export_csv, accelerator='⌘E')
        file_menu.add_command(label='Import to Outlook...', command=self.show_import_instructions)
        
        if platform.system() != 'Darwin':
            file_menu.add_separator()
            file_menu.add_command(label='Exit', command=self.root.quit)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='Help', menu=help_menu)
        help_menu.add_command(label='Gmail App Password Help', command=self.show_help)
        help_menu.add_command(label='Import Instructions', command=self.show_import_help)
        
        # Bind keyboard shortcuts
        self.root.bind('<Command-e>', lambda e: self.export_csv())
        
    def setup_ui(self):
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title with larger font
        title = ttk.Label(main_frame, text="Gmail to Outlook Autocomplete Builder", 
                         font=('-apple-system', 16, 'bold'))
        title.grid(row=0, column=0, columnspan=2, pady=(0, 5))
        
        # Subtitle
        subtitle = ttk.Label(main_frame, text="Restore Outlook's email autocomplete from Gmail history", 
                            font=('-apple-system', 11), foreground='gray')
        subtitle.grid(row=1, column=0, columnspan=2, pady=(0, 20))
        
        # Input section with frame
        input_frame = ttk.LabelFrame(main_frame, text="Gmail Account", padding="10")
        input_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        input_frame.columnconfigure(1, weight=1)
        
        # Email input
        ttk.Label(input_frame, text="Email Address:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.email_var = tk.StringVar()
        email_entry = ttk.Entry(input_frame, textvariable=self.email_var, width=35)
        email_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
        # Password input
        ttk.Label(input_frame, text="App Password:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar()
        password_entry = ttk.Entry(input_frame, textvariable=self.password_var, show="•", width=35)
        password_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
        # Help button for app password
        help_btn = ttk.Button(input_frame, text="Get App Password", command=self.show_help)
        help_btn.grid(row=2, column=1, sticky=tk.W, pady=5, padx=(10, 0))
        
        # Settings section
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        settings_frame.columnconfigure(1, weight=1)
        
        # Number of messages
        ttk.Label(settings_frame, text="Messages to scan:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.messages_var = tk.StringVar(value="500")
        messages_spinbox = ttk.Spinbox(settings_frame, from_=100, to=5000, increment=100, 
                                       textvariable=self.messages_var, width=15)
        messages_spinbox.grid(row=0, column=1, sticky=tk.W, pady=5, padx=(10, 0))
        
        # Output file
        ttk.Label(settings_frame, text="Save as:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_var = tk.StringVar(value="~/Desktop/outlook_contacts.csv")
        output_frame = ttk.Frame(settings_frame)
        output_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        output_frame.columnconfigure(0, weight=1)
        
        output_entry = ttk.Entry(output_frame, textvariable=self.output_var)
        output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        browse_btn = ttk.Button(output_frame, text="Browse", command=self.browse_output)
        browse_btn.grid(row=0, column=1, padx=(5, 0))
        
        # Process button
        self.process_btn = ttk.Button(main_frame, text="Scan Gmail & Create CSV", 
                                     command=self.start_processing)
        self.process_btn.grid(row=4, column=0, columnspan=2, pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, length=400, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=2, pady=(0, 10))
        
        # Status text with frame
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="5")
        status_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
        
        self.status_text = scrolledtext.ScrolledText(status_frame, height=8, width=60, 
                                                     state=tk.DISABLED, wrap=tk.WORD,
                                                     font=('Monaco', 10))
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights for resizing
        main_frame.rowconfigure(6, weight=1)
        
    def show_about(self):
        """Show about dialog"""
        about_text = """Gmail to Outlook Autocomplete Builder
Version 1.0.0

Scans your Gmail sent messages to rebuild
Outlook's email address autocomplete cache.

Created for macOS"""
        
        messagebox.showinfo("About", about_text)
    
    def show_preferences(self):
        """Show preferences dialog"""
        pref_window = tk.Toplevel(self.root)
        pref_window.title("Preferences")
        pref_window.geometry("400x200")
        
        # Center the window
        pref_window.transient(self.root)
        pref_window.grab_set()
        
        frame = ttk.Frame(pref_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Default Settings", font=('-apple-system', 14, 'bold')).pack(anchor=tk.W)
        
        ttk.Label(frame, text="Messages to scan:").pack(anchor=tk.W, pady=(10, 0))
        messages_var = tk.StringVar(value=self.messages_var.get())
        ttk.Spinbox(frame, from_=100, to=5000, increment=100, 
                   textvariable=messages_var, width=20).pack(anchor=tk.W)
        
        ttk.Label(frame, text="Default output location:").pack(anchor=tk.W, pady=(10, 0))
        output_var = tk.StringVar(value="Desktop")
        ttk.Combobox(frame, textvariable=output_var, 
                    values=["Desktop", "Documents", "Downloads"], 
                    state="readonly", width=20).pack(anchor=tk.W)
        
        ttk.Button(frame, text="Save", 
                  command=lambda: self.save_preferences(messages_var.get(), pref_window)).pack(pady=20)
    
    def save_preferences(self, messages, window):
        """Save preferences"""
        self.messages_var.set(messages)
        window.destroy()
        messagebox.showinfo("Preferences", "Preferences saved")
    
    def show_help(self):
        """Show Gmail App Password help"""
        # Try to open in browser first
        try:
            subprocess.run(['open', 'https://myaccount.google.com/apppasswords'])
        except:
            pass
        
        # Also show dialog
        help_window = tk.Toplevel(self.root)
        help_window.title("Gmail App Password Help")
        help_window.geometry("550x400")
        help_window.transient(self.root)
        
        text = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, padx=15, pady=15)
        text.pack(fill=tk.BOTH, expand=True)
        
        help_content = """How to Get a Gmail App Password

1. Open your browser and go to:
   https://myaccount.google.com/apppasswords
   
2. Sign in to your Google account

3. You may need to enable 2-factor authentication first
   (Google requires this for app passwords)

4. Click "Select app" and choose "Mail"

5. Click "Select device" and choose "Other (custom name)"

6. Enter "Outlook Autocomplete" as the name

7. Click "Generate"

8. Copy the 16-character password shown
   (It will look like: xxxx xxxx xxxx xxxx)

9. Use this password in the app (NOT your regular Gmail password)

Important Notes:
• App passwords are required for security when using third-party apps
• Each app password can only be viewed once
• You can revoke app passwords anytime from your Google account
• Make sure IMAP is enabled in Gmail settings

Troubleshooting:
• Can't find app passwords? Enable 2-factor authentication first
• IMAP must be enabled: Gmail Settings → Forwarding and POP/IMAP"""
        
        text.insert(1.0, help_content)
        text.config(state=tk.DISABLED)
        
        # Add button to open in browser
        btn_frame = ttk.Frame(help_window)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Open in Browser", 
                  command=lambda: subprocess.run(['open', 'https://myaccount.google.com/apppasswords'])).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Close", command=help_window.destroy).pack(side=tk.LEFT, padx=5)
    
    def show_import_help(self):
        """Show Outlook import instructions"""
        self.show_import_instructions()
    
    def browse_output(self):
        """Browse for output file location"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile="outlook_contacts.csv",
            initialdir=os.path.expanduser("~/Desktop")
        )
        if filename:
            self.output_var.set(filename)
    
    def export_csv(self):
        """Export current data to CSV"""
        if not self.email_addresses:
            messagebox.showwarning("No Data", "No email addresses to export. Run a scan first.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile="outlook_contacts.csv"
        )
        
        if filename:
            self.export_to_csv(filename)
            messagebox.showinfo("Export Complete", f"Exported to:\n{filename}")
    
    def log_message(self, message, level="info"):
        """Log message to status text"""
        self.status_text.config(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if level == "error":
            self.status_text.insert(tk.END, f"[{timestamp}] ✗ {message}\n", "error")
            self.status_text.tag_config("error", foreground="red")
        elif level == "success":
            self.status_text.insert(tk.END, f"[{timestamp}] ✓ {message}\n", "success")
            self.status_text.tag_config("success", foreground="green")
        else:
            self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
        self.root.update_idletasks()
    
    def start_processing(self):
        """Start processing emails"""
        if self.processing:
            return
        
        # Validate inputs
        if not self.email_var.get():
            messagebox.showerror("Error", "Please enter your Gmail address")
            return
        
        if not self.password_var.get():
            messagebox.showerror("Error", "Please enter your app password")
            return
        
        # Clear status
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.config(state=tk.DISABLED)
        
        # Start processing
        self.processing = True
        self.process_btn.config(state=tk.DISABLED, text="Processing...")
        self.progress.start(10)
        
        # Run in thread
        thread = threading.Thread(target=self.process_emails)
        thread.daemon = True
        thread.start()
    
    def process_emails(self):
        """Process Gmail messages"""
        try:
            # Connect
            self.log_message("Connecting to Gmail...")
            if not self.connect_gmail():
                return
            
            # Scan
            self.log_message("Scanning sent messages...")
            if not self.scan_sent_folder():
                return
            
            # Export
            output_file = os.path.expanduser(self.output_var.get())
            self.log_message(f"Exporting to {output_file}...")
            self.export_to_csv(output_file)
            
            # Success
            self.log_message(f"Successfully exported {len(self.email_addresses)} contacts!", "success")
            self.log_message(f"File saved to: {output_file}", "success")
            
            # Show import instructions
            self.root.after(100, lambda: self.show_import_instructions(output_file))
            
        except Exception as e:
            self.log_message(f"Error: {str(e)}", "error")
            self.root.after(100, lambda: messagebox.showerror("Error", f"Processing failed: {str(e)}"))
        
        finally:
            self.disconnect_gmail()
            self.processing = False
            self.progress.stop()
            self.process_btn.config(state=tk.NORMAL, text="Scan Gmail & Create CSV")
    
    def connect_gmail(self):
        """Connect to Gmail via IMAP"""
        try:
            context = ssl.create_default_context()
            self.imap = imaplib.IMAP4_SSL('imap.gmail.com', 993, ssl_context=context)
            self.imap.login(self.email_var.get(), self.password_var.get())
            self.log_message(f"Connected to {self.email_var.get()}", "success")
            return True
        except Exception as e:
            self.log_message(f"Connection failed: {str(e)}", "error")
            return False
    
    def scan_sent_folder(self):
        """Scan Gmail sent folder"""
        try:
            max_messages = int(self.messages_var.get())
            
            # Find sent folder
            sent_folders = ['[Gmail]/Sent Mail', 'Sent', 'INBOX.Sent']
            folder_found = False
            
            for folder in sent_folders:
                try:
                    self.imap.select(f'"{folder}"', readonly=True)
                    folder_found = True
                    self.log_message(f"Found sent folder: {folder}")
                    break
                except:
                    continue
            
            if not folder_found:
                self.log_message("Could not find sent folder", "error")
                return False
            
            # Get messages
            _, message_ids = self.imap.search(None, 'ALL')
            message_ids = message_ids[0].split()
            
            # Limit messages
            message_ids = message_ids[-max_messages:] if len(message_ids) > max_messages else message_ids
            total = len(message_ids)
            
            self.log_message(f"Processing {total} messages...")
            
            # Process messages
            for idx, msg_id in enumerate(message_ids, 1):
                if idx % 50 == 0:
                    self.log_message(f"Processed {idx}/{total} messages...")
                
                try:
                    _, msg_data = self.imap.fetch(msg_id, '(RFC822)')
                    raw_email = msg_data[0][1]
                    msg = email.message_from_bytes(raw_email)
                    
                    # Process recipients
                    for field in ['To', 'Cc', 'Bcc']:
                        recipients = msg.get(field, '')
                        if recipients:
                            addresses = self.extract_email_addresses(recipients)
                            for email_addr, name in addresses:
                                self.email_addresses[email_addr]['count'] += 1
                                if name and not self.email_addresses[email_addr]['name']:
                                    self.email_addresses[email_addr]['name'] = name
                except:
                    continue
            
            self.log_message(f"Found {len(self.email_addresses)} unique addresses", "success")
            return True
            
        except Exception as e:
            self.log_message(f"Error scanning: {str(e)}", "error")
            return False
    
    def extract_email_addresses(self, email_string):
        """Extract email addresses from header string"""
        addresses = []
        email_pattern = r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
        
        # Decode if needed
        if isinstance(email_string, str):
            decoded_string = email_string
        else:
            decoded = decode_header(email_string)[0]
            if decoded[1]:
                decoded_string = decoded[0].decode(decoded[1])
            elif isinstance(decoded[0], bytes):
                decoded_string = decoded[0].decode('utf-8', errors='ignore')
            else:
                decoded_string = decoded[0]
        
        parts = decoded_string.split(',')
        
        for part in parts:
            name = ''
            email_addr = ''
            
            if '<' in part and '>' in part:
                name_part = part.split('<')[0].strip()
                email_part = part.split('<')[1].split('>')[0].strip()
                name = name_part.strip('"\'')
                email_addr = email_part
            else:
                matches = re.findall(email_pattern, part)
                if matches:
                    email_addr = matches[0]
            
            if email_addr and email_addr.lower() != self.email_var.get().lower():
                addresses.append((email_addr.lower(), name))
        
        return addresses
    
    def export_to_csv(self, filename=None):
        """Export to CSV file"""
        if not filename:
            filename = os.path.expanduser(self.output_var.get())
        
        sorted_addresses = sorted(self.email_addresses.items(),
                                 key=lambda x: x[1]['count'],
                                 reverse=True)
        
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['First Name', 'Last Name', 'E-mail Address', 'E-mail Display As']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for email_addr, info in sorted_addresses:
                display_name = info['name'] if info['name'] else email_addr
                name_parts = info['name'].split() if info['name'] else []
                first_name = name_parts[0] if name_parts else ''
                last_name = ' '.join(name_parts[1:]) if len(name_parts) > 1 else ''
                
                writer.writerow({
                    'First Name': first_name,
                    'Last Name': last_name,
                    'E-mail Address': email_addr,
                    'E-mail Display As': f"{display_name} ({email_addr})" if info['name'] else email_addr
                })
        
        return filename
    
    def disconnect_gmail(self):
        """Disconnect from Gmail"""
        if self.imap:
            try:
                self.imap.close()
                self.imap.logout()
            except:
                pass
    
    def show_import_instructions(self, filename=None):
        """Show Outlook import instructions"""
        if not filename:
            filename = "the CSV file"
        
        instructions = f"""Import Instructions for Outlook

File created: {filename}

To import into Outlook:

1. Open Microsoft Outlook
2. Go to File → Open & Export → Import/Export
3. Choose "Import from another program or file"
4. Select "Comma Separated Values"
5. Browse to: {filename}
6. Select your Contacts folder as destination
7. Map the fields if prompted
8. Click Finish

The imported contacts will appear in autocomplete!

Note: You may need to restart Outlook for changes to take effect."""
        
        # Create instruction window
        inst_window = tk.Toplevel(self.root)
        inst_window.title("Import Instructions")
        inst_window.geometry("600x400")
        inst_window.transient(self.root)
        
        text = scrolledtext.ScrolledText(inst_window, wrap=tk.WORD, padx=15, pady=15)
        text.pack(fill=tk.BOTH, expand=True)
        text.insert(1.0, instructions)
        text.config(state=tk.DISABLED)
        
        # Button frame
        btn_frame = ttk.Frame(inst_window)
        btn_frame.pack(pady=10)
        
        if filename and filename != "the CSV file":
            ttk.Button(btn_frame, text="Open File Location", 
                      command=lambda: subprocess.run(['open', '-R', filename])).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Copy Instructions", 
                  command=lambda: self.copy_to_clipboard(instructions)).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Close", command=inst_window.destroy).pack(side=tk.LEFT, padx=5)
    
    def copy_to_clipboard(self, text):
        """Copy text to clipboard"""
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("Copied", "Instructions copied to clipboard")

def main():
    root = tk.Tk()
    app = GmailAutocompleteMac(root)
    root.mainloop()

if __name__ == '__main__':
    main()