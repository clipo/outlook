#!/usr/bin/env python3
"""
Gmail Autocomplete Builder for Outlook - GUI Version
Standalone Windows application for rebuilding Outlook autocomplete from Gmail
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
from datetime import datetime

class GmailAutocompleteGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Gmail to Outlook Autocomplete Builder")
        self.root.geometry("600x550")
        
        # Set icon if bundled with PyInstaller
        try:
            if hasattr(sys, '_MEIPASS'):
                icon_path = os.path.join(sys._MEIPASS, 'icon.ico')
                if os.path.exists(icon_path):
                    self.root.iconbitmap(icon_path)
        except:
            pass
        
        self.email_addresses = defaultdict(lambda: {'count': 0, 'name': '', 'last_used': None})
        self.imap = None
        self.processing = False
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title = ttk.Label(main_frame, text="Gmail to Outlook Autocomplete Builder", 
                         font=('Arial', 14, 'bold'))
        title.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Instructions
        instructions = ttk.Label(main_frame, text="This tool scans your Gmail sent messages\nand creates a contact list for Outlook autocomplete", 
                                font=('Arial', 10))
        instructions.grid(row=1, column=0, columnspan=2, pady=5)
        
        # Email input
        ttk.Label(main_frame, text="Gmail Address:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.email_var = tk.StringVar()
        email_entry = ttk.Entry(main_frame, textvariable=self.email_var, width=40)
        email_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Password input
        ttk.Label(main_frame, text="App Password:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar()
        password_entry = ttk.Entry(main_frame, textvariable=self.password_var, show="*", width=40)
        password_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # App password help link
        help_text = ttk.Label(main_frame, text="Get App Password", 
                             foreground="blue", cursor="hand2", font=('Arial', 9, 'underline'))
        help_text.grid(row=4, column=1, sticky=tk.W)
        help_text.bind("<Button-1>", lambda e: self.show_help())
        
        # Number of messages
        ttk.Label(main_frame, text="Messages to scan:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.messages_var = tk.StringVar(value="500")
        messages_spinbox = ttk.Spinbox(main_frame, from_=100, to=5000, increment=100, 
                                       textvariable=self.messages_var, width=10)
        messages_spinbox.grid(row=5, column=1, sticky=tk.W, pady=5)
        
        # Output file
        ttk.Label(main_frame, text="Output file:").grid(row=6, column=0, sticky=tk.W, pady=5)
        self.output_var = tk.StringVar(value="outlook_contacts.csv")
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=6, column=1, sticky=(tk.W, tk.E), pady=5)
        output_entry = ttk.Entry(output_frame, textvariable=self.output_var, width=30)
        output_entry.pack(side=tk.LEFT)
        browse_btn = ttk.Button(output_frame, text="Browse...", command=self.browse_output)
        browse_btn.pack(side=tk.LEFT, padx=5)
        
        # Process button
        self.process_btn = ttk.Button(main_frame, text="Start Processing", 
                                     command=self.start_processing, width=20)
        self.process_btn.grid(row=7, column=0, columnspan=2, pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, length=400, mode='indeterminate')
        self.progress.grid(row=8, column=0, columnspan=2, pady=5)
        
        # Status text
        self.status_text = scrolledtext.ScrolledText(main_frame, height=10, width=70, 
                                                     state=tk.DISABLED, wrap=tk.WORD)
        self.status_text.grid(row=9, column=0, columnspan=2, pady=10)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
    def show_help(self):
        help_window = tk.Toplevel(self.root)
        help_window.title("How to get Gmail App Password")
        help_window.geometry("500x300")
        
        help_text = """How to get a Gmail App Password:

1. Go to your Google Account settings:
   https://myaccount.google.com/apppasswords

2. You may need to enable 2-factor authentication first

3. Click "Select app" and choose "Mail"

4. Click "Select device" and choose "Other"

5. Enter "Outlook Autocomplete" as the name

6. Click "Generate"

7. Copy the 16-character password shown

8. Use this password in the app (not your regular Gmail password)

Note: Also make sure IMAP is enabled in Gmail:
Settings → Forwarding and POP/IMAP → Enable IMAP"""
        
        text_widget = tk.Text(help_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert(1.0, help_text)
        text_widget.config(state=tk.DISABLED)
        
    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=self.output_var.get()
        )
        if filename:
            self.output_var.set(filename)
    
    def log_message(self, message, level="info"):
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
        
        # Start processing in thread
        self.processing = True
        self.process_btn.config(state=tk.DISABLED, text="Processing...")
        self.progress.start(10)
        
        thread = threading.Thread(target=self.process_emails)
        thread.daemon = True
        thread.start()
    
    def process_emails(self):
        try:
            # Connect
            self.log_message("Connecting to Gmail...")
            if not self.connect_gmail():
                return
            
            # Scan messages
            self.log_message("Scanning sent messages...")
            if not self.scan_sent_folder():
                return
            
            # Export
            self.log_message("Exporting to CSV...")
            output_file = self.export_to_csv()
            
            # Success
            self.log_message(f"Successfully exported {len(self.email_addresses)} contacts!", "success")
            self.log_message(f"Output file: {output_file}", "success")
            
            # Show import instructions
            self.show_import_instructions(output_file)
            
        except Exception as e:
            self.log_message(f"Error: {str(e)}", "error")
            messagebox.showerror("Error", f"Processing failed: {str(e)}")
        
        finally:
            self.disconnect_gmail()
            self.processing = False
            self.progress.stop()
            self.process_btn.config(state=tk.NORMAL, text="Start Processing")
    
    def connect_gmail(self):
        try:
            context = ssl.create_default_context()
            self.imap = imaplib.IMAP4_SSL('imap.gmail.com', 993, ssl_context=context)
            self.imap.login(self.email_var.get(), self.password_var.get())
            self.log_message(f"Connected to {self.email_var.get()}", "success")
            return True
        except Exception as e:
            self.log_message(f"Connection failed: {str(e)}", "error")
            self.log_message("Check your email, password, and IMAP settings", "error")
            return False
    
    def scan_sent_folder(self):
        try:
            max_messages = int(self.messages_var.get())
            
            # Find sent folder
            sent_folders = ['[Gmail]/Sent Mail', 'Sent', 'INBOX.Sent', '[Gmail]/Sent']
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
            
            # Search messages
            _, message_ids = self.imap.search(None, 'ALL')
            message_ids = message_ids[0].split()
            
            # Limit messages
            message_ids = message_ids[-max_messages:] if len(message_ids) > max_messages else message_ids
            total = len(message_ids)
            
            self.log_message(f"Processing {total} messages...")
            
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
            
            self.log_message(f"Found {len(self.email_addresses)} unique email addresses", "success")
            return True
            
        except Exception as e:
            self.log_message(f"Error scanning messages: {str(e)}", "error")
            return False
    
    def extract_email_addresses(self, email_string):
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
    
    def export_to_csv(self):
        filename = self.output_var.get()
        
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
        
        # Create report
        report_file = filename.replace('.csv', '_report.txt')
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("Top 50 Most Contacted Email Addresses\n")
            f.write("=" * 50 + "\n\n")
            
            for email_addr, info in sorted_addresses[:50]:
                f.write(f"{email_addr:<40} - {info['count']} messages\n")
                if info['name']:
                    f.write(f"  Name: {info['name']}\n")
        
        return filename
    
    def disconnect_gmail(self):
        if self.imap:
            try:
                self.imap.close()
                self.imap.logout()
            except:
                pass
    
    def show_import_instructions(self, filename):
        instructions = f"""SUCCESS! Your contacts have been exported.

To import into Outlook:

1. Open Outlook
2. Go to File → Open & Export → Import/Export
3. Choose "Import from another program or file"
4. Select "Comma Separated Values"
5. Browse to: {os.path.abspath(filename)}
6. Select your Contacts folder as destination
7. Map the fields if prompted
8. Click Finish

The imported contacts will appear in autocomplete!"""
        
        messagebox.showinfo("Import Instructions", instructions)

def main():
    root = tk.Tk()
    app = GmailAutocompleteGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()