#!/usr/bin/env python3
"""
Gmail Autocomplete Builder for Outlook
Scans Gmail sent messages and creates an importable contact list for Outlook autocomplete
"""

import imaplib
import email
from email.header import decode_header
import re
import csv
import json
from datetime import datetime
from collections import defaultdict
import getpass
import argparse
import ssl

class GmailAutocompleteBuilder:
    def __init__(self, email_address, password=None, app_password=None):
        self.email_address = email_address
        self.password = password or app_password
        self.imap = None
        self.email_addresses = defaultdict(lambda: {'count': 0, 'name': '', 'last_used': None})
        
    def connect(self):
        """Connect to Gmail via IMAP"""
        try:
            # Create SSL context
            context = ssl.create_default_context()
            
            # Connect to Gmail IMAP
            self.imap = imaplib.IMAP4_SSL('imap.gmail.com', 993, ssl_context=context)
            
            # Login
            self.imap.login(self.email_address, self.password)
            print(f"✓ Connected to Gmail account: {self.email_address}")
            return True
            
        except Exception as e:
            print(f"✗ Failed to connect: {e}")
            print("\nTroubleshooting tips:")
            print("1. Enable 'Less secure app access' or use an App Password")
            print("2. Enable IMAP in Gmail settings")
            print("3. For App Password: https://myaccount.google.com/apppasswords")
            return False
    
    def extract_email_addresses(self, email_string):
        """Extract email addresses from various email header formats"""
        addresses = []
        
        # Pattern for email addresses
        email_pattern = r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
        
        # Split by comma for multiple recipients
        parts = email_string.split(',')
        
        for part in parts:
            # Extract name and email
            name = ''
            email_addr = ''
            
            # Check for "Name <email>" format
            if '<' in part and '>' in part:
                name_part = part.split('<')[0].strip()
                email_part = part.split('<')[1].split('>')[0].strip()
                name = name_part.strip('"\'')
                email_addr = email_part
            else:
                # Just email address
                matches = re.findall(email_pattern, part)
                if matches:
                    email_addr = matches[0]
            
            if email_addr and email_addr.lower() != self.email_address.lower():
                addresses.append((email_addr.lower(), name))
        
        return addresses
    
    def scan_sent_folder(self, max_messages=500):
        """Scan sent messages for recipient email addresses"""
        print(f"\nScanning sent messages (up to {max_messages})...")
        
        try:
            # Select Sent folder (try different names)
            sent_folders = ['[Gmail]/Sent Mail', 'Sent', 'INBOX.Sent', '[Gmail]/Sent']
            
            folder_found = False
            for folder in sent_folders:
                try:
                    self.imap.select(f'"{folder}"', readonly=True)
                    folder_found = True
                    print(f"✓ Found sent folder: {folder}")
                    break
                except:
                    continue
            
            if not folder_found:
                print("✗ Could not find sent folder")
                return False
            
            # Search for all messages
            _, message_ids = self.imap.search(None, 'ALL')
            message_ids = message_ids[0].split()
            
            # Limit to most recent messages
            message_ids = message_ids[-max_messages:] if len(message_ids) > max_messages else message_ids
            
            print(f"Processing {len(message_ids)} messages...")
            
            for idx, msg_id in enumerate(message_ids, 1):
                if idx % 50 == 0:
                    print(f"  Processed {idx}/{len(message_ids)} messages...")
                
                try:
                    # Fetch message
                    _, msg_data = self.imap.fetch(msg_id, '(RFC822)')
                    raw_email = msg_data[0][1]
                    
                    # Parse email
                    msg = email.message_from_bytes(raw_email)
                    
                    # Get date
                    date_str = msg.get('Date', '')
                    
                    # Process To, Cc, and Bcc fields
                    for field in ['To', 'Cc', 'Bcc']:
                        recipients = msg.get(field, '')
                        if recipients:
                            # Decode header if needed
                            decoded = decode_header(recipients)[0]
                            if decoded[1]:
                                recipients = decoded[0].decode(decoded[1])
                            elif isinstance(decoded[0], bytes):
                                recipients = decoded[0].decode('utf-8', errors='ignore')
                            else:
                                recipients = decoded[0]
                            
                            # Extract addresses
                            addresses = self.extract_email_addresses(recipients)
                            
                            for email_addr, name in addresses:
                                self.email_addresses[email_addr]['count'] += 1
                                if name and not self.email_addresses[email_addr]['name']:
                                    self.email_addresses[email_addr]['name'] = name
                                if date_str and not self.email_addresses[email_addr]['last_used']:
                                    self.email_addresses[email_addr]['last_used'] = date_str
                
                except Exception as e:
                    continue
            
            print(f"✓ Found {len(self.email_addresses)} unique email addresses")
            return True
            
        except Exception as e:
            print(f"✗ Error scanning messages: {e}")
            return False
    
    def export_to_csv(self, filename='outlook_contacts.csv'):
        """Export to CSV format that Outlook can import"""
        print(f"\nExporting to CSV: {filename}")
        
        # Sort by frequency of use
        sorted_addresses = sorted(self.email_addresses.items(), 
                                 key=lambda x: x[1]['count'], 
                                 reverse=True)
        
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            # Outlook-compatible CSV headers
            fieldnames = ['First Name', 'Last Name', 'E-mail Address', 'E-mail Display As']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for email_addr, info in sorted_addresses:
                display_name = info['name'] if info['name'] else email_addr
                
                # Split name into first/last
                name_parts = info['name'].split() if info['name'] else []
                first_name = name_parts[0] if name_parts else ''
                last_name = ' '.join(name_parts[1:]) if len(name_parts) > 1 else ''
                
                writer.writerow({
                    'First Name': first_name,
                    'Last Name': last_name,
                    'E-mail Address': email_addr,
                    'E-mail Display As': f"{display_name} ({email_addr})" if info['name'] else email_addr
                })
        
        print(f"✓ Exported {len(sorted_addresses)} contacts to {filename}")
        
        # Also create a frequency report
        report_file = filename.replace('.csv', '_report.txt')
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("Email Address Frequency Report\n")
            f.write("=" * 50 + "\n\n")
            
            for email_addr, info in sorted_addresses[:50]:  # Top 50
                f.write(f"{email_addr:<40} - {info['count']} messages\n")
                if info['name']:
                    f.write(f"  Name: {info['name']}\n")
        
        print(f"✓ Created frequency report: {report_file}")
        
        return filename
    
    def disconnect(self):
        """Disconnect from Gmail"""
        if self.imap:
            try:
                self.imap.close()
                self.imap.logout()
            except:
                pass

def main():
    parser = argparse.ArgumentParser(description='Build Outlook autocomplete from Gmail sent messages')
    parser.add_argument('email', help='Your Gmail email address')
    parser.add_argument('--password', help='Your Gmail password or app password (will prompt if not provided)')
    parser.add_argument('--max-messages', type=int, default=500, help='Maximum messages to scan (default: 500)')
    parser.add_argument('--output', default='outlook_contacts.csv', help='Output CSV filename')
    
    args = parser.parse_args()
    
    # Get password if not provided
    password = args.password
    if not password:
        print("\nYou'll need to use an App Password for Gmail:")
        print("1. Go to https://myaccount.google.com/apppasswords")
        print("2. Generate an app-specific password")
        print("3. Use that password here\n")
        password = getpass.getpass(f"Enter app password for {args.email}: ")
    
    # Create builder
    builder = GmailAutocompleteBuilder(args.email, password)
    
    # Process
    if builder.connect():
        if builder.scan_sent_folder(max_messages=args.max_messages):
            csv_file = builder.export_to_csv(args.output)
            
            print("\n" + "=" * 50)
            print("SUCCESS! Next steps to import into Outlook:")
            print("=" * 50)
            print("\n1. Open Outlook")
            print("2. Go to File → Open & Export → Import/Export")
            print("3. Choose 'Import from another program or file'")
            print("4. Select 'Comma Separated Values'")
            print(f"5. Browse to: {csv_file}")
            print("6. Select your Contacts folder as destination")
            print("7. Map the fields if needed")
            print("8. Click Finish")
            print("\nThe imported contacts will appear in autocomplete!")
            
        builder.disconnect()
    
    print("\nDone!")

if __name__ == '__main__':
    main()