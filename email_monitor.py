#!/usr/bin/env python3
"""
Email Excel Monitor with Telegram Notifications
Monitors email inbox for Excel attachments, scans for specified name, and sends Telegram alerts.
Only processes emails from the last X minutes to avoid processing large backlogs.
"""

import imaplib
import email
from email.header import decode_header
import os
import sys
from datetime import datetime, timedelta
import requests
import openpyxl
import io


def send_telegram_message(bot_token, chat_id, message):
    """Send a message via Telegram bot."""
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {
        'chat_id': chat_id,
        'text': message,
        'parse_mode': 'HTML'
    }
    
    try:
        response = requests.post(url, json=payload, timeout=10)
        response.raise_for_status()
        print(f"✓ Telegram notification sent successfully")
        return True
    except requests.exceptions.RequestException as e:
        print(f"✗ Failed to send Telegram message: {e}")
        return False


def scan_excel_for_name(file_data, search_name, filename):
    """Scan all sheets in an Excel file for the search name."""
    matches = []
    
    try:
        # Load workbook from bytes
        workbook = openpyxl.load_workbook(io.BytesIO(file_data), data_only=True)
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            
            # Scan all cells in the sheet
            for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
                for col_idx, cell_value in enumerate(row, start=1):
                    if cell_value and search_name.lower() in str(cell_value).lower():
                        col_letter = openpyxl.utils.get_column_letter(col_idx)
                        matches.append({
                            'sheet': sheet_name,
                            'cell': f"{col_letter}{row_idx}",
                            'value': str(cell_value)
                        })
        
        workbook.close()
        
    except Exception as e:
        print(f"✗ Error scanning Excel file {filename}: {e}")
        return []
    
    return matches


def decode_email_subject(subject):
    """Decode email subject handling various encodings."""
    if subject is None:
        return "No Subject"
    
    decoded_parts = []
    for part, encoding in decode_header(subject):
        if isinstance(part, bytes):
            try:
                decoded_parts.append(part.decode(encoding or 'utf-8', errors='ignore'))
            except:
                decoded_parts.append(part.decode('utf-8', errors='ignore'))
        else:
            decoded_parts.append(str(part))
    
    return ''.join(decoded_parts)


def connect_to_email(email_address, app_password, imap_server='imap.gmail.com'):
    """Connect to email inbox via IMAP."""
    try:
        print(f"Connecting to {imap_server}...")
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_address, app_password)
        print(f"✓ Successfully logged in as {email_address}")
        return mail
    except imaplib.IMAP4.error as e:
        print(f"✗ IMAP login failed: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"✗ Connection error: {e}")
        sys.exit(1)


def process_emails(mail, search_name, bot_token, chat_id, check_last_minutes=10):
    """Process only emails received in the last X minutes."""
    
    # Select inbox
    mail.select('INBOX')
    
    # Calculate the date to search from (e.g., last 10 minutes)
    # Add a small buffer to account for clock differences
    since_time = datetime.now() - timedelta(minutes=check_last_minutes)
    since_date = since_time.strftime("%d-%b-%Y")
    
    print(f"Searching for emails since {since_time.strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Search for emails received after the specified date (read or unread)
    status, messages = mail.search(None, f'SINCE {since_date}')
    
    if status != 'OK':
        print("✗ Failed to search emails")
        return
    
    email_ids = messages[0].split()
    
    if not email_ids:
        print(f"No emails found in the last {check_last_minutes} minutes at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        return
    
    print(f"Found {len(email_ids)} email(s) in the last {check_last_minutes} minutes")
    
    total_matches = 0
    processed_count = 0
    
    for email_id in email_ids:
        try:
            # Fetch email
            status, msg_data = mail.fetch(email_id, '(RFC822)')
            
            if status != 'OK':
                continue
            
            # Parse email
            raw_email = msg_data[0][1]
            email_message = email.message_from_bytes(raw_email)
            
            # Get email details
            subject = decode_email_subject(email_message.get('Subject'))
            from_email = email_message.get('From')
            date = email_message.get('Date')
            
            # Parse the date to check if it's actually within our time window
            try:
                from email.utils import parsedate_to_datetime
                email_datetime = parsedate_to_datetime(date)
                
                # Skip if email is older than our time window
                if email_datetime < since_time:
                    continue
            except:
                # If we can't parse the date, process it anyway to be safe
                pass
            
            processed_count += 1
            
            print(f"\n📧 Processing: {subject}")
            print(f"   From: {from_email}")
            print(f"   Date: {date}")
            
            # Check for attachments
            excel_found = False
            
            for part in email_message.walk():
                # Check if it's an attachment
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    
                    if filename:
                        filename = decode_email_subject(filename)
                        
                        # Check if it's an Excel file
                        if filename.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                            excel_found = True
                            print(f"   📎 Found Excel attachment: {filename}")
                            
                            # Get file data
                            file_data = part.get_payload(decode=True)
                            
                            # Scan for name
                            matches = scan_excel_for_name(file_data, search_name, filename)
                            
                            if matches:
                                total_matches += len(matches)
                                print(f"   🎯 Found {len(matches)} match(es) for '{search_name}'!")
                                
                                # Prepare Telegram message
                                message = f"🔔 <b>Name Match Found!</b>\n\n"
                                message += f"📧 <b>Email:</b> {subject}\n"
                                message += f"👤 <b>From:</b> {from_email}\n"
                                message += f"📅 <b>Date:</b> {date}\n"
                                message += f"📎 <b>File:</b> {filename}\n"
                                message += f"🔍 <b>Search term:</b> {search_name}\n\n"
                                message += f"<b>Matches ({len(matches)}):</b>\n"
                                
                                for i, match in enumerate(matches[:10], 1):  # Limit to 10 matches
                                    message += f"{i}. Sheet: {match['sheet']}, Cell: {match['cell']}\n"
                                    message += f"   Value: {match['value'][:100]}\n"
                                
                                if len(matches) > 10:
                                    message += f"\n... and {len(matches) - 10} more match(es)"
                                
                                # Send Telegram notification
                                send_telegram_message(bot_token, chat_id, message)
                            else:
                                print(f"   ℹ️ No matches for '{search_name}' in {filename}")
            
            if not excel_found:
                print(f"   ℹ️ No Excel attachments found")
                
        except Exception as e:
            print(f"✗ Error processing email {email_id}: {e}")
            continue
    
    print(f"\n{'='*60}")
    print(f"Processed {processed_count} recent email(s)")
    print(f"Scan complete: {total_matches} total match(es) found")
    print(f"{'='*60}\n")


def main():
    """Main function to run the email monitor."""
    
    print(f"\n{'='*60}")
    print(f"Email Excel Monitor - Started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}\n")
    
    # Get environment variables
    email_address = os.getenv('EMAIL_ADDRESS')
    app_password = os.getenv('EMAIL_APP_PASSWORD')
    search_name = os.getenv('SEARCH_NAME')
    bot_token = os.getenv('TELEGRAM_BOT_TOKEN')
    chat_id = os.getenv('TELEGRAM_CHAT_ID')
    imap_server = os.getenv('IMAP_SERVER', 'imap.gmail.com')
    
    # Time window for checking emails (default 10 minutes)
    try:
        check_last_minutes = int(os.getenv('CHECK_LAST_MINUTES', '10'))
    except ValueError:
        check_last_minutes = 10
    
    # Validate required variables
    required_vars = {
        'EMAIL_ADDRESS': email_address,
        'EMAIL_APP_PASSWORD': app_password,
        'SEARCH_NAME': search_name,
        'TELEGRAM_BOT_TOKEN': bot_token,
        'TELEGRAM_CHAT_ID': chat_id
    }
    
    missing_vars = [var for var, value in required_vars.items() if not value]
    
    if missing_vars:
        print(f"✗ Missing required environment variables: {', '.join(missing_vars)}")
        print("\nPlease configure these in GitHub Secrets:")
        for var in missing_vars:
            print(f"  - {var}")
        sys.exit(1)
    
    print(f"Configuration:")
    print(f"  Email: {email_address}")
    print(f"  IMAP Server: {imap_server}")
    print(f"  Search Name: {search_name}")
    print(f"  Telegram Chat ID: {chat_id}")
    print(f"  Time Window: Last {check_last_minutes} minutes")
    print()
    
    # Connect to email
    mail = connect_to_email(email_address, app_password, imap_server)
    
    try:
        # Process emails
        process_emails(mail, search_name, bot_token, chat_id, check_last_minutes)
    finally:
        # Logout
        try:
            mail.logout()
            print("✓ Logged out from email")
        except:
            pass


if __name__ == "__main__":
    main()