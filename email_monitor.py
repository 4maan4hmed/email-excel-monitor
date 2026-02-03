#!/usr/bin/env python3
"""
Email Excel Monitor with Telegram Notifications
Monitors email inbox for Excel attachments, scans for specified name, and sends Telegram alerts.
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
import html  # Added for escaping HTML characters

def escape_html(text):
    """Escape HTML special characters to prevent Telegram parsing errors."""
    if text:
        return html.escape(str(text))
    return ""

def send_telegram_message(bot_token, chat_id, message):
    """Send a message via Telegram bot."""
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    
    # Telegram has a hard limit of 4096 characters
    if len(message) > 4090:
        message = message[:4000] + "\n\n[Message truncated due to size limit]"

    payload = {
        'chat_id': chat_id,
        'text': message,
        'parse_mode': 'HTML'  # This will now work correctly with escaped inputs
    }
    
    try:
        response = requests.post(url, json=payload, timeout=10)
        response.raise_for_status()
        print(f"✓ Telegram notification sent successfully")
        return True
    except requests.exceptions.RequestException as e:
        print(f"✗ Failed to send Telegram message: {e}")
        # Fallback: Strip HTML tags if parsing still fails
        try:
            print("  Retrying with raw text...")
            payload['text'] = message.replace('<b>', '').replace('</b>', '').replace('<pre>', '').replace('</pre>', '')
            payload.pop('parse_mode')
            requests.post(url, json=payload, timeout=10)
        except:
            pass
        return False

def scan_excel_for_name(file_data, search_name, filename):
    """Scan all sheets in an Excel file for the search name."""
    matches = []
    
    try:
        workbook = openpyxl.load_workbook(io.BytesIO(file_data), data_only=True)
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
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

def get_email_body(msg):
    """Extract plain text body from email message."""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain" and "attachment" not in str(part.get_content_disposition()):
                try:
                    body = part.get_payload(decode=True).decode()
                    break
                except:
                    pass
    else:
        try:
            body = msg.get_payload(decode=True).decode()
        except:
            pass
    
    return body if body else "[No plain text body found]"

def connect_to_email(email_address, app_password, imap_server='imap.gmail.com'):
    """Connect to email inbox via IMAP."""
    try:
        print(f"Connecting to {imap_server}...")
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_address, app_password)
        print(f"✓ Successfully logged in as {email_address}")
        return mail
    except Exception as e:
        print(f"✗ Connection error: {e}")
        sys.exit(1)

def process_emails(mail, search_name, bot_token, chat_id, check_last_minutes=10):
    """Process only emails received in the last X minutes."""
    mail.select('INBOX')
    
    since_time = datetime.now() - timedelta(minutes=check_last_minutes)
    since_date = since_time.strftime("%d-%b-%Y")
    
    print(f"Searching for emails since {since_time.strftime('%Y-%m-%d %H:%M:%S')}")
    status, messages = mail.search(None, f'SINCE {since_date}')
    
    if status != 'OK': return
    
    email_ids = messages[0].split()
    if not email_ids:
        print(f"No emails found in the last {check_last_minutes} minutes")
        return
        
    print(f"Found {len(email_ids)} email(s)")
    
    for email_id in email_ids:
        try:
            status, msg_data = mail.fetch(email_id, '(RFC822)')
            if status != 'OK': continue
            
            raw_email = msg_data[0][1]
            email_message = email.message_from_bytes(raw_email)
            
            # Extract and Safe Decode
            subject = decode_email_subject(email_message.get('Subject'))
            from_email = email_message.get('From')
            date = email_message.get('Date')
            
            # Date Check
            try:
                from email.utils import parsedate_to_datetime
                email_datetime = parsedate_to_datetime(date)
                if email_datetime < since_time: continue
            except:
                pass
            
            print(f"\nProcessing: {subject}")
            
            # Check Attachments
            for part in email_message.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    if filename:
                        filename = decode_email_subject(filename)
                        if filename.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                            print(f"   Scanning: {filename}")
                            
                            file_data = part.get_payload(decode=True)
                            matches = scan_excel_for_name(file_data, search_name, filename)
                            
                            if matches:
                                print(f"   Found {len(matches)} matches!")
                                
                                # Prepare Data for Message (Escaping HTML characters)
                                safe_search = escape_html(search_name)
                                safe_subject = escape_html(subject)
                                safe_from = escape_html(from_email)
                                safe_date = escape_html(date)
                                safe_file = escape_html(filename)
                                
                                # Clean Body
                                raw_body = get_email_body(email_message)
                                safe_body = escape_html(" ".join(raw_body.split()))
                                
                                # --- New Formatting Layout ---
                                message = (
                                    f"<b>MATCH REPORT:</b> {safe_search}\n"
                                    f"────────────────────\n"
                                    f"<b>File:</b> {safe_file}\n"
                                    f"<b>Matches:</b> {len(matches)}\n\n"
                                    
                                    f"<b>EMAIL DETAILS</b>\n"
                                    f"Subject: {safe_subject}\n"
                                    f"From: {safe_from}\n"
                                    f"Date: {safe_date}\n\n"
                                    
                                    f"<b>MATCH DATA</b>\n"
                                )
                                
                                # List matches nicely
                                for i, match in enumerate(matches[:5], 1):
                                    safe_val = escape_html(match['value'][:50])
                                    message += f"{i}. {match['sheet']} ({match['cell']}): {safe_val}\n"
                                
                                if len(matches) > 5:
                                    message += f"... and {len(matches) - 5} more\n"
                                
                                message += (
                                    f"\n<b>CONTENT SNAPSHOT</b>\n"
                                    f"<pre>{safe_body[:500]}</pre>"
                                )

                                send_telegram_message(bot_token, chat_id, message)

        except Exception as e:
            print(f"✗ Error processing email: {e}")
            continue

def main():
    # ... (Env var loading same as before) ...
    # Quick setup for convenience
    email_address = os.getenv('EMAIL_ADDRESS')
    app_password = os.getenv('EMAIL_APP_PASSWORD')
    search_name = os.getenv('SEARCH_NAME')
    bot_token = os.getenv('TELEGRAM_BOT_TOKEN')
    chat_id = os.getenv('TELEGRAM_CHAT_ID')
    imap_server = os.getenv('IMAP_SERVER', 'imap.gmail.com')
    
    if not all([email_address, app_password, search_name, bot_token, chat_id]):
        print("✗ Missing environment variables")
        sys.exit(1)

    try:
        check_last_minutes = int(os.getenv('CHECK_LAST_MINUTES', '10'))
    except:
        check_last_minutes = 10

    print(f"--- Monitoring Started ({datetime.now().strftime('%H:%M:%S')}) ---")
    mail = connect_to_email(email_address, app_password, imap_server)
    
    try:
        process_emails(mail, search_name, bot_token, chat_id, check_last_minutes)
    finally:
        try:
            mail.logout()
            print("✓ Done")
        except:
            pass

if __name__ == "__main__":
    main()