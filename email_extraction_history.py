import mailbox
import os
import re
from email.utils import parsedate_to_datetime, mktime_tz, parsedate_tz
from datetime import timezone, datetime, timedelta
import time
from colorama import init, Fore, Style

# Initialize colorama
init(autoreset=True)

# ANSI escape codes for formatting using colorama
BOLD = Style.BRIGHT
RESET = Style.RESET_ALL
BIG_FONT = Fore.YELLOW + Style.BRIGHT

# Paths and constants
THUNDERBIRD_PATH = r"C:\Users\tomasz.plewka\AppData\Roaming\Thunderbird\Profiles\fws18x1p.default-release\ImapMail\mail.selt.com\INBOX.sbd\2025.sbd\Zam&APM-wienia.sbd"
SAVE_ROOT = r"M:\TECZKI KLIENTÓW"
INTERVAL = 1  # Interval in minutes

# Location for saving email history
HISTORY_FILE_PATH = r"C:\Users\tomasz.plewka\desktop\Python\email_history.txt"

CLIENT_FOLDER_MAPPING = {
    "ZOUNEK": "ZOUNEK DESIGN",
    "INTERIER TENDE": "INTERIER-TENDDE-JASMIN ALIHODIC",
    "TT GRADNJA": "TT GRADNJA",
    "ALU_PROFI BZ": "ALU-PROFI",
    "ALU_PROFI 2003KG": "ALU-PROFI 2003KG",
    "BORDASROLO": "BORDASROLO KFT",
    "RED_REL": "REDREL",
    "ROLLSTAR": "ROLLSTAR KFT",
    "SKM PROTECT": "SKM_PROTECT",
    "SPANNO": "SPANNO KFT",
    "IVPA": "IVPA OKNA",
}

# Ensure the save root folder exists
if not os.path.exists(SAVE_ROOT):
    os.makedirs(SAVE_ROOT)

def sanitize_filename(string):
    """Sanitize a string to be a valid filename."""
    return re.sub(r'[\\/*?:"<>|]', "_", string)

def save_eml(msg, save_path, filename):
    """Save the email as a .eml file."""
    sanitized_filename = sanitize_filename(filename)
    filepath = os.path.join(save_path, f"{sanitized_filename}.eml")
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(msg.as_string())
    print(f"Saved: {filepath}")

def get_last_processed_email():
    """Read the last processed email's ID and timestamp from the history file."""
    processed_emails = set()
    last_received_date = None

    if os.path.exists(HISTORY_FILE_PATH):
        with open(HISTORY_FILE_PATH, "r") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                parts = line.split(',')
                if len(parts) == 3:
                    try:
                        email_id, processing_timestamp, email_received_time = parts
                        processing_timestamp = datetime.fromisoformat(processing_timestamp)
                        email_received_time = datetime.fromisoformat(email_received_time)
                        processed_emails.add((email_id, processing_timestamp, email_received_time))
                    except ValueError:
                        print(f"Skipping malformed line in history file: {line}")
                        continue
                else:
                    print(f"Skipping malformed line in history file: {line}")

    # Determine the last received date based on the processed emails
    if processed_emails:
        last_received_date = max(email[2] for email in processed_emails)  # Most recent email_received_time
    else:
        last_received_date = None  # No emails processed yet

    return processed_emails, last_received_date

def save_email_history(processed_emails):
    """Save processed email IDs with timestamps to the history file."""
    try:
        with open(HISTORY_FILE_PATH, "w") as f:
            for email_id, timestamp, email_received_time in sorted(processed_emails, key=lambda x: x[2]):
                f.write(f"{email_id},{timestamp.isoformat()},{email_received_time.isoformat()}\n")
        print(f"History saved with {len(processed_emails)} processed emails.")
    except Exception as e:
        print(f"Error saving email history: {e}")

def email_is_new(msg, processed_emails, last_timestamp):
    """Check if the email is new (not already processed based on timestamp)."""
    email_id = msg.get("Message-ID", "").strip()
    date_tuple = msg.get("date")

    # Parse email received time
    if date_tuple:
        email_received_time = datetime.fromtimestamp(mktime_tz(parsedate_tz(date_tuple)))
    else:
        email_received_time = datetime.now()

    # Log the comparison
    print(f"Checking email: {email_id} - Received time: {email_received_time} vs Last processed time: {last_timestamp}")

    # Check if the email is newer than the last processed timestamp
    if email_received_time > last_timestamp:
        # Add email to processed set
        processed_emails.add((email_id, datetime.now(), email_received_time))
        return True

    # If email is too old or already processed
    print(f"Email {email_id} is too old or already processed.")
    return False

def extract_order_details_from_subject(subject):
    """Extract order number from the email subject."""
    match = re.search(r"(Nr\.|no\.)\s*(\d+)", subject, re.IGNORECASE)
    if match:
        order_number = match.group(2)
        return order_number
    return None

def extract_order_name_from_content(content):
    """Extract order name from the email content (text between parentheses)."""
    match = re.search(r"\((.*?)\)", content)
    if match:
        order_name = match.group(1).strip()
        if order_name in ["230, 0, 0", ""]:
            return None
        return order_name
    return None

def process_email(msg, processed_emails):
    """Process email to extract order number and order name from a specific sender."""
    sender = msg.get("From", "")
    if "noreply@selt.com" not in sender:
        print(f"Skipping email from {sender}")
        return None, None  # Skip emails from other senders
    
    # Extract order details from the subject and content
    subject = msg.get("subject", "")
    order_number = extract_order_details_from_subject(subject)
    if not order_number:
        print(f"Could not extract order details from subject: {subject}")
        return None, None
    
    content = get_email_content(msg)
    order_name = extract_order_name_from_content(content)
    if not order_name:
        order_name = None  # If no order name, set to None
    
    # Add the email to the processed_emails set
    email_id = msg.get("Message-ID", "").strip()
    received_time = get_email_received_time(msg)

    if not email_id:
        print(f"Skipping email with missing Message-ID. Subject: {subject}")
        return None, None

    processed_emails.add((email_id, datetime.now(), received_time))
    return order_number, order_name

def get_email_received_time(msg):
    """
    Extract the received time from the email.
    Tries to use the 'Received' header first, then falls back to 'Date'.
    """
    try:
        received_header = msg.get_all("Received")
        if received_header:
            # Parse the most recent 'Received' header (the last server processing timestamp)
            last_received_line = received_header[-1]
            timestamp_start = last_received_line.find(";")
            if timestamp_start != -1:
                timestamp_str = last_received_line[timestamp_start + 1:].strip()
                received_time = parsedate_to_datetime(timestamp_str)

                # If received_time is naive, make it timezone-aware using UTC
                if received_time.tzinfo is None:
                    received_time = received_time.replace(tzinfo=timezone.utc)

                return received_time
    except Exception as e:
        print(f"Error parsing 'Received' header: {e}")

    try:
        date_header = msg.get("Date", None)
        if date_header:
            received_time = parsedate_to_datetime(date_header)

            # If received_time is naive, make it timezone-aware using UTC
            if received_time.tzinfo is None:
                received_time = received_time.replace(tzinfo=timezone.utc)

            return received_time
    except Exception as e:
        print(f"Error parsing 'Date' header: {e}")

    print("Could not extract a received time for the email.")
    return None

def get_email_content(msg):
    """Extract text content from an email, handling different content types."""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                return part.get_payload(decode=True).decode("utf-8", errors="ignore")
            elif content_type == "text/html":
                return part.get_payload(decode=True).decode("utf-8", errors="ignore")
    else:
        return msg.get_payload(decode=True).decode("utf-8", errors="ignore")
    return None

def process_mbox(folder_path, save_root_path, last_received_date, processed_emails):
    """Process Thunderbird folder and save emails."""
    if not os.path.exists(folder_path):
        print(f"Folder {folder_path} not found!")
        return False

    relative_path = folder_path.replace(THUNDERBIRD_PATH, '').strip(os.sep).split(os.sep)
    relative_path = [folder.rstrip('.sbd') for folder in relative_path]

    if len(relative_path) < 2:
        print(f"Invalid folder structure: {folder_path}")
        return False

    region, client = relative_path[0], relative_path[1]
    print(f"{BIG_FONT}Processing emails for client: {client} in region: {region}{RESET}")

    if client in CLIENT_FOLDER_MAPPING:
        client = CLIENT_FOLDER_MAPPING[client]
    else:
        print(f"Client {client} not found in mappings! Using original client name.")

    save_path = os.path.join(save_root_path, "______2025", region, client, "1_ZAMÓWIENIA")

    print(f"Save Path: {save_path}")
    new_email_found = False  # Track if new emails were processed

    try:
        mbox = mailbox.mbox(folder_path)
        if len(mbox) == 0:
            print(f"No emails in folder: {folder_path}")

        for msg in mbox:
            try:
                subject = msg.get("subject", "No Subject")
                received_time = get_email_received_time(msg)

                # If received_time is naive, make it timezone-aware using UTC
                if received_time and received_time.tzinfo is None:
                    received_time = received_time.replace(tzinfo=timezone.utc)

                # If last_received_date is naive, make it timezone-aware using UTC
                if last_received_date and last_received_date.tzinfo is None:
                    last_received_date = last_received_date.replace(tzinfo=timezone.utc)

                # **ONLY process emails received after the last processed email timestamp**
                if received_time > last_received_date:
                    print(f"Processing email with subject: {BOLD}{subject}{RESET}")
                    order_number, order_name = process_email(msg, processed_emails)
                    new_email_found = True

            except Exception as e:
                print(f"Error processing email: {e}")
    except Exception as e:
        print(f"Error opening the mbox file {folder_path}: {e}")

    return new_email_found

def print_time_remaining(last_run_time, interval_minutes):
    """Print the remaining time dynamically until the next run."""
    next_run_time = last_run_time + timedelta(minutes=interval_minutes)
    while True:
        remaining_time = next_run_time - datetime.now()
        if remaining_time.total_seconds() <= 0:
            break
        mins, secs = divmod(remaining_time.seconds, 60)
        print(f"\rNext run in: {mins:02d}:{secs:02d} minutes", end="")
        time.sleep(1)
    print("\nNext run starting now!")

def get_all_folders(base_path):
    """Recursively find all folders and subfolders under the given base path."""
    folders = []
    for root, dirs, files in os.walk(base_path):
        for file in files:
            if not file.endswith('.msf'):  # Skip .msf files
                folders.append(os.path.join(root, file))
    return folders

def process_folders(base_path):
    processed_emails, last_received_date = get_last_processed_email()
    if last_received_date is None:
        print("No history found. Processing all emails.")
        last_received_date = datetime.min

    try:
        folders = get_all_folders(base_path)
        print(f"Found {len(folders)} folders to process.")
        if not folders:
            print(f"No folders found in base path: {base_path}")
            return

        new_email_found = False
        for folder_path in folders:
            if process_mbox(folder_path, SAVE_ROOT, last_received_date, processed_emails):
                new_email_found = True

        if new_email_found:
            save_email_history(processed_emails)
        else:
            print("No new emails received.")
    except Exception as e:
        print(f"Error processing folders: {e}")

if __name__ == "__main__":
    try:
        while True:
            start_time = datetime.now()
            process_folders(THUNDERBIRD_PATH)
            print_time_remaining(start_time, INTERVAL)
    except KeyboardInterrupt:
        print("\nProcess terminated by user.")
