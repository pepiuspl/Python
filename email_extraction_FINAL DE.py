import mailbox
import os
import re
from email.utils import parsedate_tz, mktime_tz
from datetime import datetime, timedelta
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
TIME_WINDOW_MINUTES = 4320
INTERVAL = 30  # Interval in minutes

CLIENT_FOLDER_MAPPING = {
    "ZOUNEK": "ZOUNEK DESIGN",
    "INTERIER TENDE": "INTERIER-TENDDE-JASMIN ALIHODIC",
    "TT GRADNJA": "TT-GRADNJA",
    "ALU-PROFI": "ALU_PROFI BZ",
    "ALU_PROFI 2003KG": "ALU-PROFI 2003KG",
    "BORDASROLO KFT": "BORDASROLO",
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

def email_is_new(msg, time_window_minutes, processed_emails):
    """Check if the email is within the time window and not already processed."""
    current_time = datetime.now()
    received = msg.get('Date')
    if received:
        email_time = parsedate_tz(received)
        if email_time:
            email_timestamp = mktime_tz(email_time)
            email_time_dt = datetime.fromtimestamp(email_timestamp)
            if current_time - email_time_dt <= timedelta(minutes=time_window_minutes):
                email_id = msg.get("Message-ID", "").strip()
                if email_id not in processed_emails:
                    processed_emails.add(email_id)
                    return True
    return False

def get_all_folders(base_path):
    """Recursively find all folders and subfolders under the given base path."""
    folders = []
    for root, dirs, files in os.walk(base_path):
        for file in files:
            if not file.endswith('.msf'):
                folders.append(os.path.join(root, file))
    return folders

def process_folders(base_path, time_window_minutes):
    """Recursively process folders and their subfolders."""
    processed_emails = set()
    folders = get_all_folders(base_path)

    if not folders:
        print(f"No folders found in base path: {base_path}")
        return

    for folder_path in folders:
        try:
            process_mbox(folder_path, SAVE_ROOT, time_window_minutes, processed_emails)
        except Exception as e:
            print(f"Error processing folder {folder_path}: {e}")

def extract_order_details(subject, content):
    """Extract order number and name from German email subject and content."""
    order_number = None
    order_name = None

    # Search for order number in subject (in German: Bestellung Nr.)
    number_match = re.search(r'Bestellung\s*(Nr\.)?\s*(\d+)', subject, re.IGNORECASE)
    if number_match:
        order_number = number_match.group(2)

    # If no order number found in the subject, try extracting from the content
    if not order_number:
        number_match = re.search(r'Bestellung\s*(Nr\.)?\s*(\d+)', content, re.IGNORECASE)
        if number_match:
            order_number = number_match.group(2)

    # Extract order name from content (German terms: Firma, Kunde)
    name_match = re.search(r'(Firma|Kunde)[\s:]*([^<\n]+)', content, re.IGNORECASE)
    if name_match:
        order_name = name_match.group(2).strip()

    return order_number, order_name

def process_mbox(folder_path, save_root_path, time_window_minutes, processed_emails):
    """Process Thunderbird folder and save emails."""
    if not os.path.exists(folder_path):
        print(f"Folder {folder_path} not found!")
        return

    relative_path = folder_path.replace(THUNDERBIRD_PATH, '').strip(os.sep).split(os.sep)
    relative_path = [folder.rstrip('.sbd') for folder in relative_path]

    if len(relative_path) < 2:
        print(f"Invalid folder structure: {folder_path}")
        return

    region, client = relative_path[0], relative_path[1]
    print(f"{BIG_FONT}Processing emails for client: {client} in region: {region}{RESET}")

    save_path = os.path.join(save_root_path, "______2025", region, client, "1_ZAMÓWIENIA")
    if not os.path.exists(save_path):
        os.makedirs(save_path)

    try:
        mbox = mailbox.mbox(folder_path)
        if len(mbox) == 0:
            print(f"No emails in folder: {folder_path}")

        for msg in mbox:
            try:
                subject = msg.get("subject", "No Subject")
                content = msg.get_payload(decode=True)

                # Decode and safeguard against None content
                if content:
                    content = content.decode(errors="ignore")
                else:
                    content = ""  # Default to an empty string

                # Check if the email is in German (contains German keywords in subject or content)
                if 'Bestellung' in subject or 'Bestellung' in content:
                    if email_is_new(msg, time_window_minutes, processed_emails):
                        print(f"Processing German email with subject: {BOLD}{subject}{RESET}")
                        order_number, order_name = extract_order_details(subject, content)

                        if order_number:
                            folder_name = f"Order_{order_number}_{sanitize_filename(order_name or 'Unknown')}"
                        else:
                            folder_name = "Unknown_Order"

                        # Define the save path for the specific order folder
                        save_path_order = os.path.join(save_path, folder_name)

                        # Create the folder if it doesn't exist
                        if not os.path.exists(save_path_order):
                            os.makedirs(save_path_order)

                        save_eml(msg, save_path_order, subject)
                    else:
                        print(f"Skipping email: {subject}")
                else:
                    print(f"Skipping non-German email: {subject}")
            except Exception as e:
                print(f"Error processing email: {e}")
    except Exception as e:
        print(f"Error opening the mbox file {folder_path}: {e}")

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

if __name__ == "__main__":
    try:
        while True:
            start_time = datetime.now()
            process_folders(THUNDERBIRD_PATH, TIME_WINDOW_MINUTES)
            print_time_remaining(start_time, INTERVAL)
    except KeyboardInterrupt:
        print("\nProcess terminated by user.")
    except Exception as e:
        print(f"Error during processing: {e}")
