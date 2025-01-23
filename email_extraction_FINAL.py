import mailbox
import os
from email.utils import parsedate_to_datetime, mktime_tz, parsedate_tz
from datetime import timezone, datetime, timedelta
from colorama import init, Fore, Style

# Initialize colorama
init(autoreset=True)

# ANSI escape codes for formatting using colorama
BOLD = Style.BRIGHT
RESET = Style.RESET_ALL
BIG_FONT = Fore.YELLOW + Style.BRIGHT

# Paths and constants
processing_fodler =r"C:\Users\tomasz.plewka\AppData\Roaming\Thunderbird\Profiles\fws18x1p.default-release\ImapMail\mail.selt.com\INBOX.sbd\Processing.sbd"
main_inbox_folder:r"C:\Users\tomasz.plewka\AppData\Roaming\Thunderbird\Profiles\fws18x1p.default-release\ImapMail\mail.selt.com\INBOX.sbd"
save_root_path = r"M:\TECZKI KLIENTÓW"
INTERVAL = 1  # Interval in minutes

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
if not os.path.exists(save_root_path):
    os.makedirs(save_root_path)

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

def move_to_folders(msg, mbox_path):
    """Move an email to another mbox folder"""
    with mailbox.mbox(mbox_path) as mbox:
        mbox.lock()
        mbox.add()
        mbox.flush()
        mbox.unlock()
    print(f"Moved email to folder: {mbox_path}")


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

def process_email(msg):
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

def clear_processing_folder(folder_path):
    if os.path.exists (folder_path):
        shutil.rmtree(folder_path)
        os.makedirs(folder_path)
        print("Cleared the processing folder")

def process_mbox(processing_folder, main_inbox_folder, save_root_path):
    try:
        mbox = mailbox.mbox(processing_folder)
        if len(mbox) == 0:
            print ("no emails found in processing folder")
            return
    for msg in mbox:
        try:
            sender, order_numer, order_name = process_email(msg)
            if not order_numer:
                move_to_folders(msg, main_inbox_folder)
                continue
            client_folder = CLIENT_FOLDER_MAPPING.get(order_name.upper(), order_name)
            client_folder_path = os.path.join(save_root_path, client_folder)
            if not os.path.exists(client_folder_path):
                os.makedirs(client_folder_path)
            save_eml(msg,client_folder_path, f"Order_{order_numer}_{order_name}")
            except Exception as e:
            print(f"Error processing email: {e}")

        clear_processing_folder(processing_folder)
        except Exception as e:
        print(f"Error processing folder: {e}")

def print_time_remaining(last_run_time, interval_minutes):
    """Print the remaining time dynamically until the next run."""
    next_run_time = last_run_time + timedelta(minutes=interval_minutes)
    while True:
        remaining_time = next_run_time - datetime.now()
        if remaining_time.total_seconds() <= 0:
            break
        mins, secs = divmod(remaining_time.seconds, 60)
        print(f"\rNext run in: {mins:02d}:{secs:02d} minutes", end="")
    print("\nNext run starting now!")

if __name__ == "__main__":
    try:
        while True:
            start_time = datetime.now()
            print_time_remaining(start_time, INTERVAL)
    except KeyboardInterrupt:
        print("\nProcess terminated by user.")