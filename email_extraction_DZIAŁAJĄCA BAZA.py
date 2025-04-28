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
THUNDERBIRD_PATH = r"<_thunderbird_path>" #To be changed
SAVE_ROOT = r"<save_root>" #To be changed
TIME_WINDOW_MINUTES = 4320
INTERVAL = 30  # Interval in minutes

CLIENT_FOLDER_MAPPING = {
    "<name_in_thunderbird>": "<name_in_save_path>", #To be changed
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
    """Process email to extract order number and order name from specific sender."""
    sender = msg.get("From", "")
    if "<email_address_to_execute>" not in sender: #To be changed
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
    
    return order_number, order_name

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

    if client in CLIENT_FOLDER_MAPPING:
        client = CLIENT_FOLDER_MAPPING[client]
    else:
        print(f"Client {client} not found in mappings! Using original client name.")
    
    save_path = os.path.join(save_root_path, "<save_path_folder_name>", region, client, "<orders_folder_name>") #To be changed

    try:
        mbox = mailbox.mbox(folder_path)
        if len(mbox) == 0:
            print(f"No emails in folder: {folder_path}")

        for msg in mbox:
            try:
                subject = msg.get("subject", "No Subject")
                if email_is_new(msg, time_window_minutes, processed_emails):
                    print(f"Processing email with subject: {BOLD}{subject}{RESET}")

                    # Extract order number and order name
                    order_number, order_name = process_email(msg)
                    if order_number:
                        # Generate the folder name based on the order number and name
                        if order_name:
                            order_save_path = os.path.join(save_path, f"Order_{order_number}_{order_name}")
                        else:
                            order_save_path = os.path.join(save_path, f"Order_{order_number}")

                        # Check if the folder exists
                        if not os.path.exists(order_save_path):
                            os.makedirs(order_save_path)
                            print(f"Created folder: {order_save_path}")
                        else:
                            print(f"Folder already exists: {order_save_path}")

                        # Save the email
                        save_eml(msg, order_save_path, subject)
                    else:
                        print(f"Skipping email: {subject}")
                else:
                    print(f"Skipping email: {subject}")
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
