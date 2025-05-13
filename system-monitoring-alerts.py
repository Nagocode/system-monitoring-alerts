import time
import socket
import getpass
import pyautogui
from pathlib import Path
import smtplib
from email.message import EmailMessage
from datetime import datetime, timedelta
import win32evtlog
import re

# === System Information ===
hostname = socket.gethostname()
username = getpass.getuser()
current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

# === Email Configuration ===
sender_email = "example.sender@gmail.com"
receiver_email = "example.receiver@gmail.com"
app_password = "your-app-password-here"  # Use an App Password, not your actual Gmail password.

# === Function to retrieve last shutdown timestamp from Windows Event Logs ===
def get_last_shutdown_time():
    server = 'localhost'
    log_type = 'System'
    handle = win32evtlog.OpenEventLog(server, log_type)
    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    shutdown_time = None
    max_events = 200
    events_read = 0

    while events_read < max_events:
        events_read += 1
        events = win32evtlog.ReadEventLog(handle, flags, 0)
        if events:
            for event in events:
                # Event ID 1 indicates a system time change (commonly occurs at system startup)
                if event.EventID == 1:
                    if event.StringInserts and len(event.StringInserts) > 2:
                        try:
                            # Avoid false positives: check time difference is greater than 5 seconds
                            if int(event.StringInserts[2]) > 5000:
                                match = re.search(
                                    r'(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})\.(\d+Z)',
                                    event.StringInserts[1]
                                )
                                if match:
                                    year, month, day = match.group(1), match.group(2), match.group(3)
                                    time_str = f"{match.group(4)}:{match.group(5)}:{match.group(6)}.{match.group(7)[:6]}"
                                    utc_dt_str = f"{year}-{month}-{day} {time_str}"
                                    utc_dt = datetime.fromisoformat(utc_dt_str.replace('Z', '+00:00'))
                                    local_time = utc_dt - timedelta(hours=5)  # UTC-5 (e.g. Colombia)
                                    shutdown_time = local_time.strftime('%Y-%m-%d %H:%M:%S')
                                    break
                        except Exception:
                            continue
        if shutdown_time:
            break
    return shutdown_time

# === Compose and send shutdown alert email ===
last_shutdown = get_last_shutdown_time()
shutdown_email = EmailMessage()
shutdown_email["Subject"] = f"Last shutdown by {username} - {last_shutdown}"
shutdown_email["From"] = sender_email
shutdown_email["To"] = receiver_email
shutdown_email.set_content(f"The last shutdown for user {username} on device {hostname} occurred at {last_shutdown}.")

try:
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender_email, app_password)
        smtp.send_message(shutdown_email)
except Exception as e:
    print(f"Failed to send shutdown alert: {e}")

# === Compose and send login alert email ===
login_email = EmailMessage()
login_email["Subject"] = f"Login by {username} - {current_time}"
login_email["From"] = sender_email
login_email["To"] = receiver_email
login_email.set_content(f"User {username} logged into device {hostname} at {current_time}.")

try:
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender_email, app_password)
        smtp.send_message(login_email)
except Exception as e:
    print(f"Failed to send login alert: {e}")

# === Wait before taking first screenshot ===
time.sleep(5)

def take_and_send_screenshot():
    # Capture timestamp for file naming
    capture_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    filename_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    screenshot_filename = f"screenshot_{hostname}_{filename_timestamp}.png"

    # Define generic path to store screenshot
    save_path = Path(r"C:\Monitoring\Screenshots")
    save_path.mkdir(parents=True, exist_ok=True)
    screenshot_file = save_path / screenshot_filename

    # Take screenshot and save to file
    pyautogui.screenshot().save(screenshot_file)

    # Prepare screenshot alert email
    screenshot_email = EmailMessage()
    screenshot_email["Subject"] = f"Screenshot from {hostname} - {capture_timestamp}"
    screenshot_email["From"] = sender_email
    screenshot_email["To"] = receiver_email
    screenshot_email.set_content(f"Attached is a screenshot from device {hostname} captured at {capture_timestamp}.")

    # Attach screenshot file to the email
    with open(screenshot_file, "rb") as f:
        screenshot_email.add_attachment(f.read(), maintype="image", subtype="png", filename=screenshot_filename)

    # Send the screenshot alert email
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(sender_email, app_password)
            smtp.send_message(screenshot_email)
    except Exception as e:
        print(f"Failed to send screenshot: {e}")

# === Take and send two screenshots with delay ===
take_and_send_screenshot()
time.sleep(5)
take_and_send_screenshot()
