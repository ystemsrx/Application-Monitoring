import sys
import struct
import zlib
import time
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QFileDialog, QComboBox, QPushButton
from PyQt5.QtCore import Qt

key_encoding = {
    'a': 0x61, 'b': 0x62, 'c': 0x63, 'd': 0x64, 'e': 0x65,
    'f': 0x66, 'g': 0x67, 'h': 0x68, 'i': 0x69, 'j': 0x6A,
    'k': 0x6B, 'l': 0x6C, 'm': 0x6D, 'n': 0x6E, 'o': 0x6F,
    'p': 0x70, 'q': 0x71, 'r': 0x72, 's': 0x73, 't': 0x74,
    'u': 0x75, 'v': 0x76, 'w': 0x77, 'x': 0x78, 'y': 0x79,
    'z': 0x7A,
    '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34,
    '5': 0x35, '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39,
    'F1': 0x70, 'F2': 0x71, 'F3': 0x72, 'F4': 0x73, 'F5': 0x74,
    'F6': 0x75, 'F7': 0x76, 'F8': 0x77, 'F9': 0x78, 'F10': 0x79,
    'F11': 0x7A, 'F12': 0x7B,
    'space': 0x20, 'enter': 0x0D, 'backspace': 0x08, 'tab': 0x09,
    'escape': 0x1B, 'ctrl': 0x81, 'shift': 0x82, 'alt': 0x83,
    'caps_lock': 0x84, 'num_lock': 0x85, 'scroll_lock': 0x86,
    '-': 0x2D, '=': 0x3D, '[': 0x5B, ']': 0x5D, '\\': 0x5C,
    ';': 0x3B, '\'': 0x27, ',': 0x2C, '.': 0x2E, '/': 0x2F,
    '!': 0x21, '@': 0x40, '#': 0x23, '$': 0x24, '%': 0x25,
    '^': 0x5E, '&': 0x26, '*': 0x2A, '(': 0x28, ')': 0x29,
    '_': 0x5F, '+': 0x2B, '{': 0x7B, '}': 0x7D, ':': 0x3A,
    '"': 0x22, '<': 0x3C, '>': 0x3E, '?': 0x3F, '|': 0x7C,
    '~': 0x7E, '《': 0x300A, '》': 0x300B, '？': 0xFF1F, 
    '：': 0xFF1A, '“': 0x201C, '”': 0x201D, '｛': 0xFF5B, 
    '｝': 0xFF5D, '——': 0x2014, '+': 0x2B, '~': 0x7E, 
    '！': 0xFF01, '￥': 0xFFE5, '%': 0xFF05, '…': 0x2026, 
    '＆': 0xFF06, '*': 0xFF0A, '（': 0xFF08, '）': 0xFF09,
    'insert': 0x90, 'delete': 0x91, 'home': 0x92, 'end': 0x93,
    'page_up': 0x94, 'page_down': 0x95, 'arrow_up': 0x96,
    'arrow_down': 0x97, 'arrow_left': 0x98, 'arrow_right': 0x99,
    'numpad_decimal': 0x6E, 'numpad_add': 0x6B,
    'numpad_subtract': 0x6D, 'numpad_multiply': 0x6A, 'numpad_divide': 0x6F,
    'print_screen': 0x9A, 'pause': 0x9B, 'menu': 0x9C, 'windows': 0x9D
}

def read_compressed_file(filename):
    window_title_dict = {}
    reverse_window_title_dict = {}
    data = []
    
    try:
        with open(filename, 'rb') as file:
            compressed_data = file.read()
            print(f"Compressed data size: {len(compressed_data)} bytes")  # 调试输出
            raw_data = zlib.decompress(compressed_data)
            print(f"Decompressed data size: {len(raw_data)} bytes")  # 调试输出

            i = 0
            while i < len(raw_data):
                if i + 1 >= len(raw_data): break  # 检查是否越界
                title_index = raw_data[i]
                i += 1

                if title_index not in reverse_window_title_dict:
                    if i >= len(raw_data): break  # 检查是否越界
                    title_len = raw_data[i]
                    i += 1
                    if i + title_len > len(raw_data): break  # 检查是否越界
                    title = raw_data[i:i+title_len].decode('utf-8')
                    i += title_len
                    reverse_window_title_dict[title_index] = title
                else:
                    title = reverse_window_title_dict[title_index]

                if i >= len(raw_data): break  # 检查是否越界
                key_code = raw_data[i]
                i += 1

                if i >= len(raw_data): break  # 检查是否越界
                count = raw_data[i]
                i += 1

                if i + 8 > len(raw_data): break  # 检查是否越界
                timestamp, = struct.unpack('d', raw_data[i:i+8])
                i += 8
                formatted_time = time.strftime('%Y/%m/%d %H:%M:%S', time.localtime(timestamp))

                key = [k for k, v in key_encoding.items() if v == key_code]
                if key:
                    key = key[0]
                else:
                    key = 'Unknown'  # 如果没有找到对应的按键编码

                data.append((title, key, count, formatted_time))
            
            print(f"Loaded {len(data)} entries")  # 调试输出
    except Exception as e:
        print(f"Error reading file: {e}")
    
    return data

class DataViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Data Viewer')
        self.setGeometry(300, 300, 800, 600)

        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(4)
        self.table_widget.setHorizontalHeaderLabels(["应用", "按键", "数量", "时间"])
        self.table_widget.horizontalHeader().setStretchLastSection(True)

        self.app_selector = QComboBox()
        self.app_selector.addItem("All")
        self.app_selector.currentTextChanged.connect(self.filter_data)

        self.export_button = QPushButton("导出")
        self.export_button.clicked.connect(self.export_data)

        layout = QVBoxLayout()
        layout.addWidget(self.app_selector)
        layout.addWidget(self.table_widget)
        layout.addWidget(self.export_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.setAcceptDrops(True)
        self.data = []
        self.filtered_data = []  # 用于存储当前显示的数据

    def load_data(self, file_name):
        self.data = read_compressed_file(file_name)
        self.update_app_selector()
        self.display_data(self.data)

    def update_app_selector(self):
        apps = sorted(set([row[0] for row in self.data]))
        self.app_selector.clear()
        self.app_selector.addItem("All")
        self.app_selector.addItems(apps)

    def filter_data(self):
        selected_app = self.app_selector.currentText()
        if selected_app == "All":
            self.filtered_data = self.data
        else:
            self.filtered_data = [row for row in self.data if row[0] == selected_app]
        self.display_data(self.filtered_data)

    def display_data(self, data):
        self.table_widget.setRowCount(len(data))
        for row, (title, key, count, timestamp) in enumerate(data):
            item_title = QTableWidgetItem(title)
            item_title.setTextAlignment(Qt.AlignCenter)
            self.table_widget.setItem(row, 0, item_title)
            
            item_key = QTableWidgetItem(key)
            item_key.setTextAlignment(Qt.AlignCenter)
            self.table_widget.setItem(row, 1, item_key)
            
            item_count = QTableWidgetItem(str(count))
            item_count.setTextAlignment(Qt.AlignCenter)
            self.table_widget.setItem(row, 2, item_count)
            
            item_timestamp = QTableWidgetItem(timestamp)
            item_timestamp.setTextAlignment(Qt.AlignCenter)
            self.table_widget.setItem(row, 3, item_timestamp)

        self.table_widget.resizeColumnsToContents()

    def export_data(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "导出数据", "", "CSV Files (*.csv);;Excel Files (*.xlsx)", options=options)
        if file_name:
            df = pd.DataFrame(self.filtered_data, columns=["应用", "按键", "数量", "时间"])
            if file_name.endswith('.csv'):
                df.to_csv(file_name, index=False, encoding='utf-8-sig')
            elif file_name.endswith('.xlsx'):
                df.to_excel(file_name, index=False)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_name = url.toLocalFile()
            self.load_data(file_name)

def main():
    app = QApplication(sys.argv)
    viewer = DataViewer()
    viewer.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
