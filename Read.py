import sys
import struct
import zlib
import time
import pandas as pd
from datetime import datetime, timezone, timedelta
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, 
                             QFileDialog, QComboBox, QPushButton, QHeaderView, QTextEdit, QSizePolicy)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QTextOption

def read_compressed_file(filename):
    data = []
    
    try:
        with open(filename, 'rb') as file:
            compressed_data = file.read()
            print(f"Compressed data size: {len(compressed_data)} bytes")
            raw_data = zlib.decompress(compressed_data)
            print(f"Decompressed data size: {len(raw_data)} bytes")

            index = 0
            while index < len(raw_data):
                try:
                    data_type = raw_data[index]
                    index += 1
                    if data_type == 1:  # 按键数据
                        title_length = struct.unpack('I', raw_data[index:index+4])[0]
                        index += 4
                        window_title = raw_data[index:index+title_length].decode('utf-8')
                        index += title_length

                        key_length = struct.unpack('I', raw_data[index:index+4])[0]
                        index += 4
                        key_name = raw_data[index:index+key_length].decode('utf-8')
                        index += key_length

                        count = struct.unpack('I', raw_data[index:index+4])[0]
                        index += 4

                        timestamp = struct.unpack('<d', raw_data[index:index+8])[0]
                        index += 8

                        # 将 UTC 时间转换为 UTC+8
                        utc_time = datetime.fromtimestamp(timestamp, tz=timezone.utc)
                        beijing_time = utc_time.astimezone(timezone(timedelta(hours=8)))
                        formatted_time = beijing_time.strftime('%Y/%m/%d %H:%M:%S')
                        
                        # 只保留冒号前的应用名称
                        app_name = window_title.split(':')[0]
                        data.append((app_name, key_name, count, formatted_time))
                except struct.error as e:
                    print(f"Warning: Corrupt data at index {index}, error: {e}")
                    index += 1

            print(f"Loaded {len(data)} entries")
    except Exception as e:
        print(f"Error reading file: {e}")
    
    return data

class DataViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Data Viewer')
        self.setGeometry(300, 300, 1000, 600)

        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(4)
        self.table_widget.setHorizontalHeaderLabels(["应用", "按键", "数量", "时间"])
        
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)

        self.app_selector = QComboBox()
        self.app_selector.addItem("All")
        self.app_selector.currentTextChanged.connect(self.filter_data)

        # Corrected key display box
        self.key_display = QTextEdit()
        self.key_display.setReadOnly(True)
        self.key_display.setLineWrapMode(QTextEdit.WidgetWidth)
        self.key_display.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.key_display.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.key_display.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.key_display.setMinimumHeight(self.key_display.fontMetrics().lineSpacing() * 3)
        self.key_display.setMaximumHeight(self.key_display.fontMetrics().lineSpacing() * 10)

        self.export_button = QPushButton("导出")
        self.export_button.clicked.connect(self.export_data)

        layout = QVBoxLayout()
        layout.addWidget(self.app_selector)
        layout.addWidget(self.table_widget)
        layout.addWidget(self.key_display)
        layout.addWidget(self.export_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.setAcceptDrops(True)
        self.data = []
        self.filtered_data = []

        self.resize_timer = QTimer(self)
        self.resize_timer.timeout.connect(self.adjust_column_widths)
        self.resize_timer.start(100)

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
        keys_with_count = []
        for row, (title, key, count, timestamp) in enumerate(data):
            for col, item in enumerate([title, key, str(count), timestamp]):
                table_item = QTableWidgetItem(item)
                table_item.setTextAlignment(Qt.AlignCenter)
                self.table_widget.setItem(row, col, table_item)
            keys_with_count.append((key, count))

        self.adjust_column_widths()
        self.update_key_display(keys_with_count)

    def update_key_display(self, keys_with_count):
        all_keys = []
        for key, count in keys_with_count:
            formatted_key = key if len(key) == 1 else f"*{key}*"
            all_keys.extend([formatted_key] * int(count))
        self.key_display.setText("".join(all_keys))

    def adjust_column_widths(self):
        header = self.table_widget.horizontalHeader()
        total_width = self.table_widget.viewport().width()
        column_widths = [0] * self.table_widget.columnCount()

        # Calculate the content width for each column
        for col in range(self.table_widget.columnCount()):
            header_width = self.table_widget.fontMetrics().width(self.table_widget.horizontalHeaderItem(col).text()) + 20
            column_widths[col] = max(column_widths[col], header_width)
            for row in range(self.table_widget.rowCount()):
                item = self.table_widget.item(row, col)
                if item:
                    column_widths[col] = max(column_widths[col], 
                                             self.table_widget.fontMetrics().width(item.text()) + 20)  # Add 20px padding

        # Ensure minimum width for each column
        column_widths = [max(width, 50) for width in column_widths]  # Minimum width of 50 pixels

        # Adjust widths based on available space
        total_content_width = sum(column_widths)
        if total_content_width > 0:
            available_width = max(total_width - total_content_width, 0)
            for col in range(self.table_widget.columnCount()):
                extra = int(available_width * (column_widths[col] / total_content_width))
                column_widths[col] += extra

        # Set the calculated widths
        for col, width in enumerate(column_widths):
            header.setSectionResizeMode(col, QHeaderView.Interactive)
            self.table_widget.setColumnWidth(col, width)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.adjust_column_widths()

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