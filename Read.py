import sys
import struct
import zlib
import time
import pandas as pd
from datetime import datetime, timezone, timedelta
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QFileDialog, QComboBox, QPushButton, QHeaderView
from PyQt5.QtCore import Qt

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
        
        # 初始设置表格列的伸缩模式
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)

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
            for col, item in enumerate([title, key, str(count), timestamp]):
                table_item = QTableWidgetItem(item)
                table_item.setTextAlignment(Qt.AlignCenter)
                self.table_widget.setItem(row, col, table_item)

        self.adjustTableSize()

    def adjustTableSize(self):
        # 先设置为 ResizeToContents 模式以获取内容的实际宽度
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        total_width = self.table_widget.verticalHeader().width() + 20  # 添加一些额外的空间
        for i in range(self.table_widget.columnCount()):
            total_width += self.table_widget.columnWidth(i)

        # 如果内容宽度超过当前窗口宽度，调整窗口大小
        if total_width > self.width():
            self.resize(total_width, self.height())

        # 然后将模式设置回 Stretch，以便列可以随窗口大小变化
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # 当窗口大小改变时，重新调整表格列宽
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

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
