import sys
import struct
import zlib
import time
import pandas as pd
import re
import requests
import json
from datetime import datetime, timezone, timedelta
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, 
                             QFileDialog, QComboBox, QPushButton, QHeaderView, QTextEdit, QSizePolicy, QHBoxLayout,
                             QLineEdit, QLabel, QMessageBox)
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
from PyQt5.QtGui import QTextOption

# API列表
API_KEYS = [
    "sk-xxx",
    "xxx",
    "xxx"
]

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

class ContentGenerationThread(QThread):
    content_generated = pyqtSignal(str)
    generation_finished = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def __init__(self, url, api_key, model, content):
        super().__init__()
        self.url = url
        self.api_key = api_key
        self.model = model
        self.content = content
        self.is_running = True

    def run(self):
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }
        data = {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": "你是一个解密专家，你负责解密一段文本。这段文本是程序监测键盘的行为，然后将键盘行为转换成了一段文本。比如，按下了字母键，就会是对应的字母，但若按下了空格、回退等格式化键，会显示*Space*或*Backspace*等。你需要一步一步推理，一定要写出详细逐步的推理过程，最后得到用户实际打出的文本，这些文本可能是各种语言，也可能是中文拼音，因为可以用输入法的原因，拼音可能不全或有误，比如'多少'可能会写作'duos*Space*'，也就是说先打出'duos'，然后用空格键*Space*选中输入法给出的第一个选项得到'多少'，所以你需要认真仔细分析辨别，一步一步进行分析，然后还原这一句中文内容，还原后再重新仔细检查句式表达是否有误的地方，如果有就仔细分析并修正，保证最后的内容不要有语句上的错误，并一定将最后推理出的内容放入代码块，且代码块中只包含推理出的内容原文，不要有其他任何东西，并且可以适当调整内容以保证格式、标点符号正确！"
                },
                {
                    "role": "user",
                    "content": self.content
                }
            ],
            "stream": True
        }
        try:
            response = requests.post(self.url, headers=headers, json=data, stream=True)
            response.raise_for_status()
            for line in response.iter_lines():
                if not self.is_running:
                    break
                if line:
                    line_content = line.decode('utf-8').strip()
                    if line_content.startswith("data:"):
                        line_content = line_content[len("data:"):].strip()
                        if line_content and line_content != '[DONE]':
                            try:
                                json_data = json.loads(line_content)
                                delta = json_data.get("choices", [])[0].get("delta", {}).get("content", "")
                                if delta:
                                    self.content_generated.emit(delta)
                            except json.JSONDecodeError as e:
                                print(f"Error decoding JSON: {e}")
            self.generation_finished.emit()
        except requests.RequestException as e:
            self.error_occurred.emit(str(e))

    def stop(self):
        self.is_running = False

class DataLoadingThread(QThread):
    data_loaded = pyqtSignal(list)
    error_occurred = pyqtSignal(str)

    def __init__(self, filename):
        super().__init__()
        self.filename = filename

    def run(self):
        try:
            data = read_compressed_file(self.filename)
            self.data_loaded.emit(data)
        except Exception as e:
            self.error_occurred.emit(str(e))

class DataExportThread(QThread):
    export_finished = pyqtSignal(bool)
    error_occurred = pyqtSignal(str)

    def __init__(self, data, file_name):
        super().__init__()
        self.data = data
        self.file_name = file_name

    def run(self):
        try:
            df = pd.DataFrame(self.data, columns=["应用", "按键", "数量", "时间"])
            if self.file_name.endswith('.csv'):
                df.to_csv(self.file_name, index=False, encoding='utf-8-sig')
            elif self.file_name.endswith('.xlsx'):
                df.to_excel(self.file_name, index=False)
            self.export_finished.emit(True)
        except Exception as e:
            self.error_occurred.emit(str(e))
            self.export_finished.emit(False)

class DataViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Data Viewer')
        self.setGeometry(300, 300, 1000, 800)

        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(4)
        self.table_widget.setHorizontalHeaderLabels(["应用", "按键", "数量", "时间"])
        
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)

        self.app_selector = QComboBox()
        self.app_selector.addItem("All")
        self.app_selector.currentTextChanged.connect(self.filter_data)

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

        self.generate_button = QPushButton("生成解释")
        self.generate_button.setEnabled(False)
        self.generate_button.clicked.connect(self.generate_explanation)

        self.api_key_input = QComboBox()
        self.api_key_input.setEditable(True)
        self.api_key_input.setInsertPolicy(QComboBox.NoInsert)  # 防止重复添加
        self.api_key_input.setPlaceholderText("请输入或选择有效的OpenAI、质谱清言或通义千问的API KEY")
        self.api_key_input.currentTextChanged.connect(self.validate_api_key)
        self.api_key_input.addItems(API_KEYS)

        self.model_selector = QComboBox()
        self.model_selector.hide()

        self.explanation_display = QTextEdit()
        self.explanation_display.setReadOnly(True)
        self.explanation_display.hide()

        self.expand_button = QPushButton("展开过程")
        self.expand_button.clicked.connect(self.toggle_expand)
        self.expand_button.hide()

        self.stop_button = QPushButton("终止生成")
        self.stop_button.clicked.connect(self.stop_generation)
        self.stop_button.hide()

        self.copy_button = QPushButton("复制")
        self.copy_button.clicked.connect(self.copy_explanation)
        self.copy_button.hide()

        layout = QVBoxLayout()
        layout.addWidget(self.app_selector)
        layout.addWidget(self.table_widget)
        layout.addWidget(self.key_display)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.export_button)
        button_layout.addWidget(self.generate_button)
        button_layout.addWidget(self.api_key_input)
        layout.addLayout(button_layout)

        layout.addWidget(self.model_selector)
        layout.addWidget(self.explanation_display)

        bottom_button_layout = QHBoxLayout()
        bottom_button_layout.addWidget(self.expand_button)
        bottom_button_layout.addWidget(self.stop_button)
        bottom_button_layout.addWidget(self.copy_button)
        layout.addLayout(bottom_button_layout)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.setAcceptDrops(True)
        self.data = []
        self.filtered_data = []

        self.resize_timer = QTimer(self)
        self.resize_timer.timeout.connect(self.adjust_column_widths)
        self.resize_timer.start(100)

        self.generation_thread = None
        self.full_explanation = ""
        self.is_expanded = False

    def validate_api_key(self, api_key):
        if re.match(r'^sk-[a-zA-Z0-9]{48,}$', api_key):  # OpenAI
            self.setup_model_selector('openai')
        elif re.match(r'^[a-f0-9]{32}\.[a-zA-Z0-9]{16}$', api_key):  # 质谱清言
            self.setup_model_selector('zhispu')
        elif re.match(r'^sk-[a-zA-Z0-9]{32}$', api_key):  # 通义千问
            self.setup_model_selector('tongyi')
        else:
            self.model_selector.hide()
        
        self.update_generate_button()

    def setup_model_selector(self, api_type):
        self.model_selector.clear()
        if api_type == 'openai':
            models = ['gpt-4o', 'gpt-4o-mini', 'gpt-4-turbo', 'gpt-4', 'gpt-3.5-turbo']
            default = 'gpt-4o-mini'
        elif api_type == 'zhispu':
            models = ['glm-4-0520', 'glm-4-long', 'glm-4-airx', 'glm-4-air', 'glm-4-flash', 'glm-4-alltools']
            default = 'glm-4-flash'
        elif api_type == 'tongyi':
            models = ['qwen-max', 'qwen-max-longcontext', 'qwen-plus', 'qwen-turbo']
            default = 'qwen-turbo'
    
        self.model_selector.addItems(models)
        self.model_selector.setCurrentText(default)
        self.model_selector.show()
        self.model_selector.currentTextChanged.connect(self.update_generate_button)

    def update_generate_button(self):
        has_content = bool(self.key_display.toPlainText())
        valid_api_key = bool(re.match(r'^sk-[a-zA-Z0-9]{48,}$|^[a-f0-9]{32}\.[a-zA-Z0-9]{16}$|^sk-[a-zA-Z0-9]{32}$', self.api_key_input.currentText()))
        valid_model = self.model_selector.isVisible() and bool(self.model_selector.currentText())
        
        self.generate_button.setEnabled(has_content and valid_api_key and valid_model)

    def load_data(self, file_name):
        self.data_loading_thread = DataLoadingThread(file_name)
        self.data_loading_thread.data_loaded.connect(self.on_data_loaded)
        self.data_loading_thread.start()

    def on_data_loaded(self, data):
        self.data = data
        self.update_app_selector()
        self.display_data(self.data)
        self.update_generate_button()

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
        self.update_generate_button()

    def adjust_column_widths(self):
        header = self.table_widget.horizontalHeader()
        total_width = self.table_widget.viewport().width()
        column_widths = [0] * self.table_widget.columnCount()

        for col in range(self.table_widget.columnCount()):
            header_width = self.table_widget.fontMetrics().width(self.table_widget.horizontalHeaderItem(col).text()) + 20
            column_widths[col] = max(column_widths[col], header_width)
            for row in range(self.table_widget.rowCount()):
                item = self.table_widget.item(row, col)
                if item:
                    column_widths[col] = max(column_widths[col], 
                                            self.table_widget.fontMetrics().width(item.text()) + 20)

        column_widths = [max(width, 50) for width in column_widths]

        total_content_width = sum(column_widths)
        if total_content_width > 0:
            available_width = max(total_width - total_content_width, 0)
            for col in range(self.table_widget.columnCount()):
                extra = int(available_width * (column_widths[col] / total_content_width))
                column_widths[col] += extra

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
            self.export_thread = DataExportThread(self.filtered_data, file_name)
            self.export_thread.export_finished.connect(self.on_export_finished)
            self.export_thread.start()

    def on_export_finished(self, success):
        if success:
            QMessageBox.information(self, "导出成功", "数据已成功导出。")
        else:
            QMessageBox.warning(self, "导出失败", "导出数据时发生错误。")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_name = url.toLocalFile()
            self.load_data(file_name)
        self.clear_explanation()

    def generate_explanation(self):
        content = self.key_display.toPlainText()
        api_key = self.api_key_input.currentText()
        model = self.model_selector.currentText()

        if re.match(r'^sk-[a-zA-Z0-9]{48,}$', api_key):  # OpenAI
            url = "https://api.openai.com/v1/chat/completions"
        elif re.match(r'^[a-f0-9]{32}\.[a-zA-Z0-9]{16}$', api_key):  # 质谱清言
            url = "https://open.bigmodel.cn/api/paas/v4/chat/completions"
        elif re.match(r'^sk-[a-zA-Z0-9]{32}$', api_key):  # 通义千问
            url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
        else:
            QMessageBox.warning(self, "API KEY错误", "请重新输入有效的API KEY")
            return

        self.generation_thread = ContentGenerationThread(url, api_key, model, content)
        self.generation_thread.content_generated.connect(self.update_explanation)
        self.generation_thread.generation_finished.connect(self.on_generation_finished)
        self.generation_thread.start()

        self.explanation_display.clear()
        self.explanation_display.show()
        self.expand_button.show()
        self.stop_button.show()
        self.stop_button.setEnabled(True)
        self.full_explanation = ""
        self.is_expanded = False
        self.update_explanation("正在解析中|")

    def update_explanation(self, content):
        self.full_explanation += content
        if not self.is_expanded:
            self.explanation_display.setText("正在解析中" + "|—/\\"[len(self.full_explanation) % 4])
        else:
            self.explanation_display.setText(self.full_explanation)

        if "```" in self.full_explanation:
            parts = self.full_explanation.split("```")
            if len(parts) > 1:
                self.explanation_display.setText(parts[1])

    def on_generation_finished(self):
        self.copy_button.show()
        self.stop_button.setEnabled(False)
        self.explanation_display.setText(self.full_explanation.split("```")[1] if "```" in self.full_explanation else self.full_explanation)
        self.expand_button.setEnabled(True)

    def toggle_expand(self):
        self.is_expanded = not self.is_expanded
        if self.is_expanded:
            self.expand_button.setText("收起过程")
            self.explanation_display.setText(self.full_explanation)
        else:
            self.expand_button.setText("展开过程")
            parts = self.full_explanation.split("```")
            if len(parts) > 1:
                self.explanation_display.setText(parts[1])
            else:
                self.explanation_display.setText(self.full_explanation)

    def stop_generation(self):
        if self.generation_thread and self.generation_thread.isRunning():
            self.generation_thread.stop()
            self.generation_thread.wait()
        self.on_generation_finished()

    def copy_explanation(self):
        clipboard = QApplication.clipboard()
        clipboard.setText(self.explanation_display.toPlainText())

    def clear_explanation(self):
        self.explanation_display.clear()
        self.explanation_display.hide()
        self.expand_button.hide()
        self.stop_button.hide()
        self.copy_button.hide()
        self.full_explanation = ""
        self.is_expanded = False

class DataLoadingThread(QThread):
    data_loaded = pyqtSignal(list)
    error_occurred = pyqtSignal(str)

    def __init__(self, filename):
        super().__init__()
        self.filename = filename

    def run(self):
        try:
            data = read_compressed_file(self.filename)
            self.data_loaded.emit(data)
        except Exception as e:
            self.error_occurred.emit(str(e))

class DataExportThread(QThread):
    export_finished = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def __init__(self, data, file_name):
        super().__init__()
        self.data = data
        self.file_name = file_name

    def run(self):
        try:
            df = pd.DataFrame(self.data, columns=["应用", "按键", "数量", "时间"])
            if self.file_name.endswith('.csv'):
                df.to_csv(self.file_name, index=False, encoding='utf-8-sig')
            elif self.file_name.endswith('.xlsx'):
                df.to_excel(self.file_name, index=False)
            self.export_finished.emit()
        except Exception as e:
            self.error_occurred.emit(str(e))

def main():
    app = QApplication(sys.argv)
    viewer = DataViewer()
    viewer.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()