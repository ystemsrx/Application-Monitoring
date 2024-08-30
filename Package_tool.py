import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, 
    QListWidget, QTextEdit, QHBoxLayout, QMessageBox, QAbstractItemView, QProgressDialog
)
from PyQt5.QtCore import Qt, QProcess, QThread, pyqtSignal, QTimer

class WorkerThread(QThread):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()
    cancel_signal = pyqtSignal()

    def __init__(self, command):
        super().__init__()
        self.command = command
        self._process = None

    def run(self):
        self._process = QProcess()
        self._process.setProcessChannelMode(QProcess.MergedChannels)
        self._process.start(self.command)
        self._process.waitForFinished()

        if self._process.state() != QProcess.NotRunning:
            self._process.kill()

        output = self._process.readAll().data().decode()
        self.log_signal.emit(output)
        self.finished_signal.emit()

    def cancel(self):
        if self._process and self._process.state() == QProcess.Running:
            self._process.kill()
            self.log_signal.emit("打包任务已取消")
            self.finished_signal.emit()
            self.cancel_signal.emit()

class PyInstallerGUI(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("打包工具")
        self.setGeometry(300, 300, 600, 400)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QAbstractItemView.NoSelection)  # 禁止选择
        self.layout.addWidget(self.file_list)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.layout.addWidget(self.log_output)

        self.button_layout = QHBoxLayout()

        self.package_button = QPushButton("开始打包")
        self.package_button.clicked.connect(self.start_packaging)
        self.button_layout.addWidget(self.package_button)

        self.clear_button = QPushButton("清空列表")
        self.clear_button.clicked.connect(self.clear_list)
        self.button_layout.addWidget(self.clear_button)

        self.layout.addLayout(self.button_layout)

        self.file_list.setDragEnabled(True)
        self.file_list.viewport().setAcceptDrops(True)
        self.file_list.setDropIndicatorShown(True)
        self.file_list.setDragDropMode(QListWidget.DragDrop)

        # 用于存放用户拖入的图标文件路径
        self.additional_icon_file = None

        # 存储线程的引用，防止线程被垃圾回收
        self.threads = []

        # 使窗口接受拖放
        self.setAcceptDrops(True)

        # 用于跟踪打包任务的完成情况
        self.tasks_in_progress = 0

        # 进度对话框
        self.progress_dialog = None

        # 用于控制进度条滑动的计时器
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_progress)
        self.progress_value = 0

    def start_packaging(self):
        if self.file_list.count() == 0:
            QMessageBox.warning(self, "警告", "没有文件可以打包！")
            return

        self.tasks_in_progress = self.file_list.count()
        self.progress_dialog = QProgressDialog("打包进行中...", "取消", 0, 100, self)
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.canceled.connect(self.cancel_packaging)
        self.progress_dialog.show()

        # 设置初始值并启动计时器
        self.progress_value = 0
        self.timer.start(150)  # 每150ms更新一次进度

        for index in range(self.file_list.count()):
            file_path = self.file_list.item(index).text()
            if self.additional_icon_file:
                command = f"pyinstaller --onefile --noconsole --icon={self.additional_icon_file} {file_path}"
            else:
                command = f"pyinstaller --onefile --noconsole {file_path}"

            self.log_output.append(f"正在执行: {command}")

            # 创建并启动新线程来执行打包命令
            thread = WorkerThread(command)
            thread.log_signal.connect(self.update_log)
            thread.finished_signal.connect(self.on_task_finished)
            thread.cancel_signal.connect(self.on_task_canceled)
            thread.start()

            # 存储线程引用
            self.threads.append(thread)

    def update_progress(self):
        # 每次更新将进度条值增加
        if self.progress_value < 85:
            self.progress_value += 1
            self.progress_dialog.setValue(self.progress_value)

    def update_log(self, log):
        self.log_output.append(log)

    def on_task_finished(self):
        self.tasks_in_progress -= 1
        if self.tasks_in_progress == 0:
            # 停止计时器并将进度条滑动到终点
            self.timer.stop()
            self.progress_dialog.setValue(100)
            self.progress_dialog.close()
            QMessageBox.information(self, "完成", "所有打包任务已完成！")

    def on_task_canceled(self):
        self.timer.stop()
        self.progress_dialog.close()
        QMessageBox.information(self, "取消", "打包任务已取消！")

    def cancel_packaging(self):
        for thread in self.threads:
            thread.cancel()

    def clear_list(self):
        self.file_list.clear()
        self.log_output.clear()
        self.additional_icon_file = None

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            for url in urls:
                file_path = url.toLocalFile()
                if self.file_list.count() < 5:
                    if file_path.endswith('.py') and not self.is_duplicate(file_path):
                        self.file_list.addItem(file_path)
                    elif file_path.endswith(('.png', '.jpg', '.ico', '.svg')):
                        if self.additional_icon_file:
                            QMessageBox.warning(self, "图标冲突", "已经有图标文件了，请删除后再拖入新的图标文件。")
                        else:
                            self.additional_icon_file = file_path
                            self.log_output.append(f"使用用户提供的图标文件: {file_path}")
                    else:
                        QMessageBox.warning(self, "文件类型错误", "只能拖入.py文件或图标文件")
                else:
                    QMessageBox.warning(self, "数量限制", "最多只能拖入5个文件")
                    break

    def is_duplicate(self, file_path):
        for i in range(self.file_list.count()):
            if self.file_list.item(i).text() == file_path:
                return True
        return False

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PyInstallerGUI()
    window.show()
    sys.exit(app.exec_())
