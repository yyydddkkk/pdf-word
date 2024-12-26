import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, QFileDialog, QVBoxLayout, QWidget, QProgressBar, QHBoxLayout)
from PyQt5.QtCore import Qt, QMimeData, QThread, pyqtSignal, QObject
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
from pdf2docx import Converter as PDFConverter
from docx2pdf import convert as docx2pdf

# 添加转换工作线程类
class ConversionThread(QThread):
    progress = pyqtSignal(int, str)  # 进度信号(百分比, 文件名)
    status = pyqtSignal(str)    # 状态信号
    finished = pyqtSignal(bool) # 完成信号
    error = pyqtSignal(str, str)  # 错误信号(错误信息, 文件名)

    def __init__(self, input_path, output_path, conversion_type):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.conversion_type = conversion_type

    def run(self):
        try:
            filename = os.path.basename(self.input_path)
            self.status.emit(f'正在转换: {filename}')
            self.progress.emit(10, filename)

            if self.conversion_type == 'pdf2word':
                cv = PDFConverter(self.input_path)
                self.progress.emit(30, filename)
                cv.convert(self.output_path, start=0, end=None)
                cv.close()
                
                # 清理临时文件
                output_dir = os.path.dirname(self.output_path)
                temp_file = os.path.join(output_dir, f"~${os.path.basename(self.output_path)}")
                if os.path.exists(temp_file):
                    try:
                        os.remove(temp_file)
                    except:
                        pass
                        
            else:  # word2pdf
                self.progress.emit(30, filename)
                docx2pdf(self.input_path, self.output_path)

            self.progress.emit(100, filename)
            self.finished.emit(True)

        except Exception as e:
            self.error.emit(str(e), filename)
            self.finished.emit(False)

# 添加批量转换管理器
class BatchConversionManager(QObject):
    all_completed = pyqtSignal()  # 所有转换完成的信号
    file_completed = pyqtSignal(str, bool)  # 单个文件完成的信号(文件名, 是否成功)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.threads = []
        self.active_threads = 0
        self.max_threads = 3  # 最大同时转换数量

    def add_conversion(self, input_path, output_path, conversion_type):
        thread = ConversionThread(input_path, output_path, conversion_type)
        thread.finished.connect(lambda success: self.thread_finished(thread, success))
        self.threads.append(thread)
        self.start_next_thread()

    def start_next_thread(self):
        if self.active_threads < self.max_threads:
            for thread in self.threads:
                if not thread.isRunning():
                    self.active_threads += 1
                    thread.start()
                    break

    def thread_finished(self, thread, success):
        self.active_threads -= 1
        self.threads.remove(thread)
        self.file_completed.emit(os.path.basename(thread.input_path), success)
        
        if self.threads:
            self.start_next_thread()
        elif self.active_threads == 0:
            self.all_completed.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        # 存储所有打开的子窗口
        self.child_windows = []
        
    def initUI(self):
        # 设置窗口基本属性
        self.setWindowTitle('文档格式转换工具')
        self.setGeometry(300, 300, 400, 250)
        self.setFixedSize(400, 250)
        
        # 创建中心部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setAlignment(Qt.AlignCenter)
        
        # 标题标签
        title_label = QLabel('文档格式转换工具')
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet('''
            QLabel {
                font-size: 24px;
                font-weight: bold;
                color: #333;
                margin-bottom: 20px;
            }
        ''')
        
        # PDF转Word按钮
        self.pdf_to_word_btn = QPushButton('PDF转Word')
        self.pdf_to_word_btn.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 15px;
                border-radius: 8px;
                border: none;
                min-width: 200px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        ''')
        self.pdf_to_word_btn.clicked.connect(self.open_pdf_to_word)
        
        # Word转PDF按钮
        self.word_to_pdf_btn = QPushButton('Word转PDF')
        self.word_to_pdf_btn.setStyleSheet('''
            QPushButton {
                background-color: #008CBA;
                color: white;
                padding: 15px;
                border-radius: 8px;
                border: none;
                min-width: 200px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #007B9A;
            }
        ''')
        self.word_to_pdf_btn.clicked.connect(self.open_word_to_pdf)
        
        # 添加部件到布局
        layout.addWidget(title_label)
        layout.addSpacing(20)
        layout.addWidget(self.pdf_to_word_btn)
        layout.addSpacing(15)
        layout.addWidget(self.word_to_pdf_btn)
        
    def open_pdf_to_word(self):
        self.pdf_converter = PDFToWordWindow()
        self.pdf_converter.show()
        # 将窗口添加到子窗口列表
        self.child_windows.append(self.pdf_converter)
        
    def open_word_to_pdf(self):
        self.word_converter = WordToPDFWindow()
        self.word_converter.show()
        # 将窗口添加到子窗口列表
        self.child_windows.append(self.word_converter)
    
    def closeEvent(self, event):
        """重写关闭事件"""
        # 关闭所有子窗口
        for window in self.child_windows:
            window.close()
        event.accept()

class BaseConverterWindow(QMainWindow):
    def __init__(self, title, input_format, output_format):
        super().__init__()
        self.input_format = input_format
        self.output_format = output_format
        self.input_path = ''
        self.output_path = ''
        self.initUI(title)
        
        # 启用拖放
        self.setAcceptDrops(True)
        
        # 添加转换线程属性
        self.conversion_thread = None
        self.batch_manager = BatchConversionManager(self)
        self.batch_manager.file_completed.connect(self.file_completed)
        self.batch_manager.all_completed.connect(self.all_completed)
        self.files_status = {}  # 记录每个文件的转换状态
        self.is_batch_mode = False  # 添加批量模式标志
        self.batch_files = []       # 存储批量文件列表

    def initUI(self, title):
        self.setWindowTitle(title)
        # 增加窗口大小
        self.setGeometry(350, 350, 600, 400)
        self.setFixedSize(600, 400)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(15)  # 增加控件之间的间距
        layout.setContentsMargins(20, 20, 20, 20)  # 增加边距
        
        # 文件显示标签
        self.file_label = QLabel(f'拖拽{self.input_format}文件到这里\n或点击选择按钮')
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.setWordWrap(True)  # 允许文字换行
        self.file_label.setStyleSheet('''
            QLabel {
                background-color: #f0f0f0;
                padding: 20px;
                border-radius: 5px;
                border: 2px dashed #aaa;
                min-height: 80px;
                font-size: 14px;
            }
            QLabel:hover {
                background-color: #e8e8e8;
                border-color: #999;
            }
        ''')
        
        # 按钮容器
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        button_layout.setSpacing(20)  # 按钮之间的间距
        
        # 选择文件按钮
        self.select_btn = QPushButton(f'选择{self.input_format}文件')
        self.select_btn.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 12px 20px;
                border-radius: 5px;
                border: none;
                min-width: 200px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        ''')
        self.select_btn.clicked.connect(self.select_file)
        
        # 转换按钮
        self.convert_btn = QPushButton('开始转换')
        self.convert_btn.setStyleSheet('''
            QPushButton {
                background-color: #008CBA;
                color: white;
                padding: 12px 20px;
                border-radius: 5px;
                border: none;
                min-width: 200px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #007B9A;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        ''')
        self.convert_btn.clicked.connect(self.convert_file)
        self.convert_btn.setEnabled(False)
        
        # 添加按钮到按钮容器
        button_layout.addWidget(self.select_btn)
        button_layout.addWidget(self.convert_btn)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet('''
            QProgressBar {
                border: 2px solid #ddd;
                border-radius: 5px;
                text-align: center;
                height: 25px;
                font-size: 12px;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
            }
        ''')
        self.progress_bar.hide()
        
        # 状态标签
        self.status_label = QLabel('')
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet('''
            QLabel {
                color: #666;
                font-size: 14px;
                min-height: 30px;
            }
        ''')
        
        # 快捷键提示
        shortcut_label = QLabel('快捷键：Ctrl+V 粘贴文件路径')
        shortcut_label.setAlignment(Qt.AlignCenter)
        shortcut_label.setStyleSheet('''
            QLabel {
                color: #666;
                font-size: 12px;
                margin-top: 5px;
            }
        ''')
        
        # 添加所有部件到主布局
        layout.addWidget(self.file_label)
        layout.addWidget(button_container)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.status_label)
        layout.addWidget(shortcut_label)

    def dragEnterEvent(self, event: QDragEnterEvent):
        """处理拖入文件事件"""
        if event.mimeData().hasUrls():
            valid_files = [url.toLocalFile() for url in event.mimeData().urls()
                         if url.toLocalFile().lower().endswith(f'.{self.input_format.lower()}')]
            if valid_files:
                event.acceptProposedAction()
                self.file_label.setStyleSheet('''
                    QLabel {
                        background-color: #e8f5e9;
                        padding: 20px;
                        border-radius: 5px;
                        border: 2px dashed #4CAF50;
                        min-height: 80px;
                        font-size: 14px;
                    }
                ''')

    def dragLeaveEvent(self, event):
        """处理拖放离开事件"""
        self.file_label.setStyleSheet('''
            QLabel {
                background-color: #f0f0f0;
                padding: 15px;
                border-radius: 5px;
                border: 2px dashed #aaa;
                min-height: 60px;
            }
        ''')
    
    def dropEvent(self, event: QDropEvent):
        """处理文件拖���事件"""
        files = [url.toLocalFile() for url in event.mimeData().urls()
                if url.toLocalFile().lower().endswith(f'.{self.input_format.lower()}')]
        
        if len(files) > 1:
            self.is_batch_mode = True
            self.batch_files = files
            self.update_file_label()
        else:
            self.is_batch_mode = False
            self.process_file_path(files[0])
            
        self.file_label.setStyleSheet('''
            QLabel {
                background-color: #f0f0f0;
                padding: 20px;
                border-radius: 5px;
                border: 2px dashed #aaa;
                min-height: 80px;
                font-size: 14px;
            }
        ''')
    
    def keyPressEvent(self, event):
        """处理键盘事件"""
        if event.matches(Qt.KeySequence.Paste):
            clipboard = QApplication.clipboard()
            mime_data = clipboard.mimeData()
            
            if mime_data.hasText():
                file_path = mime_data.text().strip()
                # 移除可能的引号
                if file_path.startswith('"') and file_path.endswith('"'):
                    file_path = file_path[1:-1]
                self.process_file_path(file_path)
    
    def process_file_path(self, file_path):
        """处理文件路径"""
        if os.path.isfile(file_path) and file_path.lower().endswith(f'.{self.input_format.lower()}'):
            if self.is_batch_mode:
                # 批量模式下添加文件到列表
                if file_path not in self.batch_files:
                    self.batch_files.append(file_path)
                    self.update_file_label()
            else:
                # 单文件模式
                self.input_path = file_path
                self.file_label.setText(os.path.basename(file_path))
                self.convert_btn.setEnabled(True)
            self.status_label.setText('')
        else:
            self.status_label.setText(f'请选择正确的{self.input_format}文件')
            self.status_label.setStyleSheet('color: #f44336;')

    def update_file_label(self):
        """更新文件显示标签"""
        if self.is_batch_mode:
            if self.batch_files:
                files_text = '\n'.join([os.path.basename(f) for f in self.batch_files[:3]])
                if len(self.batch_files) > 3:
                    files_text += f'\n... 等{len(self.batch_files)}个文件'
                self.file_label.setText(files_text)
                self.convert_btn.setText(f'开始转换 ({len(self.batch_files)}个文件)')
                self.convert_btn.setEnabled(True)
            else:
                self.file_label.setText(f'拖拽{self.input_format}文件到这里\n或点击选择按钮')
                self.convert_btn.setText('开始转换')
                self.convert_btn.setEnabled(False)
        else:
            if self.input_path:
                self.convert_btn.setText('开始转换')
            else:
                self.file_label.setText(f'拖拽{self.input_format}文件到这里\n或点击选择按钮')
                self.convert_btn.setText('开始转换')

    def select_file(self):
        """文件选择方法"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            f"选择{self.input_format}文件",
            "",
            f"{self.input_format}文件 (*.{self.input_format.lower()})"
        )
        
        if files:
            if len(files) > 1:
                self.is_batch_mode = True
                self.batch_files = files
                self.update_file_label()
            else:
                self.is_batch_mode = False
                self.process_file_path(files[0])

    def convert_file(self):
        """统一的转换处理方法"""
        if not (self.input_path or self.batch_files):
            return
            
        if self.is_batch_mode:
            # 批量转换模式
            output_dir = QFileDialog.getExistingDirectory(
                self,
                "选择保存目录",
                ""
            )
            
            if not output_dir:
                return
                
            # 禁用按钮
            self.select_btn.setEnabled(False)
            self.convert_btn.setEnabled(False)
            
            # 显示进度条
            self.progress_bar.setRange(0, len(self.batch_files))
            self.progress_bar.setValue(0)
            self.progress_bar.show()
            
            # 开始批量转换
            self.files_status.clear()
            for input_file in self.batch_files:
                filename = os.path.basename(input_file)
                output_file = os.path.join(
                    output_dir,
                    os.path.splitext(filename)[0] + f".{self.output_format.lower()}"
                )
                self.files_status[filename] = 'pending'
                self.batch_manager.add_conversion(input_file, output_file, 
                    'pdf2word' if isinstance(self, PDFToWordWindow) else 'word2pdf'
                )
            
            self.update_status(f'正在批量转换 {len(self.batch_files)} 个文件...')
        else:
            # 单文件转换模式
            save_name, _ = QFileDialog.getSaveFileName(
                self,
                f"保存{self.output_format}文件",
                os.path.splitext(self.input_path)[0] + f".{self.output_format.lower()}",
                f"{self.output_format}文件 (*.{self.output_format.lower()})"
            )
            
            if save_name:
                self.output_path = save_name
                self.start_conversion()

    def start_conversion(self):
        try:
            # 禁用按钮
            self.select_btn.setEnabled(False)
            self.convert_btn.setEnabled(False)
            
            # 显示和重置进度条
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)
            self.progress_bar.show()
            
            # 创建并启动转换线程
            self.conversion_thread = ConversionThread(
                self.input_path, 
                self.output_path,
                'pdf2word' if isinstance(self, PDFToWordWindow) else 'word2pdf'
            )
            
            # 连接信号
            self.conversion_thread.progress.connect(self.update_progress)
            self.conversion_thread.status.connect(self.update_status)
            self.conversion_thread.finished.connect(self.conversion_finished)
            self.conversion_thread.error.connect(self.conversion_error)
            
            # 启动线程
            self.conversion_thread.start()
            
        except Exception as e:
            self.conversion_error(str(e))

    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)

    def update_status(self, message):
        """更新状态信息"""
        self.status_label.setText(message)
        self.status_label.setStyleSheet('color: #666;')

    def conversion_finished(self, success):
        """处理转换完成"""
        if success:
            self.progress_bar.setValue(100)
            self.status_label.setText('转换完成！')
            self.status_label.setStyleSheet('color: #4CAF50;')
        
        # 重新启用按钮
        self.select_btn.setEnabled(True)
        self.convert_btn.setEnabled(True)

    def conversion_error(self, error_message):
        """处理转换错误"""
        self.progress_bar.hide()
        self.status_label.setText(f'转换失败：{error_message}')
        self.status_label.setStyleSheet('color: #f44336;')
        
        # 重新启用按钮
        self.select_btn.setEnabled(True)
        self.convert_btn.setEnabled(True)

    def file_completed(self, filename, success):
        """单个文件转换完成的处理"""
        self.files_status[filename] = 'success' if success else 'failed'
        completed = sum(1 for status in self.files_status.values() if status != 'pending')
        self.progress_bar.setValue(completed)
        
        # 更新状态信息
        success_count = sum(1 for status in self.files_status.values() if status == 'success')
        failed_count = sum(1 for status in self.files_status.values() if status == 'failed')
        self.update_status(f'已完成: {success_count} 成功, {failed_count} 失败')
        
    def all_completed(self):
        """所有文件转换完成的处理"""
        success_count = sum(1 for status in self.files_status.values() if status == 'success')
        failed_count = sum(1 for status in self.files_status.values() if status == 'failed')
        
        self.status_label.setText(f'批量转换完成！成功: {success_count}, 失败: {failed_count}')
        self.status_label.setStyleSheet('color: #4CAF50;' if failed_count == 0 else 'color: #FF9800;')
        
        # 重新启用按钮
        self.select_btn.setEnabled(True)
        self.convert_btn.setEnabled(True)
        # 重置批量模式
        self.is_batch_mode = False
        self.batch_files = []
        self.update_file_label()

class PDFToWordWindow(BaseConverterWindow):
    def __init__(self):
        super().__init__('PDF转Word工具', 'PDF', 'DOCX')

class WordToPDFWindow(BaseConverterWindow):
    def __init__(self):
        super().__init__('Word转PDF工具', 'DOCX', 'PDF')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    main_window = MainWindow()
    main_window.show()
    
    sys.exit(app.exec_()) 