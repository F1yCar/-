import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QTextEdit, QLabel, QMessageBox, QInputDialog, QStyleFactory, QDialog, 
                             QSlider, QDialogButtonBox, QShortcut, QStatusBar, QComboBox, QCheckBox,
                             QLineEdit, QGridLayout, QProgressDialog, QSplitter, QAction, QTabWidget)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QSettings, QTimer
from PyQt5.QtGui import QCursor, QPalette, QColor, QKeySequence
import asyncio
import pyautogui
import winreg
import logging
from windows_live_captions import check_tesseract, capture_and_process_captions, preprocess_image
import time
import traceback

# 在文件开头添加以下函数
def exception_hook(exctype, value, tb):
    logging.error("Uncaught exception", exc_info=(exctype, value, tb))
    sys.__excepthook__(exctype, value, tb)

sys.excepthook = exception_hook

# 获取当前脚本所在的目录
current_dir = os.path.dirname(os.path.abspath(__file__))
log_file = os.path.join(current_dir, 'app.log')

# 设置日志记录
logging.basicConfig(filename=log_file, level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# 添加一条启动日志
logging.info("程序启动")

class CaptionThread(QThread):
    update_signal = pyqtSignal(str, str)
    error_signal = pyqtSignal(str)

    def __init__(self, x, y, width, height, preprocess_options, ocr_lang, translate_src, translate_dest):
        super().__init__()
        self.x = x
        self.y = y
        self.width = width
        self.height = height
        self.running = True
        self.preprocess_options = preprocess_options
        self.ocr_lang = ocr_lang
        self.translate_src = translate_src
        self.translate_dest = translate_dest
        self.start_time = None

    async def run_async(self):
        self.start_time = time.time()
        last_sentences = []
        last_translation = ""
        while self.running:
            try:
                last_sentences, last_translation, new_original, new_translation = await capture_and_process_captions(
                    self.x, self.y, self.width, self.height, last_sentences, last_translation, 
                    self.preprocess_options, self.ocr_lang, self.translate_src, self.translate_dest
                )
                if new_original or new_translation:
                    self.update_signal.emit(new_original, new_translation)
            except Exception as e:
                self.error_signal.emit(str(e))
                logging.error(f"Error in caption thread: {str(e)}")
                break
            await asyncio.sleep(1)

    def run(self):
        try:
            asyncio.run(self.run_async())
        except Exception as e:
            self.error_signal.emit(f"严重错误: {str(e)}")
            logging.error(f"Severe error in caption thread: {str(e)}")

    def stop(self):
        self.running = False

    @property
    def elapsed_time(self):
        if self.start_time is None:
            return 0
        return time.time() - self.start_time

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        layout = QVBoxLayout(self)

        # 创建选项卡小部件
        tab_widget = QTabWidget()
        layout.addWidget(tab_widget)

        # 常规设置选项卡
        general_tab = QWidget()
        general_layout = QGridLayout(general_tab)
        tab_widget.addTab(general_tab, "常规设置")

        # 字体大小设置
        self.font_size_slider = QSlider(Qt.Horizontal)
        self.font_size_slider.setMinimum(12)
        self.font_size_slider.setMaximum(48)
        self.font_size_slider.setValue(28)
        self.font_size_label = QLabel("字体大小: 28")
        self.font_size_slider.valueChanged.connect(self.update_font_size_label)

        general_layout.addWidget(QLabel("调整字体大小:"), 0, 0)
        general_layout.addWidget(self.font_size_slider, 0, 1)
        general_layout.addWidget(self.font_size_label, 0, 2)

        # 语言设置
        self.source_language = QComboBox()
        self.target_language = QComboBox()
        languages = ["日语", "中文", "英语"]  # 可以根据需要添加更多语言
        self.source_language.addItems(languages)
        self.target_language.addItems(languages)

        general_layout.addWidget(QLabel("识别语言:"), 1, 0)
        general_layout.addWidget(self.source_language, 1, 1)
        general_layout.addWidget(QLabel("翻译目标语言:"), 2, 0)
        general_layout.addWidget(self.target_language, 2, 1)

        # 快捷键设置选项卡
        shortcut_tab = QWidget()
        shortcut_layout = QGridLayout(shortcut_tab)
        tab_widget.addTab(shortcut_tab, "快捷键")

        self.shortcut_select = QLineEdit()
        self.shortcut_start = QLineEdit()
        self.shortcut_stop = QLineEdit()
        self.shortcut_settings = QLineEdit()

        shortcut_layout.addWidget(QLabel("选择区域快捷键:"), 0, 0)
        shortcut_layout.addWidget(self.shortcut_select, 0, 1)
        shortcut_layout.addWidget(QLabel("开始捕获快捷键:"), 1, 0)
        shortcut_layout.addWidget(self.shortcut_start, 1, 1)
        shortcut_layout.addWidget(QLabel("停止捕获快捷键:"), 2, 0)
        shortcut_layout.addWidget(self.shortcut_stop, 2, 1)
        shortcut_layout.addWidget(QLabel("打开设置快捷键:"), 3, 0)
        shortcut_layout.addWidget(self.shortcut_settings, 3, 1)

        # 图像预处理选项卡
        preprocess_tab = QWidget()
        preprocess_layout = QVBoxLayout(preprocess_tab)
        tab_widget.addTab(preprocess_tab, "图像预处理")

        self.preprocess_options = {
            'grayscale': QCheckBox("灰度化"),
            'denoise': QCheckBox("去噪"),
            'threshold': QCheckBox("二值化"),
            'deskew': QCheckBox("倾斜校正")
        }
        
        for checkbox in self.preprocess_options.values():
            preprocess_layout.addWidget(checkbox)

        # 确定和取消按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def update_font_size_label(self, value):
        self.font_size_label.setText(f"字体大小: {value}")

class TranslateThread(QThread):
    finished = pyqtSignal(str, str)
    progress = pyqtSignal(int)

    def __init__(self, text, src, dest):
        super().__init__()
        self.text = text
        self.src = src
        self.dest = dest

    async def translate_with_progress(self):
        from windows_live_captions import translate_text
        max_retries = 3
        for attempt in range(max_retries):
            self.progress.emit((attempt + 1) * 25)  # 更新进度
            try:
                result = await translate_text(self.text, src=self.src, dest=self.dest)
                if result:
                    self.progress.emit(100)  # 翻译成功，进度设为100%
                    return result
            except Exception as e:
                if attempt == max_retries - 1:
                    raise e
                await asyncio.sleep(1)
        return None

    def run(self):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            result = loop.run_until_complete(self.translate_with_progress())
            self.finished.emit(self.text, result if result else "翻译失败")
        except Exception as e:
            self.finished.emit(self.text, f"翻译错误: {str(e)}")
        finally:
            loop.close()

class MainWindow(QMainWindow):
    def __init__(self):
        try:
            super().__init__()
            logging.info("主窗口初始化开始")
            self.setWindowTitle("日语字幕实时翻译")
            self.setGeometry(100, 100, 800, 600)  # 恢复原来的窗口大小

            self.settings = QSettings("YourCompany", "JapaneseSubtitleTranslator")
            self.font_size = self.settings.value("font_size", 28, type=int)
            self.shortcuts = {
                'select': self.settings.value("shortcut_select", "Ctrl+A"),
                'start': self.settings.value("shortcut_start", "Ctrl+S"),
                'stop': self.settings.value("shortcut_stop", "Ctrl+X"),
                'settings': self.settings.value("shortcut_settings", "Ctrl+,")
            }
            self.preprocess_options = {
                'grayscale': self.settings.value("preprocess_grayscale", False, type=bool),
                'denoise': self.settings.value("preprocess_denoise", False, type=bool),
                'threshold': self.settings.value("preprocess_threshold", False, type=bool),
                'deskew': self.settings.value("preprocess_deskew", False, type=bool)
            }

            self.source_language = self.settings.value("source_language", "日语")
            self.target_language = self.settings.value("target_language", "中文")

            # 设置应用程序样式
            self.set_style()

            central_widget = QWidget()
            self.setCentralWidget(central_widget)

            main_layout = QHBoxLayout()
            central_widget.setLayout(main_layout)

            # 实时翻译区域
            self.realtime_widget = QWidget()
            realtime_layout = QVBoxLayout(self.realtime_widget)

            button_layout = QHBoxLayout()
            realtime_layout.addLayout(button_layout)

            self.select_area_button = QPushButton("选择捕获区域")
            self.select_area_button.clicked.connect(self.select_capture_area)
            button_layout.addWidget(self.select_area_button)

            self.start_button = QPushButton("开始捕获")
            self.start_button.clicked.connect(self.start_capture)
            self.start_button.setEnabled(False)
            button_layout.addWidget(self.start_button)

            self.stop_button = QPushButton("停止捕获")
            self.stop_button.clicked.connect(self.stop_capture)
            self.stop_button.setEnabled(False)
            button_layout.addWidget(self.stop_button)

            self.settings_button = QPushButton("设置")
            self.settings_button.clicked.connect(self.open_settings)
            button_layout.addWidget(self.settings_button)

            self.original_text = QTextEdit()
            self.original_text.setReadOnly(True)
            realtime_layout.addWidget(QLabel("实时原文:"))
            realtime_layout.addWidget(self.original_text)

            self.translated_text = QTextEdit()
            self.translated_text.setReadOnly(True)
            realtime_layout.addWidget(QLabel("实时翻译:"))
            realtime_layout.addWidget(self.translated_text)

            main_layout.addWidget(self.realtime_widget)

            # 交互式翻译区域
            self.interactive_widget = QWidget()
            self.interactive_widget.setVisible(False)  # 默认隐藏
            interactive_layout = QVBoxLayout(self.interactive_widget)

            # 使用 QTextEdit 替代 QLineEdit 作为输入框
            self.interactive_input = QTextEdit()
            self.interactive_input.setPlaceholderText("输入要翻译的中文文本")
            self.interactive_input.setMinimumHeight(100)  # 设置最小高度
            interactive_layout.addWidget(self.interactive_input)

            self.interactive_translate_button = QPushButton("翻译成日语")
            self.interactive_translate_button.clicked.connect(self.interactive_translate)
            interactive_layout.addWidget(self.interactive_translate_button)

            self.interactive_original = QTextEdit()
            self.interactive_original.setReadOnly(True)
            interactive_layout.addWidget(QLabel("交互式输入 (中文):"))
            interactive_layout.addWidget(self.interactive_original)

            self.interactive_translated = QTextEdit()
            self.interactive_translated.setReadOnly(True)
            interactive_layout.addWidget(QLabel("交互式翻译 (日语):"))
            interactive_layout.addWidget(self.interactive_translated)

            main_layout.addWidget(self.interactive_widget)

            # 添加显示/隐藏交互式翻译的动作
            self.toggle_interactive_action = QAction("显示交互式翻译", self)
            self.toggle_interactive_action.setCheckable(True)
            self.toggle_interactive_action.triggered.connect(self.toggle_interactive_translation)
            self.toolbar = self.addToolBar("工具栏")
            self.toolbar.addAction(self.toggle_interactive_action)

            # 添加状态栏
            self.statusBar = QStatusBar()
            self.setStatusBar(self.statusBar)
            self.statusBar.showMessage("就绪")

            # 添加   捷键
            self.update_shortcuts()

            # 定时更新状态栏
            self.status_timer = QTimer(self)
            self.status_timer.timeout.connect(self.update_status)
            self.status_timer.start(1000)  # 每秒更新一次

            # 添加以下行来初始化 caption_thread
            self.caption_thread = None
            self.capture_area = None
            self.progress_dialog = None
            self.translate_thread = None

            # 在设置完所有组件后，调整初始窗口大小
            initial_width = 800  # 可以根据需要调整这个值
            initial_height = 600  # 可以根据需要调整这个值
            self.resize(initial_width, initial_height)

            logging.info("主窗口初始化完成")
        except Exception as e:
            logging.error(f"主窗口初始化失败: {str(e)}")
            logging.error(traceback.format_exc())
            QMessageBox.critical(self, "初始化错误", f"主窗口初始化失败:\n{str(e)}\n\n请查看日志文件以获取更多信息。")
            raise

    def set_style(self):
        # 检测系统主题
        is_dark = self.is_windows_dark_mode()
        logging.debug(f"Is Windows in dark mode: {is_dark}")

        if is_dark:
            self.set_dark_theme()
            logging.debug("Applied dark theme")
        else:
            self.set_light_theme()
            logging.debug("Applied light theme")

    def is_windows_dark_mode(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return value == 0
        except Exception as e:
            logging.error(f"Error checking Windows dark mode: {e}")
            return False

    def set_dark_theme(self):
        dark_palette = QPalette()
        dark_palette.setColor(QPalette.Window, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.WindowText, Qt.white)
        dark_palette.setColor(QPalette.Base, QColor(25, 25, 25))
        dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ToolTipBase, Qt.white)
        dark_palette.setColor(QPalette.ToolTipText, Qt.white)
        dark_palette.setColor(QPalette.Text, Qt.white)
        dark_palette.setColor(QPalette.Button, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ButtonText, Qt.white)
        dark_palette.setColor(QPalette.BrightText, Qt.red)
        dark_palette.setColor(QPalette.Link, QColor(42, 130, 218))
        dark_palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
        dark_palette.setColor(QPalette.HighlightedText, Qt.black)
        
        QApplication.setPalette(dark_palette)
        QApplication.setStyle(QStyleFactory.create("Fusion"))
        
        # 更新深色主题的样式表
        self.setStyleSheet(f"""
            QWidget {{
                background-color: #353535;
                color: #ffffff;
            }}
            QPushButton {{
                background-color: #2a2a2a;
                border: 1px solid #5a5a5a;
                padding: 5px;
                font-size: 22px;
            }}
            QPushButton:hover {{
                background-color: #3a3a3a;
            }}
            QPushButton:pressed {{
                background-color: #505050;
            }}
            QTextEdit {{
                background-color: #1a1a1a;
                border: 1px solid #5a5a5a;
                font-size: {self.font_size}px;
            }}
            QLabel {{
                background-color: transparent;
                font-size: 22px;
            }}
            QMessageBox {{
                background-color: #353535;
            }}
            QMessageBox QLabel {{
                color: #ffffff;
                font-size: 22px;
            }}
            QMessageBox QPushButton {{
                background-color: #2a2a2a;
                color: #ffffff;
                border: 1px solid #5a5a5a;
                padding: 5px;
                font-size: 22px;
            }}
        """)

    def set_light_theme(self):
        QApplication.setPalette(QApplication.style().standardPalette())
        QApplication.setStyle(QStyleFactory.create("Fusion"))
        
        # 添加浅色主题的样式表
        self.setStyleSheet(f"""
            QTextEdit {{
                font-size: {self.font_size}px;
            }}
            QLabel {{
                font-size: 22px;
            }}
            QPushButton {{
                font-size: 22px;
            }}
        """)

    def update_shortcuts(self):
        QShortcut(QKeySequence(self.shortcuts['select']), self, self.select_capture_area)
        QShortcut(QKeySequence(self.shortcuts['start']), self, self.start_capture)
        QShortcut(QKeySequence(self.shortcuts['stop']), self, self.stop_capture)
        QShortcut(QKeySequence(self.shortcuts['settings']), self, self.open_settings)

    def select_capture_area(self):
        self.hide()  # 隐藏主窗口
        
        msg = QMessageBox()
        msg.setWindowTitle("选择捕获区域")
        msg.setText("请按照下步骤选择捕区域：\n\n1.   鼠标移动到字幕区域的左上角\n2. 按下确定按钮\n3. 将鼠标移动到字幕域的右下角\n4. 再次按下确定按钮")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

        # 获取左上角坐标
        QApplication.setOverrideCursor(QCursor(Qt.CrossCursor))
        left_top = QMessageBox.information(None, "选择左上角", "请将鼠标移动到字幕区域的左上角，然后点击确定。", QMessageBox.Ok)
        x1, y1 = pyautogui.position()
        QApplication.restoreOverrideCursor()

        # 获取右下角坐标
        QApplication.setOverrideCursor(QCursor(Qt.CrossCursor))
        right_bottom = QMessageBox.information(None, "选择右下角", "请将鼠标移动到字幕区域的右下角，然后点击确定。", QMessageBox.Ok)
        x2, y2 = pyautogui.position()
        QApplication.restoreOverrideCursor()

        self.capture_area = (x1, y1, x2 - x1, y2 - y1)
        self.show()  # 重新显示主窗口

        if self.capture_area:
            self.start_button.setEnabled(True)
            QMessageBox.information(self, "捕获区域已选择", f"捕获区域已成功选择。\n区域: ({x1}, {y1}) - ({x2}, {y2})")

    def start_capture(self):
        logging.info("开始捕获")
        if not self.capture_area:
            QMessageBox.warning(self, "错误", "请先选择捕获区域。")
            return

        try:
            check_tesseract()
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))
            logging.error(f"Tesseract check failed: {str(e)}")
            return

        x, y, width, height = self.capture_area
        self.caption_thread = CaptionThread(x, y, width, height, self.preprocess_options, 
                                            self.source_language, self.translate_src, self.translate_dest)
        self.caption_thread.update_signal.connect(self.update_text)
        self.caption_thread.error_signal.connect(self.show_error)  # 连接错误信号
        self.caption_thread.start()

        self.select_area_button.setEnabled(False)
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.statusBar.showMessage("正在捕获...")

    def stop_capture(self):
        logging.info("停止捕获")
        if self.caption_thread:
            self.caption_thread.stop()
            self.caption_thread.wait()
            self.caption_thread = None

        self.select_area_button.setEnabled(True)
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.statusBar.showMessage("捕获已停止")

    def update_text(self, original, translated):
        if original:
            self.original_text.append(original)
        if translated:
            self.translated_text.append(translated)

    def show_error(self, error_message):
        QMessageBox.critical(self, "错误", error_message)
        logging.error(f"Error occurred: {error_message}")
        self.stop_capture()  # 发生错误时停止捕获

    def closeEvent(self, event):
        self.stop_capture()
        event.accept()

    def open_settings(self):
        dialog = SettingsDialog(self)
        dialog.font_size_slider.setValue(self.font_size)
        dialog.shortcut_select.setText(self.shortcuts['select'])
        dialog.shortcut_start.setText(self.shortcuts['start'])
        dialog.shortcut_stop.setText(self.shortcuts['stop'])
        dialog.shortcut_settings.setText(self.shortcuts['settings'])
        for key, checkbox in dialog.preprocess_options.items():
            checkbox.setChecked(self.preprocess_options[key])

        dialog.source_language.setCurrentText(self.source_language)
        dialog.target_language.setCurrentText(self.target_language)

        if dialog.exec_():
            self.font_size = dialog.font_size_slider.value()
            self.settings.setValue("font_size", self.font_size)
            
            self.shortcuts['select'] = dialog.shortcut_select.text()
            self.shortcuts['start'] = dialog.shortcut_start.text()
            self.shortcuts['stop'] = dialog.shortcut_stop.text()
            self.shortcuts['settings'] = dialog.shortcut_settings.text()
            
            for key, value in self.shortcuts.items():
                self.settings.setValue(f"shortcut_{key}", value)
            
            for key, checkbox in dialog.preprocess_options.items():
                self.preprocess_options[key] = checkbox.isChecked()
                self.settings.setValue(f"preprocess_{key}", checkbox.isChecked())
            
            self.source_language = dialog.source_language.currentText()
            self.target_language = dialog.target_language.currentText()
            self.settings.setValue("source_language", self.source_language)
            self.settings.setValue("target_language", self.target_language)

            self.set_style()
            self.update_shortcuts()

            # 更新翻译设置
            self.update_translation_settings()

    def update_status(self):
        if hasattr(self, 'caption_thread') and self.caption_thread and self.caption_thread.isRunning():
            elapsed_time = self.caption_thread.elapsed_time
            self.statusBar.showMessage(f"正在捕获... 已运行 {elapsed_time:.1f} 秒")
        else:
            self.statusBar.showMessage("就绪")

    # 修改交互式翻译方法
    def interactive_translate(self):
        text = self.interactive_input.toPlainText()  # 使用 toPlainText() 而不是 text()
        if text:
            self.progress_dialog = QProgressDialog("正在翻译...", "取消", 0, 100, self)
            self.progress_dialog.setWindowModality(Qt.WindowModal)
            self.progress_dialog.setAutoClose(True)
            self.progress_dialog.setAutoReset(True)
            
            self.translate_thread = TranslateThread(text, self.translate_src, self.translate_dest)
            self.translate_thread.finished.connect(self.on_translation_finished)
            self.translate_thread.progress.connect(self.update_progress)
            self.translate_thread.start()
            
            self.progress_dialog.exec_()

    def update_progress(self, value):
        if self.progress_dialog and self.progress_dialog.isVisible():
            self.progress_dialog.setValue(value)
        if value == 100:
            self.progress_dialog.close()

    def on_translation_finished(self, original, translated):
        if self.progress_dialog and self.progress_dialog.isVisible():
            self.progress_dialog.close()
        
        if translated != "翻译失败" and not translated.startswith("翻译错误"):
            self.interactive_original.append(original)
            self.interactive_translated.append(translated)
        else:
            QMessageBox.warning(self, "翻译失败", f"无法翻译输入的文本: {translated}")
        
        self.interactive_input.clear()  # 清空输入框

    def toggle_interactive_translation(self):
        is_visible = self.interactive_widget.isVisible()
        current_size = self.size()
        self.interactive_widget.setVisible(not is_visible)
        self.toggle_interactive_action.setText("隐藏交互式翻译" if not is_visible else "显示交互式翻译")
        
        if not is_visible:
            # 如果要显示交互式翻译，增加窗口宽度
            new_width = current_size.width() + 400  # 可以根据需要调整这个值
            self.resize(new_width, current_size.height())
        else:
            # 如果要隐藏交互式翻译，减少窗口宽度
            new_width = current_size.width() - 400  # 与上面的值保持一致
            self.resize(new_width, current_size.height())

    def update_translation_settings(self):
        # 更新翻译设置的逻辑
        # 这里需要根据选择的语言更新 OCR 和翻译的设置
        lang_code = {
            "日语": "ja",
            "中文": "zh-cn",
            "英语": "en"
        }
        self.ocr_lang = lang_code.get(self.source_language, "jpn")
        self.translate_src = lang_code.get(self.source_language, "ja")
        self.translate_dest = lang_code.get(self.target_language, "zh-cn")

if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        logging.critical(f"程序发生严重错误: {str(e)}")
        logging.critical(traceback.format_exc())
        QMessageBox.critical(None, "错误", f"程序发生严重错误:\n{str(e)}\n\n请查看日志文件以获取更多信息。")
