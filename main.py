# main.py
import sys
import logging
import os
from PyQt6.QtWidgets import QApplication, QStyleFactory
from app.main_window import MainWindow

LOG_FILENAME = 'app_trace.log'

GLOBAL_STYLESHEET = """
    QWidget { 
        font: 9pt "Segoe UI", "Microsoft YaHei UI", sans-serif;
        color: black; 
        background-color: white; 
    }
    QMainWindow, QDialog {
        background-color: white;
    }
    QPushButton { 
        background-color: white; 
        color: black;
        border: 1px solid #bababa;
        padding: 5px 10px;
        border-radius: 3px;
        min-height: 20px;
    }
    QPushButton:hover {
        background-color: #f0f0f0; 
    }
    QPushButton:pressed {
        background-color: #e0e0e0; 
    }
    QPushButton:disabled {
        background-color: #f0f0f0;
        color: #a0a0a0;
    }
    QLineEdit, QComboBox, QListWidget, QTreeWidget, QTextEdit, QPlainTextEdit {
        background-color: white;
        color: black;
        border: 1px solid #bababa;
    }
    QListWidget::item:selected, QTreeWidget::item:selected {
        background-color: #cce8ff; 
        color: black; 
    }
    QHeaderView::section {
        background-color: #f0f0f0; 
        color: black;
        padding: 4px;
        border: 1px solid #d0d0d0;
        font-weight: bold;
    }
    QCheckBox { 
        background-color: transparent;
    }
    QCheckBox::indicator {
        border: 1px solid #555;
        width: 13px;
        height: 13px;
        border-radius: 2px;
        background-color: white;
    }
    QCheckBox::indicator:checked {
        /* Default system checkmark will be used */
    }
    QCheckBox::indicator:disabled {
        background-color: #e0e0e0;
        border-color: #c0c0c0;
    }
    QComboBox::drop-down {
        border-left: 1px solid #c0c0c0;
    }
    QStatusBar {
        background-color: #f0f0f0;
    }

    QPushButton#btn_generate_dfq { 
        font-weight: bold; 
        background-color: #4CAF50; 
        color: white;
        border: 1px solid #388E3C;
    }
    QPushButton#btn_generate_dfq:hover {
        background-color: #45a049;
    }
    QPushButton#btn_generate_dfq:pressed {
        background-color: #3e8e41;
    }
"""

def setup_logging():
    if os.path.exists(LOG_FILENAME):
        try: os.remove(LOG_FILENAME)
        except Exception as e: print(f"警告：无法删除旧的日志文件 {LOG_FILENAME}: {e}")
    log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(funcName)s:%(lineno)d - %(message)s')
    file_handler = logging.FileHandler(LOG_FILENAME, encoding='utf-8')
    file_handler.setFormatter(log_formatter)
    file_handler.setLevel(logging.DEBUG)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(log_formatter)
    console_handler.setLevel(logging.INFO)
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)
    if not root_logger.hasHandlers():
        root_logger.addHandler(file_handler)
        root_logger.addHandler(console_handler)
    logging.info("日志系统已启动。")
    sys.excepthook = handle_unhandled_exception

def handle_unhandled_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logging.critical("未捕获的异常:", exc_info=(exc_type, exc_value, exc_traceback))

def main():
    setup_logging()
    logging.info("应用程序启动...")
    app = QApplication(sys.argv)
    app.setStyleSheet(GLOBAL_STYLESHEET)
    main_window = MainWindow()
    logging.info("主窗口已创建。")
    main_window.show()
    logging.info("主窗口已显示。")
    app.exec()
    logging.info(f"应用程序已退出。")

if __name__ == "__main__":
    main()