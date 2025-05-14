# ui/main_window_ui.py
from PyQt6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QLineEdit, QPushButton, QListWidget, QComboBox, QTreeWidget,
                             QSplitter, QFrame, QSizePolicy, QAbstractItemView, QMessageBox,
                             QFileDialog, QTreeWidgetItem, QCheckBox) # 新增 QCheckBox
from PyQt6.QtCore import Qt

class UiMainWindow(object):
    def setupUi(self, MainWindow: QMainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setWindowTitle("Excel 转 DFQ 工具 (Python版)") # 更新标题
        MainWindow.resize(1100, 750) # 稍微调大一点以便容纳新内容

        self.central_widget = QWidget(MainWindow)
        MainWindow.setCentralWidget(self.central_widget)

        self.main_layout = QHBoxLayout(self.central_widget)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.main_layout.addWidget(self.splitter)

        self.left_pane = QWidget()
        self.left_layout = QVBoxLayout(self.left_pane)
        self.left_pane.setLayout(self.left_layout)

        self.lbl_excel_files = QLabel("1. 导入 Excel 文件:")
        self.left_layout.addWidget(self.lbl_excel_files)
        self.excel_file_list = QListWidget()
        self.excel_file_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.left_layout.addWidget(self.excel_file_list)
        excel_buttons_layout = QHBoxLayout()
        self.btn_add_excel = QPushButton("添加 Excel 文件...")
        excel_buttons_layout.addWidget(self.btn_add_excel)
        self.btn_remove_excel = QPushButton("移除选中")
        excel_buttons_layout.addWidget(self.btn_remove_excel)
        self.btn_clear_excel = QPushButton("全部清除")
        excel_buttons_layout.addWidget(self.btn_clear_excel)
        self.left_layout.addLayout(excel_buttons_layout)

        self.lbl_output_path = QLabel("2. 选择 DFQ 文件输出路径:")
        self.left_layout.addWidget(self.lbl_output_path)
        output_path_layout = QHBoxLayout()
        self.txt_output_path = QLineEdit()
        self.txt_output_path.setPlaceholderText("点击“浏览...”选择输出目录")
        self.txt_output_path.setReadOnly(True)
        output_path_layout.addWidget(self.txt_output_path)
        self.btn_browse_output = QPushButton("浏览...")
        output_path_layout.addWidget(self.btn_browse_output)
        self.left_layout.addLayout(output_path_layout)

        self.lbl_header_info = QLabel("3. 选择抬头信息 (来自系统设置):")
        self.left_layout.addWidget(self.lbl_header_info)
        header_search_layout = QHBoxLayout()
        self.txt_header_search = QLineEdit()
        self.txt_header_search.setPlaceholderText("搜索抬头信息...")
        header_search_layout.addWidget(self.txt_header_search)
        self.btn_header_search_reset = QPushButton("重置搜索")
        header_search_layout.addWidget(self.btn_header_search_reset)
        self.left_layout.addLayout(header_search_layout)
        self.cmb_header_select = QComboBox()
        self.left_layout.addWidget(self.cmb_header_select)
        self.btn_manage_settings = QPushButton("管理系统设置...")
        self.left_layout.addWidget(self.btn_manage_settings)

        self.lbl_actions = QLabel("4. 操作:")
        self.left_layout.addWidget(self.lbl_actions)
        actions_layout = QHBoxLayout()
        self.btn_preview = QPushButton("预览 DFQ 结构")
        actions_layout.addWidget(self.btn_preview)
        self.btn_generate_dfq = QPushButton("生成 DFQ 文件")
        self.btn_generate_dfq.setStyleSheet("font-weight: bold; background-color: #4CAF50; color: white;")
        actions_layout.addWidget(self.btn_generate_dfq)
        self.left_layout.addLayout(actions_layout)
        self.left_layout.addStretch()

        self.right_pane = QWidget()
        self.right_layout = QVBoxLayout(self.right_pane)
        self.right_pane.setLayout(self.right_layout)
        self.lbl_preview_area = QLabel("DFQ 结构预览 (可编辑):") # 更新提示
        self.right_layout.addWidget(self.lbl_preview_area)
        self.tree_preview = QTreeWidget()
        self.tree_preview.setHeaderLabels(["K值字段 / 参数", "值 / 状态"]) # 更新表头
        self.tree_preview.setColumnWidth(0, 350) # 调整列宽
        self.tree_preview.setColumnWidth(1, 200)
        # 允许编辑特定列，具体在 app/main_window.py 中处理哪些项可编辑
        self.tree_preview.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers) # 先禁用默认编辑，后续通过代码控制
        self.right_layout.addWidget(self.tree_preview)

        self.splitter.addWidget(self.left_pane)
        self.splitter.addWidget(self.right_pane)
        self.splitter.setSizes([450, 650])

        self.statusbar = MainWindow.statusBar()
        self.statusbar.showMessage("就绪。")

        self.retranslateUi(MainWindow)

    def retranslateUi(self, MainWindow):
        # 这里可以放置用于Qt Linguist的 tr() 函数，但我们目前是直接修改文本
        pass