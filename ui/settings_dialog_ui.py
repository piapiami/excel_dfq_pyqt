# ui/settings_dialog_ui.py
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                             QPushButton, QTableWidget, QAbstractItemView,
                             QHeaderView, QDialogButtonBox)
from PyQt6.QtCore import Qt

class UiSettingsDialog(object):
    def setupUi(self, SettingsDialog: QDialog):
        SettingsDialog.setObjectName("SettingsDialog")
        SettingsDialog.setWindowTitle("系统设置管理")
        SettingsDialog.setMinimumSize(800, 500) # 调整大小以容纳 K1004
        SettingsDialog.setModal(True)

        self.layout = QVBoxLayout(SettingsDialog)

        search_layout = QHBoxLayout()
        self.lbl_search = QLabel("搜索设置:")
        search_layout.addWidget(self.lbl_search)
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("输入 K1001, K1002, K1086, K1091 或 K1004 的关键词") # 更新提示
        search_layout.addWidget(self.txt_search)
        self.btn_search_reset = QPushButton("重置搜索")
        search_layout.addWidget(self.btn_search_reset)
        self.layout.addLayout(search_layout)

        self.table_settings = QTableWidget()
        self.table_settings.setColumnCount(5) # 增加一列给 K1004
        self.table_settings.setHorizontalHeaderLabels([
            "K1001 (零件号)",
            "K1002 (零件名称)",
            "K1086 (工站)",
            "K1091 (产线)",
            "K1004 (SPC送检数)" # 新增列
        ])
        self.table_settings.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_settings.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_settings.setAlternatingRowColors(True)
        self.layout.addWidget(self.table_settings)

        table_buttons_layout = QHBoxLayout()
        self.btn_add_row = QPushButton("添加新设置")
        table_buttons_layout.addWidget(self.btn_add_row)
        self.btn_delete_row = QPushButton("删除选中设置")
        table_buttons_layout.addWidget(self.btn_delete_row)
        table_buttons_layout.addStretch()
        self.layout.addLayout(table_buttons_layout)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        # 将按钮文本改为中文 (PyQt会根据系统语言自动翻译一些标准按钮，但为了确保，可以手动设置)
        self.button_box.button(QDialogButtonBox.StandardButton.Save).setText("保存")
        self.button_box.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        self.layout.addWidget(self.button_box)

        self.retranslateUi(SettingsDialog)

    def retranslateUi(self, SettingsDialog):
        pass