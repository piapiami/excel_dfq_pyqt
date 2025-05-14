# app/settings_dialog.py
from PyQt6.QtWidgets import QDialog, QMessageBox, QTableWidgetItem
from PyQt6.QtCore import Qt
from ui.settings_dialog_ui import UiSettingsDialog
from core.config_manager import get_system_settings, update_system_settings
from typing import List, Dict


class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = UiSettingsDialog()
        self.ui.setupUi(self)

        self.all_settings: List[Dict[str, str]] = []
        self.load_initial_settings()

        self.ui.btn_add_row.clicked.connect(self.add_row)
        self.ui.btn_delete_row.clicked.connect(self.delete_row)
        self.ui.txt_search.textChanged.connect(self.filter_settings)
        self.ui.btn_search_reset.clicked.connect(self.reset_search_and_filter)
        self.ui.button_box.accepted.connect(self.save_settings)
        self.ui.button_box.rejected.connect(self.reject)

    def load_initial_settings(self):
        # 从 config_manager 获取时，确保K1004已按需填充默认值
        self.all_settings = get_system_settings()
        self.populate_table(self.all_settings)

    def populate_table(self, settings_list: List[Dict[str, str]]):
        self.ui.table_settings.setRowCount(0)
        for setting in settings_list:
            row_position = self.ui.table_settings.rowCount()
            self.ui.table_settings.insertRow(row_position)
            self.ui.table_settings.setItem(row_position, 0, QTableWidgetItem(setting.get("K1001", "")))
            self.ui.table_settings.setItem(row_position, 1, QTableWidgetItem(setting.get("K1002", "")))
            self.ui.table_settings.setItem(row_position, 2, QTableWidgetItem(setting.get("K1086", "")))
            self.ui.table_settings.setItem(row_position, 3, QTableWidgetItem(setting.get("K1091", "")))
            self.ui.table_settings.setItem(row_position, 4, QTableWidgetItem(setting.get("K1004", "5")))  # K1004

    def add_row(self):
        row_position = self.ui.table_settings.rowCount()
        self.ui.table_settings.insertRow(row_position)
        for col in range(self.ui.table_settings.columnCount()):
            self.ui.table_settings.setItem(row_position, col, QTableWidgetItem(""))
        # 为K1004设置默认值
        self.ui.table_settings.setItem(row_position, 4, QTableWidgetItem("5"))
        self.ui.table_settings.scrollToBottom()
        self.ui.table_settings.selectRow(row_position)
        self.ui.table_settings.editItem(self.ui.table_settings.item(row_position, 0))

    def delete_row(self):
        selected_rows = self.ui.table_settings.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.information(self, "无选择", "请选择要删除的行。")  # 中文
            return

        reply = QMessageBox.question(self, "确认删除",  # 中文
                                     f"确定要删除选中的 {len(selected_rows)} 项设置吗?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            for index in sorted([r.row() for r in selected_rows], reverse=True):
                self.ui.table_settings.removeRow(index)

    def filter_settings(self):
        search_term = self.ui.txt_search.text().strip().lower()
        if not search_term:
            self.populate_table(self.all_settings)
            return

        filtered = []
        for setting in self.all_settings:
            if (search_term in setting.get("K1001", "").lower() or
                    search_term in setting.get("K1002", "").lower() or
                    search_term in setting.get("K1086", "").lower() or
                    search_term in setting.get("K1091", "").lower() or
                    search_term in setting.get("K1004", "").lower()):  # 加入K1004搜索
                filtered.append(setting)
        self.populate_table(filtered)

    def reset_search_and_filter(self):
        self.ui.txt_search.clear()

    def save_settings(self):
        current_table_settings = []
        for row in range(self.ui.table_settings.rowCount()):
            k1001 = self.ui.table_settings.item(row, 0).text().strip() if self.ui.table_settings.item(row, 0) else ""
            k1002 = self.ui.table_settings.item(row, 1).text().strip() if self.ui.table_settings.item(row, 1) else ""
            k1086 = self.ui.table_settings.item(row, 2).text().strip() if self.ui.table_settings.item(row, 2) else ""
            k1091 = self.ui.table_settings.item(row, 3).text().strip() if self.ui.table_settings.item(row, 3) else ""
            k1004 = self.ui.table_settings.item(row, 4).text().strip() if self.ui.table_settings.item(row,
                                                                                                      4) else "5"  # K1004，空则给默认

            if k1001 or k1002 or k1086 or k1091:  # K1004可以独立存在，所以不作为保存行的判断依据
                current_table_settings.append({
                    "K1001": k1001, "K1002": k1002, "K1086": k1086, "K1091": k1091, "K1004": k1004
                })

        update_system_settings(current_table_settings)
        self.all_settings = current_table_settings
        QMessageBox.information(self, "设置已保存", "系统设置已成功更新。")  # 中文
        self.accept()