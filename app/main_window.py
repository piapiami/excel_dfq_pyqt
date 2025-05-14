# app/main_window.py
import logging
from PyQt6.QtWidgets import (QMainWindow, QFileDialog, QMessageBox, QTreeWidgetItem, QApplication,
                             QStyledItemDelegate, QLineEdit, QComboBox, QCheckBox, QAbstractItemView,
                             QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QSizePolicy, QListWidgetItem, QTreeWidget, QStyleOptionViewItem)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QModelIndex
from PyQt6.QtGui import QPalette, QColor

from ui.main_window_ui import UiMainWindow
from app.settings_dialog import SettingsDialog
from core import config_manager, excel_processor, dfq_writer
import os
from typing import List, Dict, Any, Tuple

logger = logging.getLogger(__name__)

# K2005 选项
K2005_OPTIONS_MAP = {"0": "次要的", "1": "略重要的", "2": "重要的", "3": "很重要的", "4": "关键的"}
K2005_VALUE_TO_DISPLAY = K2005_OPTIONS_MAP
K2005_DISPLAY_TO_VALUE = {v: k for k, v in K2005_OPTIONS_MAP.items()}

# K2009 选项 - (请确保此字典内容完整)
K2009_OPTIONS = {
    "0": "未定义", "100": "直线度", "101": "平面度", "102": "圆度", "103": "圆柱度",
    "104": "线轮廓度", "105": "面轮廓度", "106": "倾斜度", "107": "垂直度", "108": "平行度",
    "109": "位置度", "110": "同心度", "111": "对称度", "112": "跳动度", "113": "全跳动度",
    "114": "复合一同轴度", "115": "复合一图案位置度", "117": "坐标", "118": "曲面跳动",
    "120": "X坐标", "121": "Y坐标", "122": "Z坐标", "125": "偏移量", "132": "椭圆度",
    "140": "角度区域的评定值", "145": "表面光洁度", "149": "凹坑深度",
    "150": "最大轮廓高度 Rz", "151": "轮廓总高度 Rt", "152": "算术平均偏差 Ra",
    "153": "最大原始轮廓高度 Pt", "154": "轮廓峰高 Rk", "155": "缩减波峰高度",
    "156": "缩减波谷深度", "157": "轮廓波纹深度 Wt", "158": "最大波纹深度 Wz",
    "159": "基本粗糙度深度 Rmax", "160": "材料承载率 Pmr", "161": "材料比例 Mr1",
    "162": "材料比例 Mr2", "170": "油槽深度", "171": "油槽角度", "172": "油槽节距",
    "180": "平均主波纹度", "181": "最大主波纹度", "182": "主波纹长度",
    "190": "粗糙单元平均深度", "191": "轮廓不规则性最大深度", "192": "粗糙单元平均宽度",
    "193": "材料承载率 Rmr", "194": "材料比例 tp", "200": "距离", "201": "半径",
    "202": "直径", "203": "角度", "204": "椭圆短轴", "205": "椭圆长轴", "206": "锥角",
    "207": "内径", "208": "外径", "210": "球面测量杆", "211": "齿高/齿深",
    "212": "参考圆柱上的齿厚", "214": "齿厚偏差（参考圆柱处）", "215": "齿厚变动量",
    "216": "跨（k个）齿公法线长度", "220": "弹簧刚度", "230": "宽度",
    "231": "垂直度（方形度）", "232": "最大直径", "233": "最小直径", "234": "平均直径",
    "250": "温度 [°C]", "251": "温度 [F]", "255": "压力 [bar]", "260": "涂层厚度",
    "270": "体积", "280": "质量", "282": "力", "285": "硬度", "290": "粘度",
    "300": "不平衡量", "301": "扭矩", "302": "拧紧扭矩", "303": "附加扭矩",
    "310": "二维坐标系（注释）", "311": "三维坐标系（注释）", "320": "旋转角度",
    "350": "转速", "360": "角度误差", "362": "轮廓误差", "364": "速度误差",
    "370": "形状偏差", "372": "形状增量", "380": "凸轮高度", "501": "电阻",
    "502": "电容", "503": "电感", "504": "相位移", "505": "频率", "506": "电流强度",
    "507": "电压", "508": "功率", "509": "场强", "601": "节距", "602": "节距误差",
    "604": "累积节距偏差", "605": "累积节距误差", "606": "节距波动",
    "607": "总节距误差", "608": "基节偏差", "609": "轴向节距偏差",
    "610": "齿顶圆直径", "612": "齿根圆直径", "617": "参考圆柱上的槽宽",
    "620": "齿向（线）", "621": "齿向形状误差", "630": "齿廓", "631": "齿廓形状误差",
    "632": "齿廓角度偏差", "633": "齿廓扭曲", "640": "齿顶修缘", "641": "齿廓修鼓",
    "642": "修鼓量", "643": "鼓形高度", "651": "齿线角度偏差", "652": "齿线扭曲",
    "660": "径向跳动偏差", "661": "偏心量", "662": "摆差", "663": "同轴度",
    "670": "双面啮合综合偏差", "671": "双面啮合一齿综合径向偏差",
    "672": "接触跳动偏差", "673": "径向双球（柱）距", "674": "径向双滚柱距",
    "675": "径向单球（柱）距", "676": "径向单滚柱距", "800": "时间", "805": "数量",
    "820": "噪音", "910": "泄漏率", "950": "零件清洁度", "955": "残留粒子"
}
K2009_VALUE_TO_DISPLAY = K2009_OPTIONS
K2009_DISPLAY_TO_VALUE = {v: k for k, v in K2009_OPTIONS.items()}


class ReadOnlyColumnDelegate(QStyledItemDelegate):
    def createEditor(self, parent: QWidget, option: QStyleOptionViewItem, index: QModelIndex) -> QWidget:
        if index.column() == 0:
            return None
        tree_widget = None
        p = parent
        while p is not None:
            if isinstance(p, QTreeWidget):
                tree_widget = p
                break
            p = p.parent()
        if tree_widget:
            item = tree_widget.itemFromIndex(index)
            if item:
                if tree_widget.itemWidget(item, index.column()):
                    return None
                if not (item.flags() & Qt.ItemFlag.ItemIsEditable):
                    return None
        return super().createEditor(parent, option, index)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        logger.info("MainWindow 初始化开始...")
        self.ui = UiMainWindow()
        self.ui.setupUi(self)

        self.ui.central_widget.setObjectName("central_widget")
        self.ui.lbl_preview_area.setText("预览检验计划 (值列可编辑):")

        self.imported_excel_files: List[str] = []
        try:
            self.current_config = config_manager.load_config()
        except Exception as e:
            logger.critical(f"加载初始配置失败: {e}", exc_info=True)
            self.current_config = {"OutputPath": "", "SystemSettings": [], "LastExcelImportPath": ""}
            QMessageBox.critical(self, "配置错误", f"加载配置文件失败: {e}\n程序将使用默认空配置。")

        self.all_header_presets: List[Dict[str, str]] = []
        self.current_parameters_data: List[Dict[str, Any]] = []
        self.current_header_data: Dict[str, str] | None = None

        self.tree_item_delegate = ReadOnlyColumnDelegate(self.ui.tree_preview)
        self.ui.tree_preview.setItemDelegate(self.tree_item_delegate)

        self.ui.tree_preview.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked |
            QAbstractItemView.EditTrigger.SelectedClicked |
            QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self.ui.tree_preview.itemChanged.connect(self.handle_tree_item_changed)
        self.ui.tree_preview.itemClicked.connect(self.handle_tree_item_clicked)
        logger.debug("树编辑触发器和 itemChanged/itemClicked 信号已连接。")

        self.setup_parameter_search_ui()
        self.setup_parameter_reorder_buttons()

        self.load_initial_config()
        self.connect_signals()
        self.update_status("程序已启动，就绪。", duration=5000)
        logger.info("MainWindow 初始化完成。")

    # --- 所有方法的完整实现如下 ---
    def setup_parameter_search_ui(self):
        logger.debug("setup_parameter_search_ui 调用")
        self.search_container_widget = QWidget()
        search_container_layout = QHBoxLayout(self.search_container_widget)
        search_container_layout.setContentsMargins(0, 5, 0, 5)
        lbl_search = QLabel("参数搜索 (K2001/K2002):")
        search_container_layout.addWidget(lbl_search)
        self.txt_param_search = QLineEdit()
        self.txt_param_search.setPlaceholderText("输入关键词...")
        self.txt_param_search.textChanged.connect(self.filter_preview_parameters)
        search_container_layout.addWidget(self.txt_param_search)
        if hasattr(self.ui, 'right_layout') and isinstance(self.ui.right_layout, QVBoxLayout):
            self.ui.right_layout.insertWidget(1, self.search_container_widget)
        else:
            logger.error("UI结构不符合预期：找不到 self.ui.right_layout 或其类型不正确 (用于搜索框)。")

    def filter_preview_parameters(self):
        search_term = self.txt_param_search.text().strip().lower()
        logger.debug(f"filter_preview_parameters: 搜索词 '{search_term}'")
        if not self.ui.tree_preview.topLevelItemCount(): return
        root_item = self.ui.tree_preview.topLevelItem(0)
        if not root_item: return
        params_root_node = None
        for i in range(root_item.childCount()):
            child = root_item.child(i)
            if not child: continue
            data = child.data(0, Qt.ItemDataRole.UserRole)
            if data and data.get("type") == "parameter_group":
                params_root_node = child
                break
        if not params_root_node:
            logger.debug("  未找到参数组节点。")
            return
        for i in range(params_root_node.childCount()):
            param_main_node = params_root_node.child(i)
            if not param_main_node: continue
            item_data = param_main_node.data(0, Qt.ItemDataRole.UserRole)
            if not (item_data and item_data.get("type") == "parameter_main"):
                param_main_node.setHidden(True)
                continue
            param_list_idx = item_data.get("param_list_index")
            if param_list_idx is not None and 0 <= param_list_idx < len(self.current_parameters_data):
                param_dict = self.current_parameters_data[param_list_idx]
                k2001 = param_dict.get("K2001_val", "").lower()
                k2002 = param_dict.get("K2002_val", "").lower()
                if not search_term or search_term in k2001 or search_term in k2002:
                    param_main_node.setHidden(False)
                else:
                    param_main_node.setHidden(True)
            else:
                param_main_node.setHidden(True)
                logger.warning(
                    f"搜索时，UI参数节点 {param_main_node.text(0)} 的 param_list_index ({param_list_idx}) 无效。")
        logger.debug("参数过滤完成。")

    def setup_parameter_reorder_buttons(self):
        logger.debug("setup_parameter_reorder_buttons 调用")
        self.reorder_buttons_widget = QWidget()
        reorder_layout = QHBoxLayout(self.reorder_buttons_widget)
        reorder_layout.setContentsMargins(0, 2, 0, 2)
        reorder_layout.addStretch()
        self.btn_move_param_up = QPushButton("上移选中参数")
        self.btn_move_param_up.clicked.connect(lambda: self.move_selected_parameter_in_tree(-1))
        reorder_layout.addWidget(self.btn_move_param_up)
        self.btn_move_param_down = QPushButton("下移选中参数")
        self.btn_move_param_down.clicked.connect(lambda: self.move_selected_parameter_in_tree(1))
        reorder_layout.addWidget(self.btn_move_param_down)
        if hasattr(self.ui, 'right_layout') and isinstance(self.ui.right_layout, QVBoxLayout):
            self.ui.right_layout.insertWidget(2, self.reorder_buttons_widget)
        else:
            logger.error("UI结构不符合预期：找不到 self.ui.right_layout 或其类型不正确 (用于排序按钮)。")

    def move_selected_parameter_in_tree(self, direction: int):
        logger.info(f"move_selected_parameter_in_tree 调用, direction: {direction}")
        current_tree_item = self.ui.tree_preview.currentItem()
        if not current_tree_item:
            QMessageBox.information(self, "提示", "请先在预览列表中选择一个参数主节点进行移动。")
            return
        item_data = current_tree_item.data(0, Qt.ItemDataRole.UserRole)
        parent_for_check = current_tree_item.parent()
        if not (item_data and item_data.get("type") == "parameter_main"):
            if parent_for_check and parent_for_check.data(0, Qt.ItemDataRole.UserRole) and \
                    parent_for_check.data(0, Qt.ItemDataRole.UserRole).get("type") == "parameter_main":
                current_tree_item = parent_for_check
                item_data = current_tree_item.data(0, Qt.ItemDataRole.UserRole)
            else:
                QMessageBox.information(self, "提示",
                                        "请选择一个参数的主节点 (例如 '参数 X: ...') 或其子项进行移动操作。")
                return
        params_root_node = current_tree_item.parent()
        if not (params_root_node and params_root_node.data(0, Qt.ItemDataRole.UserRole) and \
                params_root_node.data(0, Qt.ItemDataRole.UserRole).get("type") == "parameter_group"):
            logger.error("无法找到参数主节点的父节点 (参数列表组)。")
            return
        param_list_idx = item_data.get("param_list_index")
        if param_list_idx is None or not (0 <= param_list_idx < len(self.current_parameters_data)):
            logger.error(
                f"选中参数的 param_list_index ({param_list_idx}) 无效或越界 (数据长度 {len(self.current_parameters_data)})。")
            QMessageBox.critical(self, "错误", "参数索引不一致，无法移动。请尝试清除搜索词再操作。")
            return
        new_data_idx = param_list_idx + direction
        if not (0 <= new_data_idx < len(self.current_parameters_data)):
            self.update_status(f"参数已在列表{'顶' if direction == -1 else '底'}端。")
            return
        param_to_move_data = self.current_parameters_data.pop(param_list_idx)
        self.current_parameters_data.insert(new_data_idx, param_to_move_data)
        logger.info(
            f"数据模型中: 参数 '{param_to_move_data.get('K2001_val')}' 从索引 {param_list_idx} 移动到 {new_data_idx}。")
        self.populate_preview_tree()
        new_params_root_node = None
        if self.ui.tree_preview.topLevelItemCount() > 0:
            root_item_preview = self.ui.tree_preview.topLevelItem(0)
            if root_item_preview:
                for i_child in range(root_item_preview.childCount()):
                    child = root_item_preview.child(i_child)
                    if not child: continue
                    data = child.data(0, Qt.ItemDataRole.UserRole)
                    if data and data.get("type") == "parameter_group":
                        new_params_root_node = child
                        break
        if new_params_root_node:
            if 0 <= new_data_idx < new_params_root_node.childCount():
                moved_item_in_tree = new_params_root_node.child(new_data_idx)
                if moved_item_in_tree and moved_item_in_tree.data(0, Qt.ItemDataRole.UserRole).get(
                        "param_list_index") == new_data_idx:
                    self.ui.tree_preview.setCurrentItem(moved_item_in_tree)
                    self.ui.tree_preview.scrollToItem(moved_item_in_tree, QAbstractItemView.ScrollHint.PositionAtCenter)
                    logger.debug(f"重新选中并滚动到移动后的参数: {moved_item_in_tree.text(0)}")
        self.update_status(f"参数已{'上移' if direction == -1 else '下移'}。")

    def load_initial_config(self):
        logger.debug("load_initial_config 调用。")
        self.ui.txt_output_path.setText(self.current_config.get("OutputPath", ""))
        if self.current_header_data is None:
            system_settings = config_manager.get_system_settings()
            if system_settings:
                self.all_header_presets = system_settings
                if self.current_header_data is None and self.all_header_presets:
                    self.current_header_data = self.all_header_presets[0].copy()
                    logger.debug(
                        f"load_initial_config: 默认设置 current_header_data 为第一个预设: {self.current_header_data.get('K1001')}")
            else:
                self.all_header_presets = []
                logger.warning("load_initial_config: 系统设置为空，无可用抬头信息。")
        self.refresh_header_combobox()

    def connect_signals(self):
        logger.debug("connect_signals 调用。")
        try:
            self.ui.btn_add_excel.clicked.connect(self.add_excel_files)
            self.ui.btn_remove_excel.clicked.connect(self.remove_selected_excel_files)
            self.ui.btn_clear_excel.clicked.connect(self.clear_all_excel_files)
            self.ui.btn_browse_output.clicked.connect(self.browse_output_path)
            self.ui.btn_manage_settings.clicked.connect(self.open_settings_dialog)
            self.ui.btn_preview.clicked.connect(self.preview_dfq)
            self.ui.btn_generate_dfq.clicked.connect(self.generate_dfq)
            self.ui.txt_header_search.textChanged.connect(self.filter_header_combobox)
            self.ui.btn_header_search_reset.clicked.connect(self.reset_header_search)
            self.ui.cmb_header_select.currentIndexChanged.connect(self.on_header_selection_changed)
            logger.info("UI信号已连接。")
        except Exception as e:
            logger.critical(f"连接UI信号时发生错误: {e}", exc_info=True)
            QMessageBox.critical(self, "严重错误", f"初始化UI事件连接失败: {e}\n程序可能无法正常工作。")

    def on_header_selection_changed(self, index: int):
        logger.info(f"on_header_selection_changed: 索引 {index} 被选中。")
        if index < 0:
            logger.debug("抬头选择无效 (可能列表被清空或索引<0)。")
            return
        new_header_data_candidate = self.ui.cmb_header_select.itemData(index)
        if isinstance(new_header_data_candidate, dict):
            is_different = False
            if self.current_header_data is None:
                is_different = True
            elif len(self.current_header_data) != len(new_header_data_candidate):
                is_different = True
            else:
                for key, value in new_header_data_candidate.items():
                    if self.current_header_data.get(key) != value:
                        is_different = True;
                        break
            if is_different:
                self.current_header_data = new_header_data_candidate.copy()
                logger.info(f"当前抬头信息已更新为K1001: {self.current_header_data.get('K1001')}")
                if self.current_parameters_data or self.ui.tree_preview.topLevelItemCount() > 0:
                    logger.debug("抬头信息已改变，将使用新抬头刷新预览树。")
                    self.populate_preview_tree()
            else:
                logger.debug("选择的抬头与当前抬头相同，不执行更新。")
        else:
            if self.current_header_data is not None:
                self.current_header_data = None
                self.clear_preview_and_data(clear_header=True)
                logger.warning(f"选择了无效的抬头项 (可能是提示信息)，已清空当前抬头和预览。")
            else:
                logger.debug("选择了无效的抬头项，但当前抬头已为空，无需操作。")

    def update_status(self, message: str, is_error: bool = False, duration: int = 4000):
        logger.debug(f"update_status: '{message}', is_error: {is_error}")
        self.ui.statusbar.showMessage(message, duration)
        if is_error: logger.error(f"状态更新 (错误): {message}")

    def refresh_header_combobox(self, settings_list: List[Dict[str, str]] = None):
        logger.debug(f"refresh_header_combobox 调用, settings_list is None: {settings_list is None}")
        header_to_restore = self.current_header_data.copy() if self.current_header_data else None
        self.ui.cmb_header_select.blockSignals(True)
        self.ui.cmb_header_select.clear()
        effective_settings_list = settings_list
        if settings_list is None:
            if not self.all_header_presets:
                self.all_header_presets = config_manager.get_system_settings()
                logger.debug(f"从配置加载了 {len(self.all_header_presets)} 个抬头预设。")
            effective_settings_list = self.all_header_presets
        if not effective_settings_list:
            self.ui.cmb_header_select.addItem("无可用抬头信息，请在“管理系统设置”中配置。")
            self.ui.cmb_header_select.setEnabled(False)
            if self.current_header_data is not None:
                self.current_header_data = None
                logger.info("抬头列表为空，已清空当前选中的抬头信息。")
        else:
            self.ui.cmb_header_select.setEnabled(True)
            restored_idx = -1
            for i, setting in enumerate(effective_settings_list):
                display_text = (f"零件: {setting.get('K1001', '无')} / {setting.get('K1002', '无')} | "
                                f"工站: {setting.get('K1086', '无')} | 产线: {setting.get('K1091', '无')} | "
                                f"SPC数: {setting.get('K1004', '无')}")
                self.ui.cmb_header_select.addItem(display_text, userData=setting.copy())
                if header_to_restore and all(
                        setting.get(k) == header_to_restore.get(k) for k in header_to_restore if k in setting) and \
                        all(header_to_restore.get(k) == setting.get(k) for k in setting if k in header_to_restore):
                    restored_idx = i
            if restored_idx != -1:
                self.ui.cmb_header_select.setCurrentIndex(restored_idx)
                if self.current_header_data != header_to_restore: self.current_header_data = header_to_restore
                logger.debug(f"refresh_header_combobox: 成功恢复之前的抬头选择: {header_to_restore.get('K1001')}")
            elif self.ui.cmb_header_select.count() > 0:
                self.ui.cmb_header_select.setCurrentIndex(0)
                new_default_header = self.ui.cmb_header_select.currentData()
                if isinstance(new_default_header, dict):
                    self.current_header_data = new_default_header.copy()
                    logger.debug(
                        f"refresh_header_combobox: 默认选中第一个抬头, current_header_data 更新为: {self.current_header_data.get('K1001')}")
            else:
                self.current_header_data = None
                logger.info("刷新后抬头列表为空，已清空当前选中的抬头信息。")
        self.ui.cmb_header_select.blockSignals(False)
        if self.current_header_data is None and self.ui.cmb_header_select.currentIndex() >= 0:
            combo_data_final_check = self.ui.cmb_header_select.currentData()
            if isinstance(combo_data_final_check,
                          dict) and "无可用抬头信息" not in self.ui.cmb_header_select.currentText():
                self.current_header_data = combo_data_final_check.copy()
                logger.debug(
                    f"refresh_header_combobox: 后备同步 current_header_data 为: {self.current_header_data.get('K1001')}")
        logger.debug(
            f"refresh_header_combobox 完成。当前抬头K1001: {self.current_header_data.get('K1001') if self.current_header_data else 'None'}")

    def filter_header_combobox(self):
        search_term = self.ui.txt_header_search.text().strip().lower()
        logger.debug(f"filter_header_combobox 调用, 搜索词: '{search_term}'")
        if not self.all_header_presets:
            self.all_header_presets = config_manager.get_system_settings()
        if not search_term:
            self.refresh_header_combobox(self.all_header_presets)
            return
        filtered_presets = [
            s for s in self.all_header_presets
            if (search_term in s.get("K1001", "").lower() or search_term in s.get("K1002", "").lower() or
                search_term in s.get("K1086", "").lower() or search_term in s.get("K1091", "").lower() or
                search_term in s.get("K1004", "").lower())
        ]
        self.refresh_header_combobox(filtered_presets)

    def reset_header_search(self):
        logger.debug("reset_header_search 调用。")
        self.ui.txt_header_search.clear()

    def add_excel_files(self):
        logger.info("add_excel_files: 方法入口。准备打开文件对话框...")
        file_paths = None
        try:
            start_path = self.current_config.get("LastExcelImportPath", "") or os.path.expanduser("~")
            file_paths, _ = QFileDialog.getOpenFileNames(
                self, "选择 Excel 文件", start_path, "Excel 文件 (*.xlsx *.xls)"
            )
            logger.debug(f"QFileDialog.getOpenFileNames 返回的原始 file_paths: {file_paths}")
        except Exception as e:
            logger.critical(f"调用 QFileDialog.getOpenFileNames 时发生异常: {e}", exc_info=True)
            QMessageBox.critical(self, "错误", f"无法打开文件选择对话框: {e}")
            return

        if file_paths and len(file_paths) > 0:
            logger.info(f"用户选择了 {len(file_paths)} 个文件: {file_paths}")
            if file_paths[0] and os.path.exists(file_paths[0]):
                try:
                    last_path = os.path.dirname(file_paths[0])
                    self.current_config["LastExcelImportPath"] = last_path
                    logger.debug(f"下次导入路径将从 '{last_path}' 开始 (当前会话)。")
                except Exception as e_path:
                    logger.warning(f"存储上次导入路径时出错: {e_path}")

            newly_added_count = 0
            for path in file_paths:
                if path not in self.imported_excel_files:
                    self.imported_excel_files.append(path)
                    new_item = QListWidgetItem(os.path.basename(path))
                    new_item.setData(Qt.ItemDataRole.UserRole, path)
                    new_item.setFlags(new_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                    new_item.setCheckState(Qt.CheckState.Checked)
                    self.ui.excel_file_list.addItem(new_item)
                    newly_added_count += 1

            if newly_added_count > 0:
                self.update_status(f"已添加 {newly_added_count} 个 Excel 文件。总计: {len(self.imported_excel_files)}。")
                self.clear_preview_and_data(clear_header=False)
            else:
                self.update_status("选择的文件已在列表中或未选择有效新文件。")
        else:
            self.update_status("未选择文件。")

    def remove_selected_excel_files(self):
        logger.info("remove_selected_excel_files (基于勾选) 调用。")
        rows_to_remove = []
        paths_to_remove_from_model = []
        for i in range(self.ui.excel_file_list.count()):
            item = self.ui.excel_file_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                rows_to_remove.append(self.ui.excel_file_list.row(item))
                paths_to_remove_from_model.append(item.data(Qt.ItemDataRole.UserRole))
        if not rows_to_remove:
            QMessageBox.information(self, "无勾选项", "请勾选列表中要移除的 Excel 文件。")
            return
        for path in paths_to_remove_from_model:
            if path in self.imported_excel_files:
                self.imported_excel_files.remove(path)
        for row in sorted(rows_to_remove, reverse=True):
            taken_item = self.ui.excel_file_list.takeItem(row)
            logger.debug(
                f"已从列表和内部存储中移除 (勾选): {taken_item.data(Qt.ItemDataRole.UserRole) if taken_item else 'N/A'}")
        if rows_to_remove:
            self.update_status(f"已移除 {len(rows_to_remove)} 个勾选的 Excel 文件。")
            self.clear_preview_and_data(clear_header=False)

    def clear_all_excel_files(self):
        logger.info("clear_all_excel_files 调用。")
        if not self.imported_excel_files:
            self.update_status("Excel 文件列表已为空。")
            return
        reply = QMessageBox.question(self, "确认清除", "确定要移除所有导入的 Excel 文件吗？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            logger.info("用户确认清除所有Excel文件。")
            self.imported_excel_files.clear()
            self.ui.excel_file_list.clear()
            self.clear_preview_and_data(clear_header=False)
            self.update_status("已清除列表中的所有 Excel 文件。")

    def clear_preview_and_data(self, clear_header: bool = True):
        logger.debug(f"clear_preview_and_data 调用, clear_header={clear_header}")
        self.ui.tree_preview.clear()
        self.current_parameters_data = []
        if clear_header:
            self.current_header_data = None
            logger.info("预览树、参数数据和当前抬头数据已清除。")
        else:
            logger.info("预览树和当前参数数据已清除 (抬头数据保留)。")
        if hasattr(self, 'txt_param_search'):
            self.txt_param_search.clear()
        self.update_status("预览数据已清除。")

    def browse_output_path(self):
        logger.info("browse_output_path 调用。")
        current_path = self.ui.txt_output_path.text() or self.current_config.get("OutputPath",
                                                                                 "") or os.path.expanduser("~")
        directory = QFileDialog.getExistingDirectory(self, "选择输出目录", current_path)
        if directory:
            self.ui.txt_output_path.setText(directory)
            self.current_config["OutputPath"] = directory
            config_manager.update_output_path(directory)
            logger.info(f"输出路径选定并保存为: {directory}")
            self.update_status(f"输出路径已设置为: {directory}")

    def open_settings_dialog(self):
        logger.info("open_settings_dialog 调用。")
        dialog = SettingsDialog(self)
        if dialog.exec():
            logger.info("设置对话框被接受 (保存)。")
            self.current_config = config_manager.load_config()
            self.all_header_presets = []
            old_header_k1001 = self.current_header_data.get('K1001') if self.current_header_data else None
            self.refresh_header_combobox()
            new_header_k1001 = self.current_header_data.get('K1001') if self.current_header_data else None
            if (old_header_k1001 != new_header_k1001 or not self.ui.tree_preview.topLevelItemCount()) and \
                    (self.current_parameters_data or self.current_header_data):
                logger.info("系统设置更新后，抬头信息可能已改变或预览为空，刷新预览。")
                self.populate_preview_tree()
            self.update_status("系统设置已更新。")
        else:
            logger.info("设置对话框被拒绝 (取消)。")

    def _validate_inputs(self, for_generation: bool = True) -> bool:
        logger.debug(f"_validate_inputs 调用, for_generation: {for_generation}")
        if not self.imported_excel_files:
            QMessageBox.warning(self, "输入错误", "请至少导入一个 Excel 文件。");
            return False
        if for_generation and not self.ui.txt_output_path.text():
            QMessageBox.warning(self, "输入错误", "请为 DFQ 文件选择一个输出路径。");
            return False
        selected_header_text = self.ui.cmb_header_select.currentText()
        if self.ui.cmb_header_select.currentIndex() == -1 or \
                (selected_header_text and "无可用抬头信息" in selected_header_text):
            QMessageBox.warning(self, "输入错误", "请选择有效的抬头信息，或在系统设置中配置。");
            return False
        return True

    def _get_selected_header_info_from_combobox(self) -> Dict[str, str] | None:
        logger.debug("_get_selected_header_info_from_combobox 调用。")
        if self.ui.cmb_header_select.currentIndex() >= 0:
            data = self.ui.cmb_header_select.currentData()
            if isinstance(data, dict) and "无可用抬头信息" not in self.ui.cmb_header_select.currentText():
                return data.copy()
        return None

    def _ensure_data_loaded_for_action(self) -> bool:
        logger.info("_ensure_data_loaded_for_action 调用。")
        action_successful = True
        try:
            if self.current_header_data is None:
                selected_combo_data = self._get_selected_header_info_from_combobox()
                if selected_combo_data:
                    self.current_header_data = selected_combo_data
                else:
                    self._validate_inputs(for_generation=False)
                    if not self.current_header_data:
                        current_data_from_combo = self.ui.cmb_header_select.currentData()
                        if isinstance(current_data_from_combo, dict):
                            self.current_header_data = current_data_from_combo.copy()
                        else:
                            return False
            if not self.current_header_data: return False
            logger.debug(f"EnsureData: 当前抬头 K1001: {self.current_header_data.get('K1001')}")

            if (not self.current_parameters_data and self.imported_excel_files) or \
                    (not self.imported_excel_files and self.current_parameters_data):
                if not self.imported_excel_files and self.current_parameters_data: self.current_parameters_data = []
                if self.imported_excel_files:
                    parameters, errors = excel_processor.read_excel_files(self.imported_excel_files)
                    if errors: QMessageBox.critical(self, "Excel 处理错误",
                                                    "Excel 处理过程中遇到以下错误:\n" + "\n".join(errors))
                    processed_parameters = []
                    if parameters:
                        for p_data in parameters:
                            if 'selected_for_output' not in p_data: p_data['selected_for_output'] = True
                            processed_parameters.append(p_data)
                    self.current_parameters_data = processed_parameters
                    self.update_status(
                        f"已加载 {len(self.current_parameters_data)} 个参数。" if self.current_parameters_data else "未找到参数或处理失败。",
                        is_error=not self.current_parameters_data and bool(errors))
            logger.info(f"_ensure_data_loaded_for_action 完成。参数数量: {len(self.current_parameters_data)}")
        except Exception as e:
            logger.critical(f"_ensure_data_loaded_for_action 发生严重错误: {e}", exc_info=True)
            QMessageBox.critical(self, "内部错误", f"数据准备过程中发生错误: {e}");
            action_successful = False
        return action_successful

    def populate_preview_tree(self):
        logger.info("populate_preview_tree: 开始填充预览树...")
        self.ui.tree_preview.clear()

        if not self.current_header_data:
            logger.warning("populate_preview_tree: 无抬头数据，无法继续。")
            self.update_status("错误：无抬头信息，无法生成预览。", is_error=True)
            return

        self.ui.tree_preview.blockSignals(True)
        logger.debug("populate_preview_tree: 树信号已阻塞。")

        try:
            root_item_text = f"预览检验计划: {self.current_header_data.get('K1001', 'N/A')}"
            root_item = QTreeWidgetItem(self.ui.tree_preview, [root_item_text, ""])
            root_item.setData(0, Qt.ItemDataRole.UserRole, {"type": "root"})
            root_item.setFlags(root_item.flags() & ~Qt.ItemFlag.ItemIsEditable)

            header_node_text = "抬头信息"
            header_node = QTreeWidgetItem(root_item, [header_node_text, ""])
            header_node.setData(0, Qt.ItemDataRole.UserRole, {"type": "header_group"})
            header_node.setFlags(header_node.flags() & ~Qt.ItemFlag.ItemIsEditable)

            selected_param_count = sum(1 for p in self.current_parameters_data if p.get('selected_for_output', True))

            header_fields_meta = [
                ("K0100", "K0100 (参数总数)", False), ("K1001", "K1001 (零件号)", True),
                ("K1002", "K1002 (零件名称)", True), ("K1004", "K1004 (SPC送检数)", True),
                ("K1086", "K1086 (工站)", True), ("K1091", "K1091 (产线)", True),
            ]
            for k_key, default_desc, editable_flag in header_fields_meta:
                val_str = str(selected_param_count) if k_key == "K0100" else self.current_header_data.get(k_key, '')
                item = QTreeWidgetItem(header_node, [default_desc, val_str])
                item.setData(0, Qt.ItemDataRole.UserRole, {"type": "header_k_value", "k_key": k_key})
                current_flags = item.flags()
                if editable_flag:
                    item.setFlags(current_flags | Qt.ItemFlag.ItemIsEditable)
                else:
                    item.setFlags(current_flags & ~Qt.ItemFlag.ItemIsEditable)

            if self.current_parameters_data:
                params_root_node_text = "参数列表"
                params_root_node = QTreeWidgetItem(root_item, [params_root_node_text, ""])
                params_root_node.setData(0, Qt.ItemDataRole.UserRole, {"type": "parameter_group"})
                params_root_node.setFlags(params_root_node.flags() & ~Qt.ItemFlag.ItemIsEditable)

                for i, param_data in enumerate(self.current_parameters_data):
                    param_data['_ui_list_index'] = i
                    param_node_text = f"参数 {i + 1}: {param_data.get('K2001_val', '未命名')}"

                    param_node = QTreeWidgetItem(params_root_node)
                    param_node.setText(0, param_node_text)
                    param_node.setData(0, Qt.ItemDataRole.UserRole, {"type": "parameter_main", "param_list_index": i})
                    param_node.setFlags(
                        param_node.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
                    param_node.setCheckState(0, Qt.CheckState.Checked if param_data.get('selected_for_output',
                                                                                        True) else Qt.CheckState.Unchecked)
                    param_node.setFlags(param_node.flags() & ~Qt.ItemFlag.ItemIsEditable)

                    param_fields_meta = [
                        ("K2001_val", "K2001 (名称)", True, "text"), ("K2002_val", "K2002 (描述)", True, "text"),
                        ("K2003_val", "K2003 (测量频次)", True, "text"),
                        ("K2005_val", "K2005 (参数等级)", True, "combo_k2005"),
                        ("K2009_val", "K2009 (公差类型)", True, "combo_k2009"),
                        ("K2101_val", "K2101 (公称值)", True, "text"),
                        ("K2113_val", "K2113 (上公差)", True, "text"), ("K2112_val", "K2112 (下公差)", True, "text"),
                        ("K2142_val", "K2142 (检验方法)", True, "text"),
                    ]
                    for k_key, default_desc, editable_flag, editor_type in param_fields_meta:
                        display_desc = default_desc
                        if editor_type == "text":
                            item_val_str = str(param_data.get(k_key, ''))
                            item = QTreeWidgetItem(param_node, [display_desc, item_val_str])
                            item.setData(0, Qt.ItemDataRole.UserRole, {
                                "type": "parameter_k_value", "param_list_index": i, "k_key": k_key,
                                "editor_type": "text"
                            })
                            if editable_flag:
                                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                            else:
                                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                        elif editor_type == "combo_k2005" or editor_type == "combo_k2009":
                            item = QTreeWidgetItem(param_node, [display_desc, ""])
                            item.setData(0, Qt.ItemDataRole.UserRole, {
                                "type": f"parameter_{editor_type}_container", "param_list_index": i, "k_key": k_key,
                                "editor_type": editor_type
                            })
                            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                            combo = QComboBox(self.ui.tree_preview)
                            options_map = K2005_OPTIONS_MAP if editor_type == "combo_k2005" else K2009_OPTIONS
                            current_value_from_model_str = str(param_data.get(k_key, "0"))  # 使用修正的变量名

                            found_in_options = False
                            for val_idx_str_opt, text_option in options_map.items():
                                combo.addItem(text_option, userData=val_idx_str_opt)
                                if val_idx_str_opt == current_value_from_model_str:  # 使用修正的变量名
                                    combo.setCurrentText(text_option)
                                    found_in_options = True

                            if not found_in_options and combo.count() > 0:
                                combo.setCurrentIndex(0)
                                actual_first_val = combo.itemData(0)
                                if param_data.get(k_key) != actual_first_val:
                                    param_data[k_key] = actual_first_val
                                    logger.warning(
                                        f"参数{i}的{k_key}值'{current_value_from_model_str}'不在选项中,已设为默认'{actual_first_val}' ({combo.itemText(0)})")  # 使用修正的变量名
                            elif combo.count() == 0:
                                logger.warning(f"Combobox for {k_key} (param {i}) has no options!")
                            combo.setProperty("param_list_index", i)
                            combo.setProperty("k_key", k_key)
                            combo.currentIndexChanged.connect(self.handle_parameter_combobox_changed_via_sender)
                            self.ui.tree_preview.setItemWidget(item, 1, combo)

                    natural_limits_meta = [
                        ("K2121_val", "K2121 (上自然界限)", "upper", "K2113_val"),
                        ("K2120_val", "K2120 (下自然界限)", "lower", "K2112_val"),
                    ]
                    for nl_k_key, default_desc, limit_type, tol_k_key in natural_limits_meta:
                        display_desc = default_desc
                        nl_item_container = QTreeWidgetItem(param_node, [display_desc, ""])
                        nl_item_container.setData(0, Qt.ItemDataRole.UserRole, {
                            "type": "natural_limit_checkbox_container", "param_list_index": i,
                            "nl_k_key": nl_k_key, "limit_type": limit_type, "tol_k_key": tol_k_key
                        })
                        nl_item_container.setFlags(nl_item_container.flags() & ~Qt.ItemFlag.ItemIsEditable)

                        cb = QCheckBox(self.ui.tree_preview)
                        current_nl_value_from_model = param_data.get(nl_k_key, '0')
                        corresponding_tol_value_in_data = param_data.get(tol_k_key, '')

                        is_tol_truly_empty = (corresponding_tol_value_in_data.strip() == "")
                        cb.setEnabled(not is_tol_truly_empty)

                        if not is_tol_truly_empty:
                            cb.setChecked(current_nl_value_from_model == '1' or current_nl_value_from_model == '2')
                        else:
                            cb.setChecked(False)
                            if current_nl_value_from_model != '0':
                                param_data[nl_k_key] = '0'
                                logger.debug(
                                    f"  populate_preview_tree: 参数 {i} 的 {nl_k_key} 因公差为空，模型值强制为 '0'")

                        cb.setProperty("param_list_index", i)
                        cb.setProperty("nl_k_key", nl_k_key)
                        cb.setProperty("tol_k_key", tol_k_key)
                        cb.stateChanged.connect(self.handle_natural_limit_checkbox_changed_via_sender)
                        self.ui.tree_preview.setItemWidget(nl_item_container, 1, cb)
                        logger.debug(
                            f"      populate_preview_tree: NL {nl_k_key} for param {i}, tol='{corresponding_tol_value_in_data}', "
                            f"enabled={cb.isEnabled()}, checked={cb.isChecked()}, model_NL_val='{param_data[nl_k_key]}'"
                        )

            self.ui.tree_preview.expandAll()
            self.update_status("DFQ 结构预览已填充/更新。")
        except Exception as e:
            logger.critical(f"populate_preview_tree 执行期间发生错误: {e}", exc_info=True)
            QMessageBox.critical(self, "预览错误", f"生成预览树时发生错误: {e}")
        finally:
            self.ui.tree_preview.blockSignals(False)
            logger.debug("populate_preview_tree: 树信号已恢复。")
            logger.info("populate_preview_tree: 预览树填充/更新完成。")

    def handle_parameter_combobox_changed_via_sender(self, combo_idx_ignore: int):
        sender_combo = self.sender()
        if not isinstance(sender_combo, QComboBox): return
        param_list_index = sender_combo.property("param_list_index")
        k_value_key = sender_combo.property("k_key")
        logger.debug(
            f"handle_parameter_combobox_changed_via_sender: param_idx={param_list_index}, k_key={k_value_key}, new_combo_idx={sender_combo.currentIndex()}")
        if not (param_list_index is not None and 0 <= param_list_index < len(self.current_parameters_data)):
            logger.warning("  无效的参数索引。");
            return
        selected_value_idx_str = sender_combo.currentData()
        if selected_value_idx_str is None:
            logger.warning(f"  Combobox for {k_value_key} (param {param_list_index}) currentData 为 None。");
            return
        param_data = self.current_parameters_data[param_list_index]
        if param_data.get(k_value_key) != selected_value_idx_str:
            param_data[k_value_key] = selected_value_idx_str
            display_text = sender_combo.currentText()
            update_message = f"参数 {param_list_index + 1} 的 {k_value_key.replace('_val', '')} 更新为 '{selected_value_idx_str}' ({display_text})"
            logger.info(update_message)
            self.update_status(update_message)
        else:
            logger.debug("  Combobox 值未改变。")

    def handle_natural_limit_checkbox_changed_via_sender(self, state: int):
        sender_checkbox = self.sender()
        if not isinstance(sender_checkbox, QCheckBox): return

        param_list_index = sender_checkbox.property("param_list_index")
        natural_limit_key = sender_checkbox.property("nl_k_key")
        tolerance_key = sender_checkbox.property("tol_k_key")

        logger.info(
            f"handle_natural_limit_checkbox_changed: 参数索引 {param_list_index}, NL键 {natural_limit_key}, 新UI状态 {state}")
        if not (param_list_index is not None and 0 <= param_list_index < len(self.current_parameters_data)):
            logger.warning(f"  参数索引 {param_list_index} 超出范围。");
            return

        param_data = self.current_parameters_data[param_list_index]

        sender_checkbox.blockSignals(True)
        try:
            current_tolerance_value_in_data = param_data.get(tolerance_key, '')
            is_tol_truly_empty_in_data = (current_tolerance_value_in_data.strip() == "")

            expected_enabled_state = not is_tol_truly_empty_in_data
            if sender_checkbox.isEnabled() != expected_enabled_state:
                sender_checkbox.setEnabled(expected_enabled_state)
                logger.debug(
                    f"  复选框 {natural_limit_key} enabled 状态根据公差'{current_tolerance_value_in_data}'校正为: {expected_enabled_state}")

            if expected_enabled_state:  # 如果复选框可用 (即公差有输入)
                if state == Qt.CheckState.Checked.value:
                    param_data[natural_limit_key] = '2'
                else:  # Unchecked
                    param_data[natural_limit_key] = '1'
            else:  # 复选框不可用 (公差为空)
                param_data[natural_limit_key] = '0'
                if sender_checkbox.isChecked(): sender_checkbox.setChecked(False)

            self.update_status(
                f"参数 {param_list_index + 1} 的 {natural_limit_key} 更新为 '{param_data[natural_limit_key]}'")
            logger.debug(
                f"  参数 {param_list_index} 的 {natural_limit_key} 更新后值为 '{param_data[natural_limit_key]}', 复选框UI状态: checked={sender_checkbox.isChecked()}, enabled={sender_checkbox.isEnabled()}")
        except Exception as e:
            logger.error(f"handle_natural_limit_checkbox_changed_via_sender 内部错误: {e}", exc_info=True)
        finally:
            sender_checkbox.blockSignals(False)

    def handle_tree_item_clicked(self, item: QTreeWidgetItem, column: int):
        logger.debug(
            f"handle_tree_item_clicked: 项 '{item.text(0)}', 列 {column} 被点击。复选状态(列0): {item.checkState(0)}")
        item_data = item.data(0, Qt.ItemDataRole.UserRole)
        if item_data and item_data.get("type") == "parameter_main":
            param_list_idx = item_data.get("param_list_index")
            if param_list_idx is not None and 0 <= param_list_idx < len(self.current_parameters_data):
                is_selected_new_ui_state = (item.checkState(0) == Qt.CheckState.Checked)
                current_model_state = self.current_parameters_data[param_list_idx].get('selected_for_output', True)
                if current_model_state != is_selected_new_ui_state:
                    self.current_parameters_data[param_list_idx]['selected_for_output'] = is_selected_new_ui_state
                    self.update_status(
                        f"参数 {param_list_idx + 1} 输出状态: {'选中' if is_selected_new_ui_state else '未选中'}")
                    self.update_k0100_in_tree()

    def update_k0100_in_tree(self):
        logger.debug("update_k0100_in_tree 调用。")
        if not self.ui.tree_preview.topLevelItemCount(): return
        root_item = self.ui.tree_preview.topLevelItem(0)
        if not (root_item and root_item.childCount() > 0): return
        header_node = root_item.child(0)
        if not (header_node and header_node.data(0, Qt.ItemDataRole.UserRole) and \
                header_node.data(0, Qt.ItemDataRole.UserRole).get("type") == "header_group"): return
        try:
            for i in range(header_node.childCount()):
                k_item = header_node.child(i)
                if not k_item: continue
                item_data = k_item.data(0, Qt.ItemDataRole.UserRole)
                if item_data and item_data.get("k_key") == "K0100":
                    selected_param_count = sum(
                        1 for p in self.current_parameters_data if p.get('selected_for_output', True))
                    if k_item.text(1) != str(selected_param_count):
                        k_item.setText(1, str(selected_param_count))
                    break
        except Exception as e:
            logger.error(f"update_k0100_in_tree 发生错误: {e}", exc_info=True)

    def handle_tree_item_changed(self, item: QTreeWidgetItem, column: int):
        logger.debug(f"handle_tree_item_changed: 项 '{item.text(0)}', 列 {column} 值更改。")
        if column != 1: return
        self.ui.tree_preview.blockSignals(True)
        try:
            item_data = item.data(0, Qt.ItemDataRole.UserRole)
            if not (item_data and isinstance(item_data, dict)):
                self.ui.tree_preview.blockSignals(False);
                return
            item_type = item_data.get("type")
            new_value_str = item.text(1)
            update_message = ""
            if item_type == "header_k_value":
                k_key = item_data.get("k_key")
                if k_key == "K0100":
                    selected_param_count = sum(
                        1 for p in self.current_parameters_data if p.get('selected_for_output', True))
                    item.setText(1, str(selected_param_count));
                    self.ui.tree_preview.blockSignals(False);
                    return
                if self.current_header_data and k_key in self.current_header_data:
                    if self.current_header_data[k_key] != new_value_str:
                        self.current_header_data[k_key] = new_value_str
                        update_message = f"抬头 {k_key} 更新为 {new_value_str}"
                        if k_key == "K1001" and self.ui.tree_preview.topLevelItemCount() > 0:
                            self.ui.tree_preview.topLevelItem(0).setText(0,
                                                                         f"预览检验计划: {new_value_str}")  # 修正2 UI文本
            elif item_type == "parameter_k_value":
                p_idx = item_data.get("param_list_index")
                k_key = item_data.get("k_key")
                editor_type = item_data.get("editor_type")
                if editor_type != "text": self.ui.tree_preview.blockSignals(False); return
                if p_idx is not None and k_key and 0 <= p_idx < len(self.current_parameters_data):
                    param_dict = self.current_parameters_data[p_idx]
                    if param_dict.get(k_key) != new_value_str:
                        param_dict[k_key] = new_value_str
                        update_message = f"参数 {p_idx + 1} 的 {k_key.replace('_val', '')} 更新为 {new_value_str}"
                        if k_key == "K2001_val" and item.parent(): item.parent().setText(0,
                                                                                         f"参数 {p_idx + 1}: {new_value_str}")
                        if (k_key == "K2113_val" or k_key == "K2112_val") and item.parent():
                            self.refresh_natural_limit_checkbox_state_for_param(item.parent(), p_idx)
            if update_message: self.update_status(update_message); logger.info(f"数据更新: {update_message}")
        except Exception as e:
            logger.critical(f"handle_tree_item_changed error: {e}", exc_info=True)
        finally:
            self.ui.tree_preview.blockSignals(False)

    def refresh_natural_limit_checkbox_state_for_param(self, param_main_node: QTreeWidgetItem, param_list_index: int):
        logger.info(f"refresh_natural_limit_checkbox_state_for_param: 参数索引 {param_list_index}")
        if not (param_main_node and 0 <= param_list_index < len(self.current_parameters_data)):
            logger.warning("  refresh_natural_limit_checkbox_state_for_param: 无效的参数节点或索引，跳过刷新。")
            return

        param_data = self.current_parameters_data[param_list_index]
        logger.debug(
            f"  刷新自然界限状态前，NL相关值: K2113='{param_data.get('K2113_val')}', K2112='{param_data.get('K2112_val')}', K2121='{param_data.get('K2121_val')}', K2120='{param_data.get('K2120_val')}'")

        try:
            for i in range(param_main_node.childCount()):
                child_item = param_main_node.child(i)
                if not child_item: continue
                child_item_data = child_item.data(0, Qt.ItemDataRole.UserRole)

                if child_item_data and child_item_data.get("type") == "natural_limit_checkbox_container":
                    cb_widget = self.ui.tree_preview.itemWidget(child_item, 1)
                    if not isinstance(cb_widget, QCheckBox): continue

                    cb_widget.blockSignals(True)

                    tol_k_key = child_item_data.get("tol_k_key")
                    nl_k_key = child_item_data.get("nl_k_key")
                    current_tol_value_in_data = param_data.get(tol_k_key, "")

                    is_tol_truly_empty = (current_tol_value_in_data.strip() == "")
                    is_enabled_new_ui = not is_tol_truly_empty

                    if cb_widget.isEnabled() != is_enabled_new_ui:
                        cb_widget.setEnabled(is_enabled_new_ui)

                    current_nl_value_from_model = param_data.get(nl_k_key, '0')
                    new_nl_value_for_model = current_nl_value_from_model
                    ui_should_be_checked = False

                    if not is_enabled_new_ui:  # 如果公差为空 (复选框应禁用)
                        new_nl_value_for_model = '0'
                        ui_should_be_checked = False
                    else:  # 公差存在 (复选框应启用)
                        # 如果模型值是'0' (例如，之前公差为空导致NL=0，现在公差刚被填上)，则应变为'1'
                        if new_nl_value_for_model == '0':
                            new_nl_value_for_model = '1'
                        # UI的勾选状态应反映模型是 '1' 或 '2'
                        ui_should_be_checked = (new_nl_value_for_model == '1' or new_nl_value_for_model == '2')

                    if param_data.get(nl_k_key) != new_nl_value_for_model:
                        param_data[nl_k_key] = new_nl_value_for_model
                        logger.debug(
                            f"    参数 {param_list_index} 的 {nl_k_key} 数据模型值因公差 {tol_k_key}='{current_tol_value_in_data}' 而更新/设置为 '{new_nl_value_for_model}'")

                    if cb_widget.isChecked() != ui_should_be_checked:
                        cb_widget.setChecked(ui_should_be_checked)
                        logger.debug(
                            f"    参数 {param_list_index} 的 {nl_k_key} 复选框UI checked 状态更新为: {ui_should_be_checked}")

                    cb_widget.blockSignals(False)
            logger.debug(
                f"  刷新后参数数据 (NL相关): K2113='{param_data.get('K2113_val')}', K2112='{param_data.get('K2112_val')}', K2121='{param_data.get('K2121_val')}', K2120='{param_data.get('K2120_val')}'")
        except Exception as e:
            logger.error(f"refresh_natural_limit_checkbox_state_for_param 内部错误: {e}", exc_info=True)

    def preview_dfq(self):
        logger.info("preview_dfq: 开始预览操作。")
        try:
            if not self._validate_inputs(for_generation=False):
                logger.warning("预览输入验证失败。")
                return
            if not self._ensure_data_loaded_for_action():
                logger.warning("预览数据加载或准备失败。")
                self.populate_preview_tree()
                return

            if self.current_parameters_data or self.current_header_data:
                logger.debug("有数据，开始填充预览树。")
                self.populate_preview_tree()
                total_width = self.ui.splitter.width()
                if total_width > 100:
                    self.ui.splitter.setSizes([int(total_width * 0.4), int(total_width * 0.6)])
                else:
                    self.ui.splitter.setSizes([480, 720])
            else:
                self.ui.tree_preview.clear()
                logger.info("无数据可预览。")
                self.update_status("未找到可预览的数据。", is_error=True)
        except Exception as e:
            logger.critical(f"preview_dfq 执行期间发生错误: {e}", exc_info=True)
            QMessageBox.critical(self, "预览错误", f"生成预览时发生未知错误: {e}")

    def generate_dfq(self):
        logger.info("generate_dfq: 开始生成DFQ文件操作。")
        try:
            if not self._validate_inputs(for_generation=True): return
            if not self._ensure_data_loaded_for_action():
                self.update_status("DFQ 生成取消，数据准备失败。", is_error=True);
                return
            if not self.current_header_data:
                QMessageBox.warning(self, "抬头信息缺失", "没有有效的抬头信息，无法生成DFQ文件。");
                return

            parameters_to_output = [p for p in self.current_parameters_data if p.get('selected_for_output', True)]
            if not parameters_to_output and self.imported_excel_files:
                QMessageBox.information(self, "无参数选中", "没有参数被选中输出，无法生成DFQ文件。");
                return
            elif not parameters_to_output and not self.imported_excel_files:
                QMessageBox.information(self, "无参数数据", "请先导入并处理Excel文件。");
                return

            dfq_content = dfq_writer.generate_dfq_content(parameters_to_output, self.current_header_data)
            output_dir = self.ui.txt_output_path.text()
            success, message_or_filepath = dfq_writer.write_dfq_file(output_dir, dfq_content,
                                                                     self.current_header_data)

            if success:
                QMessageBox.information(self, "成功", f"DFQ文件已成功生成:\n{message_or_filepath}")
                self.update_status(f"DFQ文件已生成: {os.path.basename(message_or_filepath)}")
            else:
                QMessageBox.critical(self, "生成DFQ错误", f"生成DFQ文件失败:\n{message_or_filepath}")
                self.update_status(f"DFQ生成失败: {message_or_filepath}", is_error=True)
        except Exception as e:
            logger.critical(f"generate_dfq 执行期间发生严重错误: {e}", exc_info=True)
            QMessageBox.critical(self, "生成错误", f"生成DFQ文件时发生未知错误: {e}")

    def closeEvent(self, event):
        logger.info("closeEvent: 应用程序正在关闭...")
        try:
            config_to_save = config_manager.load_config()
            config_to_save["LastExcelImportPath"] = self.current_config.get("LastExcelImportPath", "")
            config_to_save["OutputPath"] = self.current_config.get("OutputPath", "")
            config_manager.save_config(config_to_save)
            logger.info("应用程序关闭前已保存部分配置。")
        except Exception as e:
            logger.error(f"关闭时保存配置出错: {e}", exc_info=True)
        event.accept()
        logger.info("应用程序已接受关闭事件。")