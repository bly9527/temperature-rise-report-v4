# -*- coding: utf-8 -*-
import os
import sys
import json
import subprocess
import shutil
import uuid
import PyQt5
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QFileDialog,
    QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox,
    QProgressBar, QTextEdit, QGroupBox, QMessageBox, QSplitter,
    QInputDialog, QListWidget, QListWidgetItem
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QLibraryInfo
from PyQt5.QtGui import QFont

from thermal_processor import process_bulk_thermal_data, delete_condition_result

def configure_qt_plugins():
    # Some Qt5 Windows wheels lose non-ASCII characters in their built-in path.
    # Use Python's Unicode package path when that default directory is invalid.
    if not os.path.isdir(QLibraryInfo.location(QLibraryInfo.PluginsPath)):
        plugin_dir = os.path.join(os.path.dirname(PyQt5.__file__), 'Qt5', 'plugins')
        if os.path.isdir(plugin_dir):
            QApplication.addLibraryPath(plugin_dir)

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.abspath(os.path.dirname(__file__))

QSS_STYLE = """
QWidget {
    font-family: "Segoe UI", "Microsoft YaHei", "PingFang SC", sans-serif;
    font-size: 13px;
    color: #2c3e50;
    background-color: #f8f9fa;
}
QGroupBox {
    font-weight: bold;
    font-size: 14px;
    border: 1px solid #dcdfe6;
    border-radius: 8px;
    margin-top: 10px;
    padding-top: 15px;
    background-color: #ffffff;
}
QGroupBox::title { color: #409eff; }
QPushButton {
    background-color: #409eff;
    color: white;
    border: none;
    border-radius: 4px;
    padding: 6px 15px;
    font-weight: bold;
}
QPushButton:hover { background-color: #66b1ff; }
QPushButton#btn_danger { background-color: #f56c6c; }
QPushButton#btn_success { background-color: #67c23a; }
QLineEdit { border: 1px solid #dcdfe6; border-radius: 4px; padding: 5px; }
QTableWidget, QListWidget { border: 1px solid #e4e7ed; background-color: white; }
QHeaderView::section { background-color: #f5f7fa; font-weight: bold; padding: 5px; }
"""

class WorkerThread(QThread):
    progress_updated = pyqtSignal(int)
    log_updated = pyqtSignal(str)
    finished_signal = pyqtSignal()
    error_signal = pyqtSignal(str)

    def __init__(self, project_excel, condition_configs, fixed_images, channels_config,
                 project_config=None):
        super().__init__()
        self.project_excel = project_excel
        self.condition_configs = condition_configs
        self.fixed_images = fixed_images
        self.channels_config = channels_config
        self.project_config = project_config

    def run(self):
        try:
            process_bulk_thermal_data(
                self.project_excel,
                self.condition_configs,
                self.fixed_images,
                self.channels_config,
                self.progress_callback,
                self.log_callback,
                self.project_config
            )
            self.finished_signal.emit()
        except Exception as e:
            self.error_signal.emit(str(e).strip() or '报告生成失败，请检查输入数据后重试。')

    def progress_callback(self, val):
        self.progress_updated.emit(val)
    def log_callback(self, msg):
        self.log_updated.emit(msg)


class DeleteConditionThread(QThread):
    log_updated = pyqtSignal(str)
    succeeded = pyqtSignal()
    failed = pyqtSignal(str)

    def __init__(self, report, condition, config, config_path):
        super().__init__()
        self.report = report
        self.condition = condition
        self.config = config
        self.config_path = config_path

    def run(self):
        try:
            delete_condition_result(self.report, self.condition, self.config,
                                    self.config_path, self.log_updated.emit)
            self.succeeded.emit()
        except Exception as error:
            self.failed.emit(str(error).strip() or '删除失败，请检查报告文件后重试。')


class ThermalApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("温升报告自动生成系统 v4.0 - Excel原生折线图（分批更新修复）")
        self.resize(1300, 900)
        self.setStyleSheet(QSS_STYLE)

        self.base_dir = get_base_dir()
        self.cfg_dir = os.path.join(self.base_dir, "configs")
        os.makedirs(self.cfg_dir, exist_ok=True)
        self.assets_dir = os.path.join(self.base_dir, "assets")
        os.makedirs(self.assets_dir, exist_ok=True)
        resource_dir = getattr(sys, '_MEIPASS', self.base_dir)
        self.base_template = os.path.join(resource_dir, "assets", "底层模板.xlsx")

        self.current_project_excel = ""
        self.pending_image_deletions = []
        self.deleted_image_records = []
        self.init_ui()
        self.load_project_list()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 1. 工程管理
        group_proj = QGroupBox("📌 工程管理")
        lyt_proj = QHBoxLayout(group_proj)
        self.cbo_projects = QComboBox()
        self.cbo_projects.setMinimumWidth(250)
        self.cbo_projects.currentIndexChanged.connect(self.on_project_changed)
        lyt_proj.addWidget(QLabel("当前工程存档:"))
        lyt_proj.addWidget(self.cbo_projects)

        btn_new_proj = QPushButton("➕ 新建工程")
        btn_new_proj.clicked.connect(self.create_new_project)
        lyt_proj.addWidget(btn_new_proj)

        btn_del_proj = QPushButton("✖ 删除工程")
        btn_del_proj.setStyleSheet("background-color: #ff4d4f; color: white;")
        btn_del_proj.clicked.connect(self.delete_project)
        lyt_proj.addWidget(btn_del_proj)

        btn_save_proj = QPushButton("💾 保存当前配置")
        btn_save_proj.clicked.connect(self.save_project_config)
        lyt_proj.addWidget(btn_save_proj)

        btn_open_excel = QPushButton("📂 打开结果Excel")
        btn_open_excel.clicked.connect(self.open_current_excel)
        lyt_proj.addWidget(btn_open_excel)
        lyt_proj.addStretch()
        main_layout.addWidget(group_proj)

        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)

        # 左侧：固定 60 通道
        left_widget = QWidget()
        lyt_left = QVBoxLayout(left_widget)
        lyt_left.setContentsMargins(0, 0, 0, 0)
        group_channels = QGroupBox("⚙️ 固定硬件通道 (打勾并命名)")
        lyt_ch = QVBoxLayout(group_channels)

        # 通道工具栏（新增一键清空）
        lyt_ch_tools = QHBoxLayout()
        btn_clear_channels = QPushButton("🗑️ 一键清空通道")
        btn_clear_channels.setObjectName("btn_danger")
        btn_clear_channels.clicked.connect(self.clear_channels)
        lyt_ch_tools.addWidget(btn_clear_channels)
        lyt_ch_tools.addStretch()
        lyt_ch.addLayout(lyt_ch_tools)

        self.table_channels = QTableWidget(60, 3)
        self.table_channels.setHorizontalHeaderLabels(["启用/硬件名", "自定义命名(输出用)", "报警限值(℃)"])
        self.table_channels.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table_channels.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table_channels.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        lyt_ch.addWidget(self.table_channels)

        # 初始化固定的 60 个通道
        hardware_channels = []
        for i in range(101, 121): hardware_channels.append(f"{i} (℃)")
        for i in range(201, 221): hardware_channels.append(f"{i} (℃)")
        for i in range(301, 321): hardware_channels.append(f"{i} (℃)")

        for idx, ch_name in enumerate(hardware_channels):
            w_chk = QWidget()
            l_chk = QHBoxLayout(w_chk)
            l_chk.setContentsMargins(5,0,0,0)
            chk = QCheckBox(ch_name)
            l_chk.addWidget(chk)
            self.table_channels.setCellWidget(idx, 0, w_chk)
            self.table_channels.setItem(idx, 1, QTableWidgetItem(""))
            self.table_channels.setItem(idx, 2, QTableWidgetItem(""))

        lyt_left.addWidget(group_channels)
        splitter.addWidget(left_widget)

        # 右侧：公共图片与工况
        right_widget = QWidget()
        lyt_right = QVBoxLayout(right_widget)
        lyt_right.setContentsMargins(0, 0, 0, 0)

        # 固定图片区 —— 带勾选框与插入状态
        group_fixed_imgs = QGroupBox("🖼️ 公共固定图片 (点温线实物图)")
        lyt_fimg = QVBoxLayout(group_fixed_imgs)
        lyt_fimg_tools = QHBoxLayout()
        btn_add_fimg = QPushButton("➕ 添加固定图片")
        btn_add_fimg.clicked.connect(self.add_fixed_image)
        btn_del_fimg = QPushButton("删除选中行")
        btn_del_fimg.setObjectName("btn_danger")
        btn_del_fimg.clicked.connect(self.del_fixed_image)
        btn_clear_fimg = QPushButton("🗑️ 一键清空图片")
        btn_clear_fimg.setObjectName("btn_danger")
        btn_clear_fimg.clicked.connect(self.clear_fixed_images)
        lyt_fimg_tools.addWidget(btn_add_fimg)
        lyt_fimg_tools.addWidget(btn_del_fimg)
        lyt_fimg_tools.addWidget(btn_clear_fimg)
        lyt_fimg_tools.addStretch()
        lyt_fimg.addLayout(lyt_fimg_tools)

        # 图片列表：[✔ 本次插入] [状态] [图片路径]
        self.table_fixed_imgs = QTableWidget(0, 3)
        self.table_fixed_imgs.setHorizontalHeaderLabels(["本次插入", "状态", "图片路径"])
        self.table_fixed_imgs.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table_fixed_imgs.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table_fixed_imgs.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table_fixed_imgs.setMaximumHeight(130)
        lyt_fimg.addWidget(self.table_fixed_imgs)
        lyt_right.addWidget(group_fixed_imgs)

        # 工况管理 —— 带启用勾选与插入状态
        group_conditions = QGroupBox("📊 独立工况管理 (一键建表用)")
        lyt_cond = QVBoxLayout(group_conditions)
        lyt_cond_tools = QHBoxLayout()
        btn_add_cond = QPushButton("➕ 添加工况")
        btn_add_cond.clicked.connect(self.add_condition)
        lyt_cond_tools.addWidget(btn_add_cond)
        lyt_cond_tools.addStretch()
        lyt_cond.addLayout(lyt_cond_tools)

        self.table_conds = QTableWidget(0, 6)
        self.table_conds.setHorizontalHeaderLabels(["启用", "状态", "工况名(表头)", "CSV 路径", "Power-Z 截图", "操作"])
        self.table_conds.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table_conds.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table_conds.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table_conds.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.table_conds.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table_conds.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        lyt_cond.addWidget(self.table_conds)
        lyt_right.addWidget(group_conditions)

        splitter.addWidget(right_widget)

        # 底部
        group_run = QGroupBox("🚀 一键全自动生成")
        lyt_run = QVBoxLayout(group_run)
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setMaximumHeight(80)
        lyt_run.addWidget(self.txt_log)

        lyt_run_tools = QHBoxLayout()
        self.progress_bar = QProgressBar()
        lyt_run_tools.addWidget(self.progress_bar)

        self.btn_run = QPushButton("⚡ 开始一键画表并填充数据")
        self.btn_run.setObjectName("btn_success")
        self.btn_run.setMinimumHeight(40)
        self.btn_run.clicked.connect(self.run_generation)
        lyt_run_tools.addWidget(self.btn_run)
        lyt_run.addLayout(lyt_run_tools)

        main_layout.addWidget(group_run)
        splitter.setSizes([450, 800])

    def log(self, msg):
        self.txt_log.append(msg)
        self.txt_log.moveCursor(self.txt_log.textCursor().End)

    def load_project_list(self):
        self.cbo_projects.blockSignals(True)
        self.cbo_projects.clear()
        import glob
        for f in glob.glob(os.path.join(self.cfg_dir, "*.json")):
            self.cbo_projects.addItem(os.path.basename(f).replace(".json", ""))
        self.cbo_projects.blockSignals(False)
        if self.cbo_projects.count() > 0:
            self.on_project_changed()

    def create_new_project(self):
        if not os.path.exists(self.base_template):
            QMessageBox.warning(self, "错误", f"找不到底层纯净模板：{self.base_template}")
            return

        proj_name, ok = QInputDialog.getText(self, "新建工程", "请输入工程名称（如 PB519）：")
        if not ok or not proj_name.strip(): return
        proj_name = proj_name.strip()

        target_dir = QFileDialog.getExistingDirectory(self, "选择工程存放的文件夹")
        if not target_dir: return

        excel_path = os.path.join(target_dir, f"{proj_name}_报告.xlsx")
        try:
            shutil.copy2(self.base_template, excel_path)
            self.current_project_excel = excel_path

            cfg_path = os.path.join(self.cfg_dir, f"{proj_name}.json")
            with open(cfg_path, "w", encoding="utf-8") as f:
                json.dump({"project_excel": excel_path}, f, ensure_ascii=False, indent=2)

            self.log(f"成功创建工程：{proj_name}")
            self.load_project_list()
            self.cbo_projects.setCurrentText(proj_name)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"创建失败：{str(e)}")

    def delete_project(self):
        proj_name = self.cbo_projects.currentText()
        if not proj_name: return

        reply = QMessageBox.question(
            self,
            '确认删除',
            f"确定要彻底删除工程【{proj_name}】的存档配置吗？\n(这只会删除本软件中的配置，不会删除你导出的 Excel 报告文件)",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            cfg_path = os.path.join(self.cfg_dir, f"{proj_name}.json")
            if os.path.exists(cfg_path):
                try:
                    os.remove(cfg_path)
                    self.log(f"已删除工程配置：{proj_name}")
                    self.load_project_list()
                    if self.cbo_projects.count() == 0:
                        self.current_project_excel = ""
                        self.table_conds.setRowCount(0)
                        self.list_fixed_imgs.clear()
                        for r in range(self.table_channels.rowCount()):
                            self.table_channels.cellWidget(r, 0).findChild(QCheckBox).setChecked(False)
                            self.table_channels.setItem(r, 1, QTableWidgetItem(""))
                            self.table_channels.setItem(r, 2, QTableWidgetItem(""))
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"删除失败：{str(e)}")

    def on_project_changed(self):
        proj_name = self.cbo_projects.currentText()
        if not proj_name: return
        cfg_path = os.path.join(self.cfg_dir, f"{proj_name}.json")
        if not os.path.exists(cfg_path): return

        with open(cfg_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)

        self.current_project_excel = cfg.get("project_excel", "")
        self.pending_image_deletions = cfg.get('deleted_image_ids', [])
        self.deleted_image_records = cfg.get('deleted_image_records', [])
        self.log(f"已切换到工程：{proj_name}")

        # 恢复通道
        channels = cfg.get("channels", [])
        for r in range(self.table_channels.rowCount()):
            w = self.table_channels.cellWidget(r, 0)
            chk = w.findChild(QCheckBox)
            h_name = chk.text()

            matched = next((c for c in channels if c.get("hardware_name") == h_name), None)
            if matched:
                chk.setChecked(matched.get("enabled", False))
                self.table_channels.setItem(r, 1, QTableWidgetItem(matched.get("custom_name", "")))
                self.table_channels.setItem(r, 2, QTableWidgetItem(matched.get("limit", "")))
            else:
                chk.setChecked(False)
                self.table_channels.setItem(r, 1, QTableWidgetItem(""))
                self.table_channels.setItem(r, 2, QTableWidgetItem(""))

        # 恢复固定图片（兼容旧版与新版，带状态）
        fimgs = cfg.get("fixed_images", [])
        self.table_fixed_imgs.setRowCount(0)
        for entry in fimgs:
            if isinstance(entry, str):
                path, enabled, status = entry, True, "待插入"
            else:
                path = entry.get("path", "")
                enabled = entry.get("enabled", True)
                status = entry.get("status", "待插入")
            if not path: continue
            self.insert_fixed_img_row(path, enabled, status,
                                      entry.get('id') if isinstance(entry, dict) else None)

        # 恢复工况（兼容旧版与新版，带启用与状态）
        conds = cfg.get("conditions", [])
        if isinstance(conds, dict):
            conds_list = []
            for k, v in conds.items():
                conds_list.append({"name": k, "csv": v.get("csv", ""), "powerz": v.get("powerz", ""), "enabled": True, "status": "待生成"})
            conds = conds_list

        self.table_conds.setRowCount(0)
        for c in conds:
            self.add_cond_row(
                c.get("name", ""),
                c.get("csv", ""),
                c.get("powerz", ""),
                c.get("enabled", True),
                c.get("status", "待生成"),
                c.get('id')
            )

    def clear_channels(self):
        reply = QMessageBox.question(
            self,
            "确认清空通道",
            "确定要清空所有已勾选的硬件通道、自定义命名和限值吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            for r in range(self.table_channels.rowCount()):
                w = self.table_channels.cellWidget(r, 0)
                if w:
                    chk = w.findChild(QCheckBox)
                    if chk: chk.setChecked(False)
                self.table_channels.setItem(r, 1, QTableWidgetItem(""))
                self.table_channels.setItem(r, 2, QTableWidgetItem(""))
            self.log("已清空所有硬件通道配置。")

    def clear_fixed_images(self):
        if self.table_fixed_imgs.rowCount() == 0:
            return
        reply = QMessageBox.question(
            self,
            "确认清空图片",
            "确定要清空所有公共固定图片吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            for r in range(self.table_fixed_imgs.rowCount()):
                self._queue_image_deletion(r)
            self.table_fixed_imgs.setRowCount(0)
            self.log("已清空图片列表；下次生成时会删除报告中对应的图片。")

    def insert_fixed_img_row(self, fpath, enabled=True, status="待插入", record_id=None):
        r = self.table_fixed_imgs.rowCount()
        self.table_fixed_imgs.insertRow(r)

        # 列0：勾选框
        chk_widget = QWidget()
        chk_layout = QHBoxLayout(chk_widget)
        chk_layout.setContentsMargins(0, 0, 0, 0)
        chk_layout.setAlignment(Qt.AlignCenter)
        chk = QCheckBox()
        chk.setChecked(enabled)
        chk_layout.addWidget(chk)
        self.table_fixed_imgs.setCellWidget(r, 0, chk_widget)

        # 列1：状态
        status_item = QTableWidgetItem(status)
        status_item.setTextAlignment(Qt.AlignCenter)
        status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
        if status == "已插入":
            status_item.setForeground(Qt.darkGreen)
            font = status_item.font()
            font.setBold(True)
            status_item.setFont(font)
        else:
            status_item.setForeground(Qt.gray)
        self.table_fixed_imgs.setItem(r, 1, status_item)

        # 列2：路径
        path_item = QTableWidgetItem(fpath)
        path_item.setData(Qt.UserRole, record_id or uuid.uuid4().hex)
        path_item.setFlags(path_item.flags() & ~Qt.ItemIsEditable)
        self.table_fixed_imgs.setItem(r, 2, path_item)

    def add_fixed_image(self):
        fpaths, _ = QFileDialog.getOpenFileNames(self, "选择固定图片 (可多选)", "", "Images (*.png *.jpg *.bmp *.jpeg)")
        for fpath in fpaths:
            if fpath:
                self.insert_fixed_img_row(fpath, enabled=True, status="待插入")

    def del_fixed_image(self):
        rows = sorted(set(idx.row() for idx in self.table_fixed_imgs.selectedIndexes()), reverse=True)
        for r in rows:
            self._queue_image_deletion(r)
            self.table_fixed_imgs.removeRow(r)
        if rows:
            self.log("下次生成时会删除报告中对应的图片，其他图片保留。")

    def _queue_image_deletion(self, row):
        item = self.table_fixed_imgs.item(row, 2)
        if item and item.data(Qt.UserRole):
            key = item.data(Qt.UserRole)
            if key not in self.pending_image_deletions:
                self.pending_image_deletions.append(key)
                status = self.table_fixed_imgs.item(row, 1)
                self.deleted_image_records.append({'id': key, 'path': item.text(),
                                                   'status': status.text() if status else ''})

    def add_condition(self):
        cname, ok = QInputDialog.getText(self, "添加工况", "请输入输出到表头的工况名\n(如: C口 15V/3A 充电):")
        if ok and cname.strip():
            self.add_cond_row(cname.strip())

    def add_cond_row(self, cname, csv_path="", pz_path="", enabled=True, status="待生成",
                     record_id=None):
        r = self.table_conds.rowCount()
        self.table_conds.insertRow(r)

        # 列0：启用复选框
        chk_widget = QWidget()
        chk_layout = QHBoxLayout(chk_widget)
        chk_layout.setContentsMargins(0, 0, 0, 0)
        chk_layout.setAlignment(Qt.AlignCenter)
        chk = QCheckBox()
        chk.setChecked(enabled)
        chk_layout.addWidget(chk)
        self.table_conds.setCellWidget(r, 0, chk_widget)

        # 列1：状态标签
        status_item = QTableWidgetItem(status)
        status_item.setTextAlignment(Qt.AlignCenter)
        status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
        if status == "已插入":
            status_item.setForeground(Qt.darkGreen)
            font = status_item.font()
            font.setBold(True)
            status_item.setFont(font)
        else:
            status_item.setForeground(Qt.gray)
        self.table_conds.setItem(r, 1, status_item)

        # 列2：工况名
        item = QTableWidgetItem(cname)
        item.setData(Qt.UserRole, record_id or uuid.uuid4().hex)
        self.table_conds.setItem(r, 2, item)

        # 文件选择器封装
        def make_path_picker(path, filter_str):
            w = QWidget()
            l = QHBoxLayout(w)
            l.setContentsMargins(0,0,0,0)
            txt = QLineEdit(path)
            btn = QPushButton("...")
            btn.clicked.connect(lambda _, t=txt: self.pick_file(t, filter_str))
            l.addWidget(txt)
            l.addWidget(btn)
            return w

        # 列3：CSV 路径
        self.table_conds.setCellWidget(r, 3, make_path_picker(csv_path, "CSV (*.csv)"))
        # 列4：Power-Z 截图
        self.table_conds.setCellWidget(r, 4, make_path_picker(pz_path, "Image (*.png *.jpg *.bmp)"))

        # 列5：操作
        btn_del = QPushButton("删除")
        btn_del.setObjectName("btn_danger")
        condition_id = item.data(Qt.UserRole)
        btn_del.clicked.connect(lambda _, key=condition_id: self.delete_condition(key))
        self.table_conds.setCellWidget(r, 5, btn_del)

    def _operation_running(self):
        return any(worker is not None and worker.isRunning()
                   for worker in (getattr(self, 'worker', None),
                                  getattr(self, 'delete_worker', None)))

    def closeEvent(self, event):
        if self._operation_running():
            QMessageBox.warning(self, '正在处理', '报告正在处理中，请完成后再关闭软件。')
            event.ignore()
        else:
            event.accept()

    def delete_condition(self, condition_id):
        if self._operation_running():
            QMessageBox.warning(self, '正在处理', '请等待当前报告操作完成后再删除工况。')
            return
        config = self.collect_project_config()
        if not config:
            QMessageBox.warning(self, '无法删除', '请先新建或选择工程。')
            return
        condition = next((c for c in config['conditions'] if c['id'] == condition_id), None)
        if condition is None:
            return
        reply = QMessageBox.question(
            self, '确认删除工况',
            f"确定删除工况【{condition['name']}】吗？\n\n"
            '将同时删除该工况的附表和总表数据，剩余工况向左排列。\n'
            '工程报告、其他工况及原始 CSV/图片文件会保留。',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply != QMessageBox.Yes:
            return
        config_path = os.path.join(self.cfg_dir, f'{self.cbo_projects.currentText()}.json')
        self.centralWidget().setEnabled(False)
        self.log(f"正在删除工况：{condition['name']}")
        self.delete_worker = DeleteConditionThread(
            self.current_project_excel, condition, config, config_path)
        self.delete_worker.log_updated.connect(self.log)
        self.delete_worker.succeeded.connect(self.on_condition_deleted)
        self.delete_worker.failed.connect(self.on_condition_delete_error)
        self.delete_worker.start()

    def on_condition_deleted(self):
        self.centralWidget().setEnabled(True)
        self.on_project_changed()
        self.log('工况已删除，报告已保存；新增工况生成时会接在剩余工况之后。')

    def on_condition_delete_error(self, message):
        self.centralWidget().setEnabled(True)
        self.log(f'删除失败：{message}')
        QMessageBox.critical(self, '删除失败', message)

    def pick_file(self, line_edit, filter_str):
        fpath, _ = QFileDialog.getOpenFileName(self, "选择文件", "", filter_str)
        if fpath: line_edit.setText(fpath)

    def collect_project_config(self):
        proj_name = self.cbo_projects.currentText()
        if not proj_name: return

        cfg = {"project_excel": self.current_project_excel, "channels": [], "fixed_images": [], "conditions": [],
               'deleted_image_ids': self.pending_image_deletions,
               'deleted_image_records': self.deleted_image_records}

        for r in range(self.table_channels.rowCount()):
            w = self.table_channels.cellWidget(r, 0)
            chk = w.findChild(QCheckBox) if w else None
            h_name = chk.text() if chk else ""
            enabled = chk.isChecked() if chk else False
            c_name = self.table_channels.item(r, 1).text() if self.table_channels.item(r, 1) else ""
            limit = self.table_channels.item(r, 2).text() if self.table_channels.item(r, 2) else ""
            if enabled or c_name or limit:
                cfg["channels"].append({"hardware_name": h_name, "enabled": enabled, "custom_name": c_name, "limit": limit})

        # 保存固定图片列表（路径 + 勾选状态 + 插入状态）
        for r in range(self.table_fixed_imgs.rowCount()):
            item = self.table_fixed_imgs.item(r, 2)
            path = item.text() if item else ""
            chk_w = self.table_fixed_imgs.cellWidget(r, 0)
            checked = chk_w.findChild(QCheckBox).isChecked() if chk_w else True
            st_item = self.table_fixed_imgs.item(r, 1)
            status = st_item.text() if st_item else "待插入"
            if path:
                cfg["fixed_images"].append({"id": item.data(Qt.UserRole), "path": path,
                                            "enabled": checked, "status": status})

        # 保存工况列表（启用 + 状态 + 名称 + CSV + PowerZ）
        for r in range(self.table_conds.rowCount()):
            chk_w = self.table_conds.cellWidget(r, 0)
            enabled = chk_w.findChild(QCheckBox).isChecked() if chk_w else True
            st_item = self.table_conds.item(r, 1)
            status = st_item.text() if st_item else "待生成"
            cname = self.table_conds.item(r, 2).text() if self.table_conds.item(r, 2) else ""
            w_csv = self.table_conds.cellWidget(r, 3)
            csv_path = w_csv.findChild(QLineEdit).text() if w_csv else ""
            w_pz = self.table_conds.cellWidget(r, 4)
            pz_path = w_pz.findChild(QLineEdit).text() if w_pz else ""
            cfg["conditions"].append({
                "id": self.table_conds.item(r, 2).data(Qt.UserRole),
                "name": cname,
                "csv": csv_path,
                "powerz": pz_path,
                "enabled": enabled,
                "status": status
            })

        return cfg

    def save_project_config(self):
        cfg = self.collect_project_config()
        if cfg is None:
            return
        cfg_path = os.path.join(self.cfg_dir, f"{self.cbo_projects.currentText()}.json")
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        self.log(f"已保存当前配置。")
        return cfg

    def open_current_excel(self):
        if os.path.exists(self.current_project_excel):
            os.startfile(self.current_project_excel)
        else:
            QMessageBox.warning(self, "错误", "工程Excel不存在！")

    def run_generation(self):
        if self._operation_running():
            return
        if not os.path.exists(self.current_project_excel):
            QMessageBox.warning(self, "错误", "工程Excel不存在，请先新建！")
            return

        # 1. 检查已插入工况并弹出二次确认防覆盖
        already_inserted_names = []
        for r in range(self.table_conds.rowCount()):
            chk_w = self.table_conds.cellWidget(r, 0)
            enabled = chk_w.findChild(QCheckBox).isChecked() if chk_w else False
            if enabled:
                st_item = self.table_conds.item(r, 1)
                st_text = st_item.text() if st_item else ""
                if st_text == "已插入":
                    cname = self.table_conds.item(r, 2).text() if self.table_conds.item(r, 2) else f"第{r+1}行"
                    already_inserted_names.append(cname)

        if already_inserted_names:
            reply = QMessageBox.question(
                self,
                "⚠️ 覆盖提示",
                f"检测到以下工况此前已生成过数据：\n\n【{'、'.join(already_inserted_names)}】\n\n重新生成将会覆盖原有的工作表与数据，确定要继续吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        project_config = self.save_project_config()

        # 2. 收集通道配置
        channels_config = []
        for r in range(self.table_channels.rowCount()):
            w = self.table_channels.cellWidget(r, 0)
            chk = w.findChild(QCheckBox)
            if chk and chk.isChecked():
                h_name = chk.text()
                c_name = self.table_channels.item(r, 1).text() if self.table_channels.item(r, 1) else ""
                limit = self.table_channels.item(r, 2).text() if self.table_channels.item(r, 2) else ""
                if not c_name.strip():
                    QMessageBox.warning(self, "警告", f"勾选的通道 '{h_name}' 没有填写自定义命名！")
                    return
                channels_config.append({"hardware_name": h_name, "custom_name": c_name.strip(), "limit": limit})

        if not channels_config:
            QMessageBox.warning(self, "错误", "至少需要勾选并配置一个硬件通道！")
            return

        # 3. 收集本次勾选的工况
        cond_configs = []
        self.active_generating_cond_rows = []
        for r in range(self.table_conds.rowCount()):
            chk_w = self.table_conds.cellWidget(r, 0)
            enabled = chk_w.findChild(QCheckBox).isChecked() if chk_w else False
            if not enabled:
                continue
            cname = self.table_conds.item(r, 2).text() if self.table_conds.item(r, 2) else ""
            w_csv = self.table_conds.cellWidget(r, 3)
            csv_path = w_csv.findChild(QLineEdit).text() if w_csv else ""
            w_pz = self.table_conds.cellWidget(r, 4)
            pz_path = w_pz.findChild(QLineEdit).text() if w_pz else ""
            if not csv_path:
                continue
            cond_configs.append({"id": self.table_conds.item(r, 2).data(Qt.UserRole),
                                 "name": cname, "csv": csv_path, "powerz": pz_path})
            self.active_generating_cond_rows.append(r)

        if not cond_configs:
            QMessageBox.warning(self, "错误", "至少需要勾选一个配有 CSV 的工况！")
            return

        # 4. 收集本次勾选要插入的固定图片
        fixed_images = []
        self.active_generating_img_rows = []
        for r in range(self.table_fixed_imgs.rowCount()):
            chk_w = self.table_fixed_imgs.cellWidget(r, 0)
            enabled = chk_w.findChild(QCheckBox).isChecked() if chk_w else False
            if enabled:
                item = self.table_fixed_imgs.item(r, 2)
                if item and item.text():
                    fixed_images.append({'id': item.data(Qt.UserRole), 'path': item.text()})
                    self.active_generating_img_rows.append(r)

        self.centralWidget().setEnabled(False)
        self.btn_run.setEnabled(False)
        self.progress_bar.setValue(0)

        self.worker = WorkerThread(self.current_project_excel, cond_configs, fixed_images,
                                   channels_config, project_config)
        self.worker.progress_updated.connect(self.progress_bar.setValue)
        self.worker.log_updated.connect(self.log)
        self.worker.finished_signal.connect(self.on_process_finished)
        self.worker.error_signal.connect(self.on_process_error)
        self.worker.start()

    def on_process_finished(self):
        self.centralWidget().setEnabled(True)
        self.pending_image_deletions = []
        self.deleted_image_records = []
        self.btn_run.setEnabled(True)
        self.progress_bar.setValue(100)

        # 成功后：将本次生成的工况标记为绿色已插入，并自动取消勾选
        for r in getattr(self, "active_generating_cond_rows", []):
            if r < self.table_conds.rowCount():
                chk_w = self.table_conds.cellWidget(r, 0)
                if chk_w:
                    chk = chk_w.findChild(QCheckBox)
                    if chk: chk.setChecked(False)  # 自动取消勾选
                st_item = self.table_conds.item(r, 1)
                if st_item:
                    st_item.setText("已插入")
                    st_item.setForeground(Qt.darkGreen)
                    font = st_item.font()
                    font.setBold(True)
                    st_item.setFont(font)

        # 成功后：将本次插入的固定图片标记为绿色已插入，并自动取消勾选
        for r in getattr(self, "active_generating_img_rows", []):
            if r < self.table_fixed_imgs.rowCount():
                chk_w = self.table_fixed_imgs.cellWidget(r, 0)
                if chk_w:
                    chk = chk_w.findChild(QCheckBox)
                    if chk: chk.setChecked(False)  # 自动取消勾选
                st_item = self.table_fixed_imgs.item(r, 1)
                if st_item:
                    st_item.setText("已插入")
                    st_item.setForeground(Qt.darkGreen)
                    font = st_item.font()
                    font.setBold(True)
                    st_item.setFont(font)

        self.save_project_config()
        QMessageBox.information(self, "完成", "批量制作成功！已为您自动标记已插入项目并更新界面。即将打开报告。")
        self.open_current_excel()

    def on_process_error(self, err_msg):
        self.centralWidget().setEnabled(True)
        self.btn_run.setEnabled(True)
        self.log(f"执行失败: {err_msg}")
        QMessageBox.critical(self, "生成失败", err_msg)

if __name__ == "__main__":
    configure_qt_plugins()
    app = QApplication(sys.argv)
    window = ThermalApp()
    window.show()
    sys.exit(app.exec_())
