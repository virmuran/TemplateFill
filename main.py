#!/usr/bin/env python
"""
TemplateFill — 模板驱动的 Office 文档生成工具

流程：
  1. 点击「打开模板」加载 .docx 模板
  2. 自动检测模板中所有 {{tag}} 占位符
  3. 为每个标签生成输入框
  4. 填写内容后点击「生成文档」
  5. 导出格式化的 DOCX 到桌面
"""

import os
import sys
import json
from pathlib import Path
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QScrollArea,
    QFileDialog, QMessageBox, QGroupBox, QGridLayout,
    QStatusBar, QSizePolicy, QFrame, QInputDialog,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
)
from PySide6.QtCore import Qt, QEvent
from PySide6.QtGui import QFont, QIcon

from doc_filler import parse_tags, fill_template, create_sample_template
from xlsx_filler import (
    parse_tags as parse_xlsx_tags,
    fill_template as fill_xlsx_template,
    create_sample_template as create_xlsx_sample,
)
from batch_manager import BatchManager
from labels import tag_to_label, load_tag_labels, get_label, get_type, is_long_text_tag, parse_entry

# ── 样式 ──
MAIN_STYLE = """
QMainWindow { background: #f5f6f8; }
QGroupBox {
    font-size: 13px; font-weight: bold; color: #2c3e50;
    border: 1px solid #d0d5dd; border-radius: 8px;
    margin-top: 12px; padding-top: 16px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px; padding: 0 6px;
}
QLineEdit, QTextEdit {
    border: 1px solid #c0c5ce; border-radius: 5px;
    padding: 6px 10px; font-size: 13px;
    background: white;
}
QLineEdit:focus, QTextEdit:focus { border-color: #4a6fa5; }
QScrollArea { border: none; background: transparent; }
QScrollBar:vertical { background: transparent; width: 8px; }
QScrollBar::handle:vertical { background: #c0c0c0; border-radius: 4px; min-height: 30px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
"""

BUTTON_PRIMARY = """
QPushButton {
    background: #4a6fa5; color: white; font-weight: bold;
    border-radius: 6px; padding: 10px 24px; font-size: 14px;
}
QPushButton:hover { background: #3a5a8c; }
QPushButton:disabled { background: #b0b8c4; }
"""

BUTTON_SECONDARY = """
QPushButton {
    background: #e8edf2; color: #2c3e50;
    border-radius: 6px; padding: 8px 16px; font-size: 13px;
    border: 1px solid #c0c5ce;
}
QPushButton:hover { background: #d0dae5; }
"""

BUTTON_SUCCESS = """
QPushButton {
    background: #27ae60; color: white; font-weight: bold;
    border-radius: 6px; padding: 10px 24px; font-size: 14px;
}
QPushButton:hover { background: #219955; }
QPushButton:disabled { background: #a0d8b0; }
"""

BUTTON_PROJECT = """
QPushButton {
    background: #f5f6f8; color: #4a6fa5; font-weight: bold;
    border-radius: 6px; padding: 8px 14px; font-size: 12px;
    border: 1px solid #4a6fa5;
}
QPushButton:hover { background: #e8edf5; }
"""

TEMPLATE_LABEL_STYLE = """
QLabel {
    color: #666; font-size: 12px;
    background: #f0f2f5; border-radius: 4px;
    padding: 6px 10px;
}
"""

FIELD_LABEL_STYLE = "font-size: 13px; color: #2c3e50; font-weight: bold; padding-right: 8px;"


class TemplateFillWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("TemplateFill — 模板文档生成器")
        self.setMinimumSize(900, 650)
        self.resize(1000, 700)

        self._template_path = None
        self._template_format = None      # 'docx' or 'xlsx'
        self._project_path = None         # 当前项目文件路径（.tplfill）
        self._field_widgets = {}          # tag_name → (QLineEdit | QTextEdit)
        self._loop_tables = {}            # loop_var → QTableWidget
        self._loop_fields = {}            # loop_var → [field_names]
        self._custom_labels = {}          # user-edited labels (in-app, highest priority)
        self._custom_types = {}           # user-edited widget types (in-app)
        self._json_labels = {}            # labels from .labels.json file (parsed entries)

        self.batch = BatchManager(self)   # 批量填充管理器

        self._setup_ui()
        self._auto_load_sample()

    # ── UI 构建 ────────────────────────────────────────────

    def _setup_ui(self):
        self.setStyleSheet(MAIN_STYLE)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(12)
        root.setContentsMargins(16, 16, 16, 16)

        # ── 顶部：模板操作 ──
        top_bar = QHBoxLayout()
        top_bar.setSpacing(10)

        self.btn_open = QPushButton("打开模板...")
        self.btn_open.setStyleSheet(BUTTON_PRIMARY)
        self.btn_open.clicked.connect(self._on_open_template)
        top_bar.addWidget(self.btn_open)

        self.btn_sample = QPushButton("生成 Word 示例")
        self.btn_sample.setStyleSheet(BUTTON_SECONDARY)
        self.btn_sample.clicked.connect(self._on_create_sample)
        top_bar.addWidget(self.btn_sample)

        self.btn_xlsx_sample = QPushButton("生成 Excel 示例")
        self.btn_xlsx_sample.setStyleSheet(BUTTON_SECONDARY)
        self.btn_xlsx_sample.clicked.connect(self._on_create_xlsx_sample)
        top_bar.addWidget(self.btn_xlsx_sample)

        self.btn_batch = QPushButton("◇ 单人模式")
        self.btn_batch.setStyleSheet("""
            QPushButton { background: #f0c040; color: #333; font-weight: bold;
                         border-radius: 6px; padding: 8px 14px; font-size: 12px;
                         border: 1px solid #d4a020; }
            QPushButton:hover { background: #e0b030; }
        """)
        self.btn_batch.clicked.connect(self.batch.toggle)
        top_bar.addWidget(self.btn_batch)

        # 项目保存 / 加载
        self.btn_save_project = QPushButton("保存项目")
        self.btn_save_project.setStyleSheet(BUTTON_PROJECT)
        self.btn_save_project.setToolTip("将当前填写的数据保存为 .tplfill 项目文件，方便下次继续编辑")
        self.btn_save_project.clicked.connect(self._save_project)
        top_bar.addWidget(self.btn_save_project)

        self.btn_open_project = QPushButton("打开项目")
        self.btn_open_project.setStyleSheet(BUTTON_PROJECT)
        self.btn_open_project.setToolTip("打开之前保存的 .tplfill 项目文件，恢复编辑状态")
        self.btn_open_project.clicked.connect(self._open_project)
        top_bar.addWidget(self.btn_open_project)

        self.lbl_template = QLabel("未加载模板")
        self.lbl_template.setStyleSheet(TEMPLATE_LABEL_STYLE)
        self.lbl_template.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        top_bar.addWidget(self.lbl_template, 1)

        root.addLayout(top_bar)

        # 分隔线
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #e0e0e0;")
        root.addWidget(sep)

        # ── 中间：动态表单（可滚动） ──
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        self.form_container = QWidget()
        self.form_layout = QVBoxLayout(self.form_container)
        self.form_layout.setSpacing(10)
        self.form_layout.setContentsMargins(4, 4, 4, 4)

        self.placeholder_label = QLabel(
            "点击「打开模板」加载 .docx 或 .xlsx 文件，\n"
            "或点击「生成示例模板」快速体验。\n\n"
            "Word 模板中的 {{变量名}} 会自动识别为输入字段；\n"
            "Excel 模板支持简单变量和 {% for %} 循环区域。")
        self.placeholder_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.placeholder_label.setStyleSheet("color: #999; font-size: 14px; padding: 60px;")
        self.form_layout.addWidget(self.placeholder_label)
        self.form_layout.addStretch()

        self.scroll_area.setWidget(self.form_container)
        root.addWidget(self.scroll_area, 1)

        # ── 底部：操作栏 ──
        bottom_bar = QHBoxLayout()
        bottom_bar.setSpacing(10)

        self.lbl_status = QLabel("就绪")
        self.lbl_status.setStyleSheet("color: #888; font-size: 12px;")
        bottom_bar.addWidget(self.lbl_status, 1)

        self.btn_generate = QPushButton("生成文档")
        self.btn_generate.setStyleSheet(BUTTON_SUCCESS)
        self.btn_generate.setEnabled(False)
        self.btn_generate.clicked.connect(self._on_generate)
        bottom_bar.addWidget(self.btn_generate)

        root.addLayout(bottom_bar)

        # 状态栏
        self.status_bar = QStatusBar()
        self.status_bar.showMessage("欢迎使用 TemplateFill — 模板驱动文档生成器")
        self.setStatusBar(self.status_bar)

    # ── 模板加载 ────────────────────────────────────────────

    def _auto_load_sample(self):
        """启动时检查是否有示例模板可用"""
        sample_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "templates", "sample_report.docx"
        )
        if os.path.exists(sample_path):
            self._load_template(sample_path)

    def _on_open_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择模板文件", os.path.expanduser("~\\Desktop"),
            "Office 模板 (*.docx *.xlsx);;Word 文档 (*.docx);;Excel 工作簿 (*.xlsx);;所有文件 (*.*)"
        )
        if path:
            self._load_template(path)

    def _on_create_sample(self):
        """生成示例 Word 模板并自动加载"""
        try:
            sample_dir = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "templates")
            os.makedirs(sample_dir, exist_ok=True)
            sample_path = os.path.join(sample_dir, "sample_report.docx")
            create_sample_template(sample_path)
            QMessageBox.information(self, "成功",
                                    f"示例模板已生成：\n{sample_path}\n\n已自动加载，可以开始填写。")
            self._load_template(sample_path)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成示例模板失败：{e}")

    def _on_create_xlsx_sample(self):
        """生成示例 Excel 模板并自动加载"""
        try:
            sample_dir = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "templates")
            os.makedirs(sample_dir, exist_ok=True)
            sample_path = os.path.join(sample_dir, "sample_quote.xlsx")
            create_xlsx_sample(sample_path)
            QMessageBox.information(self, "成功",
                                    f"Excel 示例模板已生成：\n{sample_path}\n\n"
                                    "包含简单变量和循环区域，已自动加载。")
            self._load_template(sample_path)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成 Excel 示例失败：{e}")

    def _load_template(self, path):
        """加载模板并生成表单，自动识别 .docx / .xlsx"""
        try:
            ext = os.path.splitext(path)[1].lower()

            if ext == '.xlsx':
                self._template_format = 'xlsx'
                result = parse_xlsx_tags(path)
                tags = result['simple']
                loops = result['loops']
                self._json_labels = load_tag_labels(path)
            else:
                self._template_format = 'docx'
                tags = parse_tags(path)
                loops = {}
                self._json_labels = load_tag_labels(path)

            if not tags and not loops:
                QMessageBox.warning(self, "提示",
                                    "模板中未检测到 {{变量}} 占位符。\n\n"
                                    "请在模板中使用 {{变量名}} 格式插入占位符。")
                return

            total_fields = len(tags) + len(loops)
            fmt_name = "Excel" if self._template_format == 'xlsx' else "Word"
            self._template_path = path
            self._project_path = None
            self._custom_labels = {}
            self._custom_types = {}
            self.batch.tags = tags
            self.batch.loops = loops
            self.lbl_template.setText(
                f"模板：{os.path.basename(path)}  ({fmt_name} · {total_fields} 个字段)")
            self.lbl_template.setStyleSheet(
                "color: #4a6fa5; font-weight: bold; font-size: 12px; "
                "background: #e8edf5; border-radius: 4px; padding: 6px 10px;")

            self._build_form(tags, loops)
            self.btn_generate.setEnabled(True)
            self.status_bar.showMessage(f"已加载{fmt_name}模板：{path}")

        except Exception as e:
            QMessageBox.critical(self, "加载失败", str(e))

    # ── 表单构建 ────────────────────────────────────────────

    def _build_form(self, tags, loops=None):
        """根据标签列表动态生成输入表单。
        tags: 简单变量列表
        loops: 循环区域 dict {loop_var: {fields, item_var, ...}}
        """
        # 清空旧表单
        self._field_widgets.clear()
        self._loop_tables.clear()
        self._loop_fields.clear()
        self.batch.table = None
        while self.form_layout.count():
            item = self.form_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # 批量模式：构建批量表格
        if self.batch.mode and tags:
            self.batch.build_form(tags, loops)
            return

        # ── 简单变量 ──
        priority = ['title', 'project_name', 'department', 'author',
                    'date', 'doc_number', 'company', 'reviewer', 'phone']
        ordered = [t for t in priority if t in tags]
        ordered += [t for t in tags if t not in priority]

        for tag in ordered:
            display_name = self._resolve_label(tag)
            group = QGroupBox(display_name)
            group.setStyleSheet("""
                QGroupBox { font-size: 12px; font-weight: bold; color: #4a6fa5;
                           border: 1px solid #e0e4ea; border-radius: 6px;
                           margin-top: 8px; padding-top: 14px; }
                QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
            """)
            group.setProperty("tag", tag)
            group.setToolTip("双击标题可修改显示名和控件类型")
            group.installEventFilter(self)
            layout = QVBoxLayout(group)
            layout.setContentsMargins(10, 8, 10, 10)

            # 控件类型：用户配置 > JSON 文件 > 变量名兜底
            widget_type = self._resolve_type(tag)
            if widget_type == 'multi':
                widget = QTextEdit()
                widget.setMinimumHeight(60)
                widget.setMaximumHeight(120)
                widget.setPlaceholderText(f"请输入{display_name}...")
            else:
                widget = QLineEdit()
                widget.setPlaceholderText(f"请输入{display_name}")

            layout.addWidget(widget)
            self._field_widgets[tag] = widget
            self.form_layout.addWidget(group)

        # ── 循环区域 ──
        if loops:
            # 分隔线
            loop_sep = QFrame()
            loop_sep.setFrameShape(QFrame.Shape.HLine)
            loop_sep.setStyleSheet("color: #4a6fa5; margin: 8px 0;")
            self.form_layout.addWidget(loop_sep)

            loop_label = QLabel("▎循环区域（可添加多行数据）")
            loop_label.setStyleSheet("font-size: 13px; color: #27ae60; font-weight: bold; padding: 4px 0;")
            self.form_layout.addWidget(loop_label)

            for loop_var, loop_info in loops.items():
                fields = loop_info['fields']
                self._build_loop_section(loop_var, fields, loop_info)

        self.form_layout.addStretch()

    def _build_loop_section(self, loop_var, fields, loop_info):
        """为循环区域构建 QTableWidget 输入区"""
        item_var = loop_info.get('item_var', 'item')

        display_name = self._resolve_label(loop_var)
        group = QGroupBox(f"{display_name}（循环：{{% for {item_var} in {loop_var} %}}）")
        group.setStyleSheet("""
            QGroupBox { font-size: 12px; font-weight: bold; color: #27ae60;
                       border: 2px solid #a0d8b0; border-radius: 8px;
                       margin-top: 12px; padding-top: 16px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
        """)
        layout = QVBoxLayout(group)
        layout.setContentsMargins(10, 8, 10, 10)

        btn_row = QHBoxLayout()
        btn_add = QPushButton("+ 添加行")
        btn_add.setStyleSheet("""
            QPushButton { background: #27ae60; color: white; font-weight: bold;
                         border-radius: 4px; padding: 6px 14px; font-size: 12px; }
            QPushButton:hover { background: #219955; }
        """)
        btn_add.clicked.connect(lambda checked, lv=loop_var, f=fields: self._add_loop_row(lv, f))
        btn_row.addWidget(btn_add)

        btn_del = QPushButton("- 删除选中行")
        btn_del.setStyleSheet("""
            QPushButton { background: #e74c3c; color: white; font-weight: bold;
                         border-radius: 4px; padding: 6px 14px; font-size: 12px; }
            QPushButton:hover { background: #c0392b; }
        """)
        btn_del.clicked.connect(lambda checked, lv=loop_var: self._remove_loop_row(lv))
        btn_row.addWidget(btn_del)

        btn_row.addStretch()
        layout.addLayout(btn_row)

        table = QTableWidget()
        table.setColumnCount(len(fields))
        field_labels = [f.replace('_', ' ').title() for f in fields]
        table.setHorizontalHeaderLabels(field_labels)
        # 首列（序号）窄一些，其余列可拖拽调整，最后一列自适应填充
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setMinimumSectionSize(60)
        for col_idx in range(len(fields)):
            if fields[col_idx] in ('seq', 'index', 'no', 'id'):
                table.setColumnWidth(col_idx, 50)
            else:
                table.setColumnWidth(col_idx, 120)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setAlternatingRowColors(True)
        table.setStyleSheet("""
            QTableWidget { border: 1px solid #d0d5dd; border-radius: 4px;
                          gridline-color: #e8edf2; font-size: 12px; }
            QTableWidget::item { padding: 4px 6px; }
            QHeaderView::section { background: #f0f2f5; font-weight: bold;
                                  padding: 4px; border: 1px solid #d0d5dd; }
        """)
        table.setMinimumHeight(160)
        table.setMaximumHeight(400)
        table.verticalHeader().setDefaultSectionSize(40)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(table)

        self._loop_tables[loop_var] = table
        self._loop_fields[loop_var] = fields
        self.form_layout.addWidget(group)

    def _add_loop_row(self, loop_var, fields):
        """在循环表格中添加一行"""
        table = self._loop_tables.get(loop_var)
        if not table:
            return
        row = table.rowCount()
        table.insertRow(row)
        table.setRowHeight(row, 40)
        if fields and fields[0] in ('seq', 'index', 'no', 'id'):
            item = QTableWidgetItem(str(row + 1))
            table.setItem(row, 0, item)
        table.scrollToBottom()

    def _remove_loop_row(self, loop_var):
        """删除循环表格中选中的行"""
        table = self._loop_tables.get(loop_var)
        if not table:
            return
        rows = set()
        for item in table.selectedItems():
            rows.add(item.row())
        for row in sorted(rows, reverse=True):
            table.removeRow(row)

    def _resolve_label(self, tag):
        """解析显示名：自定义编辑 > JSON 文件 > 代码映射 > 自动生成"""
        if tag in self._custom_labels:
            return self._custom_labels[tag]
        # JSON 文件中查找（兼容新旧格式）
        label = get_label(self._json_labels, tag)
        if label:
            return label
        return tag_to_label(tag)

    def _resolve_type(self, tag):
        """解析控件类型：自定义编辑 > JSON 文件 > 变量名兜底判断"""
        if tag in self._custom_types:
            return self._custom_types[tag]
        return get_type(self._json_labels, tag)

    # ── 标签改名（双击） ─────────────────────────────────────

    def eventFilter(self, obj, event):
        """双击 GroupBox 标题 → 弹出改名对话框"""
        if event.type() == QEvent.MouseButtonDblClick:
            tag = obj.property("tag")
            if tag:
                self._edit_label(tag, obj)
                return True
        return super().eventFilter(obj, event)

    def _edit_label(self, tag, group_box):
        """弹出对话框让用户修改显示名和控件类型（单行/多行），并自动写回 .labels.json"""
        from PySide6.QtWidgets import QDialog, QDialogButtonBox, QComboBox, QFormLayout

        current_label = self._resolve_label(tag)
        current_type = self._resolve_type(tag)

        dlg = QDialog(self)
        dlg.setWindowTitle("编辑字段")
        dlg.setMinimumWidth(340)
        form = QFormLayout(dlg)

        name_edit = QLineEdit(current_label)
        name_edit.setPlaceholderText(f"变量名：{tag}")
        form.addRow("显示名称：", name_edit)

        type_combo = QComboBox()
        type_combo.addItem("单行文本", "single")
        type_combo.addItem("多行文本", "multi")
        type_combo.setCurrentIndex(0 if current_type == 'single' else 1)
        form.addRow("控件类型：", type_combo)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        form.addRow(buttons)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        new_label = name_edit.text().strip()
        new_type = type_combo.currentData()
        if not new_label:
            return

        self._custom_labels[tag] = new_label
        self._custom_types[tag] = new_type

        # 写入 _json_labels（新格式）
        old_entry = self._json_labels.get(tag)
        old_label_entry = parse_entry(old_entry, tag)
        self._json_labels[tag] = {
            'label': new_label,
            'type': new_type,
        }

        group_box.setTitle(new_label)

        # 如果控件类型变了，重建表单以应用新的控件
        if new_type != current_type:
            self._build_form(self.batch.tags, self.batch.loops)

        # 自动写回 .labels.json
        self._save_labels_json()

    def _save_labels_json(self):
        """将当前标签映射写回 .labels.json（新格式：每条为 {label, type} dict）"""
        if not self._template_path:
            return
        json_path = os.path.splitext(self._template_path)[0] + '.labels.json'
        try:
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(self._json_labels, f, ensure_ascii=False, indent=2)
        except (IOError, PermissionError):
            pass

    # ── 项目保存 / 加载 ─────────────────────────────────────

    def _collect_context(self):
        """收集当前表单所有数据，返回序列化友好的 dict"""
        context = {}
        for tag, widget in self._field_widgets.items():
            if isinstance(widget, QTextEdit):
                context[tag] = widget.toPlainText()
            else:
                context[tag] = widget.text()

        loops = {}
        for loop_var, table in self._loop_tables.items():
            fields = self._loop_fields.get(loop_var, [])
            items = []
            for row in range(table.rowCount()):
                item_data = {}
                for col in range(table.columnCount()):
                    field_name = fields[col] if col < len(fields) else f"col{col}"
                    cell_item = table.item(row, col)
                    item_data[field_name] = cell_item.text().strip() if cell_item else ''
                if any(v != '' for v in item_data.values()):
                    items.append(item_data)
            loops[loop_var] = items

        return context, loops

    def _save_project(self):
        """将当前填写数据 + 标签配置保存为 .tplfill 项目文件"""
        if not self._template_path:
            QMessageBox.information(self, "提示", "请先加载一个模板文件。")
            return

        context, loops = self._collect_context()

        # 收集批量表格数据（如果处于批量模式）
        batch_data = None
        if self.batch.mode and self.batch.table:
            batch_data = []
            for row in range(self.batch.table.rowCount()):
                row_data = {}
                for col in range(self.batch.table.columnCount()):
                    if col < len(self.batch.tags):
                        tag = self.batch.tags[col]
                        item = self.batch.table.item(row, col)
                        row_data[tag] = item.text().strip() if item else ''
                if any(v != '' for v in row_data.values()):
                    batch_data.append(row_data)

        # 模板路径存为相对路径（如果可能）
        try:
            project_dir = os.path.dirname(os.path.abspath(self._project_path or __file__))
            template_path = os.path.relpath(self._template_path, project_dir)
        except ValueError:
            template_path = self._template_path

        data = {
            'version': 1,
            'template_path': template_path.replace('\\', '/'),
            'template_format': self._template_format,
            'context': context,
            'loops': loops,
            'custom_labels': self._custom_labels,
            'custom_types': self._custom_types,
            'batch_mode': self.batch.mode,
            'batch_defaults': self.batch.defaults,
            'batch_data': batch_data,
        }

        default_name = os.path.splitext(os.path.basename(self._template_path))[0] + '.tplfill'
        if self._project_path:
            default_dir = os.path.dirname(self._project_path)
            default_name = os.path.basename(self._project_path)
        else:
            default_dir = os.path.expanduser("~\\Desktop")

        path, _ = QFileDialog.getSaveFileName(
            self, "保存项目", os.path.join(default_dir, default_name),
            "TemplateFill 项目 (*.tplfill);;所有文件 (*.*)"
        )
        if not path:
            return

        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self._project_path = path
            self._update_title()
            self.status_bar.showMessage(f"项目已保存：{path}")
        except (IOError, PermissionError) as e:
            QMessageBox.critical(self, "保存失败", f"无法写入文件：{e}")

    def _open_project(self):
        """打开 .tplfill 项目文件，恢复编辑状态"""
        path, _ = QFileDialog.getOpenFileName(
            self, "打开项目", os.path.expanduser("~\\Desktop"),
            "TemplateFill 项目 (*.tplfill);;所有文件 (*.*)"
        )
        if not path:
            return

        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            QMessageBox.critical(self, "打开失败", f"无法读取项目文件：{e}")
            return

        # 验证基本结构
        if not isinstance(data, dict) or 'template_path' not in data:
            QMessageBox.critical(self, "格式错误", "项目文件格式不正确。")
            return

        # 解析模板路径（支持相对路径）
        project_dir = os.path.dirname(os.path.abspath(path))
        template_path = os.path.normpath(os.path.join(project_dir, data['template_path']))
        if not os.path.exists(template_path):
            # 尝试用原始路径
            template_path = data['template_path']
            if not os.path.exists(template_path):
                reply = QMessageBox.question(
                    self, "模板未找到",
                    f"项目引用的模板文件不存在：\n{template_path}\n\n是否手动定位模板文件？",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.Yes:
                    new_path, _ = QFileDialog.getOpenFileName(
                        self, "选择模板文件", os.path.expanduser("~\\Desktop"),
                        "Office 模板 (*.docx *.xlsx);;所有文件 (*.*)"
                    )
                    if not new_path:
                        return
                    template_path = new_path
                else:
                    return

        # 加载模板
        self._load_template(template_path)

        # 恢复自定义标签配置
        if data.get('custom_labels'):
            self._custom_labels = data['custom_labels']
        if data.get('custom_types'):
            self._custom_types = data['custom_types']

        # 恢复简单变量数据
        context = data.get('context', {})
        for tag, widget in self._field_widgets.items():
            if tag in context:
                val = str(context[tag]) if context[tag] is not None else ''
                if isinstance(widget, QTextEdit):
                    widget.setPlainText(val)
                else:
                    widget.setText(val)

        # 恢复循环表格数据
        loops_data = data.get('loops', {})
        for loop_var, items in loops_data.items():
            table = self._loop_tables.get(loop_var)
            fields = self._loop_fields.get(loop_var, [])
            if not table or not fields:
                continue
            # 清空现有行（保留表头）
            table.setRowCount(0)
            for item_data in items:
                row = table.rowCount()
                table.insertRow(row)
                table.setRowHeight(row, 40)
                for col, field_name in enumerate(fields):
                    if col >= table.columnCount():
                        break
                    value = item_data.get(field_name, '')
                    item = QTableWidgetItem(str(value) if value else '')
                    table.setItem(row, col, item)

        # 恢复批量模式
        if data.get('batch_mode'):
            # 切换到批量模式（如果尚未处于）
            if not self.batch.mode:
                self.batch.mode = True
                self.batch.defaults = data.get('batch_defaults', {})
                self.batch.build_form(self.batch.tags, self.batch.loops)
                # 填充批量表格数据
                batch_data = data.get('batch_data')
                if batch_data and self.batch.table:
                    self.batch.table.setRowCount(0)
                    for row_data in batch_data:
                        row = self.batch.table.rowCount()
                        self.batch.table.insertRow(row)
                        self.batch.table.setRowHeight(row, 32)
                        for tag, val in row_data.items():
                            if tag in self.batch.tags:
                                col = self.batch.tags.index(tag)
                                item = QTableWidgetItem(str(val) if val else '')
                                self.batch.table.setItem(row, col, item)
                # 更新批量按钮样式
                self.batch.mode = True  # 重建 build_form 后状态
                # 这里 batch.mode 已经被 build_form 设置过，toggle 逻辑需要...
                # 直接设置已保存的状态
                self.btn_batch.setText("◆ 批量模式")
                self.btn_batch.setStyleSheet("""
                    QPushButton { background: #e67e22; color: white; font-weight: bold;
                                 border-radius: 6px; padding: 8px 14px; font-size: 12px;
                                 border: 1px solid #d35400; }
                    QPushButton:hover { background: #d35400; }
                """)
                self.btn_generate.setText("批量生成")
                self.btn_generate.setToolTip("每行数据生成一个独立文件")

        # 重建表单以应用自定义显示名
        self._build_form(self.batch.tags, self.batch.loops)

        # 重新填充重建表单后的数据（因为 build_form 会清空表单）
        context = data.get('context', {})
        for tag, widget in self._field_widgets.items():
            if tag in context:
                val = str(context[tag]) if context[tag] is not None else ''
                if isinstance(widget, QTextEdit):
                    widget.setPlainText(val)
                else:
                    widget.setText(val)

        self._project_path = path
        self._update_title()
        self.status_bar.showMessage(f"已打开项目：{path}")

    def _update_title(self):
        """更新窗口标题，显示项目文件名"""
        base = "TemplateFill — 模板文档生成器"
        if self._project_path:
            name = os.path.basename(self._project_path)
            self.setWindowTitle(f"{base}  [{name}]")
        else:
            self.setWindowTitle(base)

    # ── 文档生成 ────────────────────────────────────────────

    def _on_generate(self):
        """收集输入 → 填充模板 → 导出文档"""
        if not self._template_path:
            return

        if self.batch.mode:
            self.batch._on_batch_generate()
            return

        context, loops_raw = self._collect_context()

        # 将循环数据的字符串值转为数值（与 _collect_context 纯字符串语义分离）
        loops = {}
        for loop_var, items in loops_raw.items():
            typed_items = []
            for item_data in items:
                typed = {}
                for k, v in item_data.items():
                    try:
                        typed[k] = float(v) if '.' in str(v) else int(v)
                    except (ValueError, TypeError):
                        typed[k] = v
                typed_items.append(typed)
            if typed_items:
                loops[loop_var] = typed_items
        context.update(loops)

        try:
            if self._template_format == 'xlsx':
                out_path = fill_xlsx_template(self._template_path, context)
            else:
                out_path = fill_template(self._template_path, context)

            self.status_bar.showMessage(f"文档已生成：{out_path}")
            QMessageBox.information(
                self, "生成成功",
                f"文档已保存到：\n{out_path}\n\n请打开查看。"
            )
        except Exception as e:
            QMessageBox.critical(self, "生成失败", str(e))

def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei", 10))
    window = TemplateFillWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
