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
import csv
import io
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

from doc_filler import parse_tags, fill_template, create_sample_template, load_tag_labels
from xlsx_filler import (
    parse_tags as parse_xlsx_tags,
    fill_template as fill_xlsx_template,
    create_sample_template as create_xlsx_sample,
    load_tag_labels as load_xlsx_labels,
)

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
        self._field_widgets = {}          # tag_name → (QLineEdit | QTextEdit)
        self._loop_tables = {}            # loop_var → QTableWidget
        self._loop_fields = {}            # loop_var → [field_names]
        self._custom_labels = {}          # user-edited labels (in-app, highest priority)
        self._json_labels = {}            # labels from .labels.json file
        self._batch_mode = False          # 批量填充模式
        self._batch_table = None          # 批量模式的 QTableWidget
        self._batch_tags = []            # 批量模式的标签列表
        self._batch_loops = {}           # 批量模式的循环（暂不支持批量填充）

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
        self.btn_batch.clicked.connect(self._toggle_batch_mode)
        top_bar.addWidget(self.btn_batch)

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
                self._json_labels = load_xlsx_labels(path)
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
            self._custom_labels = {}
            self._batch_tags = tags
            self._batch_loops = loops
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
        self._batch_table = None
        while self.form_layout.count():
            item = self.form_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # 批量模式：构建批量表格
        if self._batch_mode and tags:
            self._build_batch_form(tags, loops)
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
            group.setToolTip("双击标题可修改显示名")
            group.installEventFilter(self)
            layout = QVBoxLayout(group)
            layout.setContentsMargins(10, 8, 10, 10)

            long_tags = {'overview', 'technical_plan', 'implementation',
                        'budget', 'conclusion', 'content', 'description',
                        'summary', 'abstract', 'details', 'remarks'}
            if tag in long_tags or tag.endswith('_text') or tag.endswith('_plan'):
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
        if tag in self._json_labels and self._json_labels[tag]:
            return self._json_labels[tag]
        return self._tag_to_label(tag)

    @staticmethod
    def _tag_to_label(tag):
        """将标签名转为友好的中文显示名"""
        mapping = {
            'title': '文档标题', 'author': '编制人', 'date': '日期',
            'department': '部门', 'project_name': '项目名称',
            'doc_number': '编号', 'reviewer': '审核人', 'company': '单位名称',
            'phone': '联系电话', 'overview': '项目概况', 'budget': '预算与资源',
            'technical_plan': '技术方案', 'implementation': '实施计划',
            'conclusion': '结论与建议',
        }
        return mapping.get(tag, tag.replace('_', ' ').title())

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
        """弹出 QInputDialog 让用户修改字段的显示名称，并自动写回 .labels.json"""
        current = group_box.title()
        new_label, ok = QInputDialog.getText(
            self, "重命名字段",
            f"变量名：{tag}\n新的显示名称：",
            text=current
        )
        if ok and new_label.strip():
            new_label = new_label.strip()
            self._custom_labels[tag] = new_label
            self._json_labels[tag] = new_label    # 同步到 JSON 缓存
            group_box.setTitle(new_label)

            # 自动写回 .labels.json
            self._save_labels_json()

    def _save_labels_json(self):
        """将当前标签映射写回模板同目录的 .labels.json 文件"""
        if not self._template_path:
            return
        json_path = os.path.splitext(self._template_path)[0] + '.labels.json'
        try:
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(self._json_labels, f, ensure_ascii=False, indent=2)
        except (IOError, PermissionError):
            pass  # 静默失败：文件只读 / 目录无写入权限等情况

    # ── 文档生成 ────────────────────────────────────────────

    def _on_generate(self):
        """收集输入 → 填充模板 → 导出文档"""
        if not self._template_path:
            return

        if self._batch_mode:
            self._on_batch_generate()
            return

        # 收集简单变量
        context = {}
        for tag, widget in self._field_widgets.items():
            if isinstance(widget, QTextEdit):
                context[tag] = widget.toPlainText()
            else:
                context[tag] = widget.text()

        # 收集循环表格数据
        for loop_var, table in self._loop_tables.items():
            fields = self._loop_fields.get(loop_var, [])
            items = []
            for row in range(table.rowCount()):
                item_data = {}
                for col in range(table.columnCount()):
                    field_name = fields[col] if col < len(fields) else f"col{col}"
                    cell_item = table.item(row, col)
                    value = cell_item.text().strip() if cell_item else ''
                    # 尝试转换数值
                    try:
                        if '.' in value:
                            value = float(value)
                        else:
                            value = int(value)
                    except ValueError:
                        pass
                    item_data[field_name] = value
                # 跳过全空行
                if any(v != '' and v is not None for v in item_data.values()):
                    items.append(item_data)
            if items:
                context[loop_var] = items

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

    # ── 批量填充 ────────────────────────────────────────────

    def _toggle_batch_mode(self):
        """切换单人 / 批量模式"""
        if not self._template_path:
            QMessageBox.information(self, "提示", "请先加载一个模板文件。")
            return
        self._batch_mode = not self._batch_mode
        if self._batch_mode:
            # 从单人模式收集当前填值作为默认值
            self._batch_defaults = {}
            for tag, widget in self._field_widgets.items():
                val = widget.toPlainText() if isinstance(widget, QTextEdit) else widget.text()
                if val.strip():
                    self._batch_defaults[tag] = val.strip()

            self.btn_batch.setText("◆ 批量模式")
            self.btn_batch.setStyleSheet("""
                QPushButton { background: #e67e22; color: white; font-weight: bold;
                             border-radius: 6px; padding: 8px 14px; font-size: 12px;
                             border: 1px solid #d35400; }
                QPushButton:hover { background: #d35400; }
            """)
            self.btn_generate.setText("批量生成")
            self.btn_generate.setToolTip("每行数据生成一个独立文件")
        else:
            self.btn_batch.setText("◇ 单人模式")
            self.btn_batch.setStyleSheet("""
                QPushButton { background: #f0c040; color: #333; font-weight: bold;
                             border-radius: 6px; padding: 8px 14px; font-size: 12px;
                             border: 1px solid #d4a020; }
                QPushButton:hover { background: #e0b030; }
            """)
            self.btn_generate.setText("生成文档")
            self.btn_generate.setToolTip("")
        self._build_form(self._batch_tags, self._batch_loops)

    def _build_batch_form(self, tags, loops):
        """批量模式：构建大数据表格，每行 = 一个输出文件"""
        defaults = getattr(self, '_batch_defaults', {})

        # 提示标签
        hint_lines = ["▎批量填充模式 — 每行数据生成一个独立文件"]
        if defaults:
            hint_lines.append(f"（已从单人模式带入 {len(defaults)} 个默认值）")
        hint = QLabel("\n".join(hint_lines))
        hint.setStyleSheet("font-size: 13px; color: #e67e22; font-weight: bold; padding: 4px 0;")
        self.form_layout.addWidget(hint)

        # 操作栏
        btn_row = QHBoxLayout()

        btn_csv = QPushButton("导入 CSV / Excel")
        btn_csv.setStyleSheet("""
            QPushButton { background: #3498db; color: white; font-weight: bold;
                         border-radius: 4px; padding: 6px 14px; font-size: 12px; }
            QPushButton:hover { background: #2980b9; }
        """)
        btn_csv.clicked.connect(self._import_csv)
        btn_row.addWidget(btn_csv)

        btn_paste = QPushButton("粘贴数据")
        btn_paste.setStyleSheet("""
            QPushButton { background: #9b59b6; color: white; font-weight: bold;
                         border-radius: 4px; padding: 6px 14px; font-size: 12px; }
            QPushButton:hover { background: #8e44ad; }
        """)
        btn_paste.clicked.connect(self._paste_clipboard)
        btn_row.addWidget(btn_paste)

        btn_row.addSpacing(15)

        btn_add = QPushButton("+ 添加行")
        btn_add.setStyleSheet("""
            QPushButton { background: #27ae60; color: white; font-weight: bold;
                         border-radius: 4px; padding: 6px 14px; font-size: 12px; }
            QPushButton:hover { background: #219955; }
        """)
        btn_add.clicked.connect(self._add_batch_row)
        btn_row.addWidget(btn_add)

        btn_del = QPushButton("- 删除选中行")
        btn_del.setStyleSheet("""
            QPushButton { background: #e74c3c; color: white; font-weight: bold;
                         border-radius: 4px; padding: 6px 14px; font-size: 12px; }
            QPushButton:hover { background: #c0392b; }
        """)
        btn_del.clicked.connect(self._remove_batch_row)
        btn_row.addWidget(btn_del)

        if defaults:
            btn_apply = QPushButton("应用默认值")
            btn_apply.setStyleSheet("""
                QPushButton { background: #f39c12; color: white; font-weight: bold;
                             border-radius: 4px; padding: 6px 14px; font-size: 12px; }
                QPushButton:hover { background: #e67e22; }
            """)
            btn_apply.clicked.connect(self._apply_defaults)
            btn_row.addWidget(btn_apply)

        btn_row.addStretch()
        self.form_layout.addLayout(btn_row)

        # 批量表格
        table = QTableWidget()
        table.setColumnCount(len(tags))
        headers = [self._resolve_label(t) for t in tags]
        table.setHorizontalHeaderLabels(headers)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setMinimumSectionSize(60)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setAlternatingRowColors(True)
        table.setStyleSheet("""
            QTableWidget { border: 1px solid #d0d5dd; border-radius: 4px;
                          gridline-color: #e8edf2; font-size: 12px; }
            QTableWidget::item { padding: 4px 6px; }
            QHeaderView::section { background: #f0f2f5; font-weight: bold;
                                  padding: 4px; border: 1px solid #d0d5dd; }
        """)
        table.verticalHeader().setDefaultSectionSize(32)
        table.setMinimumHeight(200)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        # 初始 3 个空行（含默认值）
        for _ in range(3):
            self._add_batch_row_to(table, defaults)

        self._batch_table = table
        self.form_layout.addWidget(table)

        # 循环区域提示（批量模式暂不支持）
        if loops:
            loop_warn = QLabel(
                "⚠ 模板含循环区域（{% for %}），批量模式暂不支持循环数据填充。\n"
                "循环区域在批量生成时将被跳过。如需循环，请使用单人模式。")
            loop_warn.setStyleSheet("color: #e67e22; font-size: 11px; padding: 4px 0;")
            self.form_layout.addWidget(loop_warn)

        self.form_layout.addStretch()

    def _add_batch_row_to(self, table, defaults=None):
        """向指定表格添加一行，可选填入默认值"""
        row = table.rowCount()
        table.insertRow(row)
        table.setRowHeight(row, 32)
        if defaults:
            for tag, val in defaults.items():
                if tag in self._batch_tags:
                    col = self._batch_tags.index(tag)
                    item = QTableWidgetItem(val)
                    table.setItem(row, col, item)

    def _add_batch_row(self):
        if self._batch_table:
            self._add_batch_row_to(self._batch_table, getattr(self, '_batch_defaults', {}))

    def _apply_defaults(self):
        """将默认值应用到所有行"""
        defaults = getattr(self, '_batch_defaults', {})
        if not defaults or not self._batch_table:
            return
        table = self._batch_table
        for row in range(table.rowCount()):
            for tag, val in defaults.items():
                if tag in self._batch_tags:
                    col = self._batch_tags.index(tag)
                    existing = table.item(row, col)
                    if not existing or not existing.text().strip():
                        item = QTableWidgetItem(val)
                        table.setItem(row, col, item)
        self.status_bar.showMessage(f"已将 {len(defaults)} 个默认值应用到所有空单元格")

    def _remove_batch_row(self):
        if not self._batch_table:
            return
        rows = set()
        for item in self._batch_table.selectedItems():
            rows.add(item.row())
        for row in sorted(rows, reverse=True):
            self._batch_table.removeRow(row)

    def _import_csv(self):
        """从 CSV 或 Excel 文件导入批量数据"""
        path, _ = QFileDialog.getOpenFileName(
            self, "导入数据文件", os.path.expanduser("~\\Desktop"),
            "表格文件 (*.csv *.xlsx *.xls);;CSV 文件 (*.csv);;Excel 工作簿 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        if not path:
            return

        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in ('.xlsx', '.xls'):
                import openpyxl
                wb = openpyxl.load_workbook(path, data_only=True)
                ws = wb.active
                rows_data = []
                for row in ws.iter_rows(values_only=True):
                    rows_data.append([str(v) if v is not None else '' for v in row])
                wb.close()
            else:
                # CSV: 尝试常见编码
                rows_data = []
                for encoding in ['utf-8-sig', 'utf-8', 'gbk', 'gb2312']:
                    try:
                        with open(path, 'r', encoding=encoding) as f:
                            reader = csv.reader(f)
                            rows_data = [list(row) for row in reader]
                        break
                    except (UnicodeDecodeError, Exception):
                        continue
        except Exception as e:
            QMessageBox.critical(self, "导入失败", f"无法读取文件：{e}")
            return

        if not rows_data:
            QMessageBox.warning(self, "提示", "文件中没有数据。")
            return

        # 第一行可能是表头 — 尝试匹配模板变量名
        header_row = rows_data[0]
        # 将模板标签映射到列索引
        tag_to_col = {}
        for tag in self._batch_tags:
            label_lower = self._resolve_label(tag).lower()
            for col_idx, h in enumerate(header_row):
                h_lower = h.strip().lower().replace(' ', '_').replace('（', '').replace('）', '')
                if h_lower == tag.lower() or h_lower == label_lower:
                    tag_to_col[tag] = col_idx
                    break

        # 如果至少有一个匹配，跳过表头行
        data_start = 1 if tag_to_col else 0
        if not tag_to_col:
            # 表头不匹配，尝试按列顺序对应
            for idx, tag in enumerate(self._batch_tags):
                if idx < len(header_row):
                    tag_to_col[tag] = idx

        # 清空表格，重新填充
        table = self._batch_table
        table.setRowCount(0)

        for row_data in rows_data[data_start:]:
            if not any(v.strip() for v in row_data):
                continue  # 跳过空行
            row = table.rowCount()
            table.insertRow(row)
            table.setRowHeight(row, 32)
            for tag, col_idx in tag_to_col.items():
                if col_idx < len(row_data):
                    value = row_data[col_idx].strip()
                    if value:
                        tag_col = self._batch_tags.index(tag)
                        item = QTableWidgetItem(value)
                        table.setItem(row, tag_col, item)

        self.status_bar.showMessage(f"已导入 {table.rowCount()} 行数据")

    def _paste_clipboard(self):
        """从剪贴板粘贴表格数据"""
        clip = QApplication.clipboard()
        text = clip.text()
        if not text:
            return

        lines = text.strip().split('\n')
        rows_data = []
        for line in lines:
            # 支持 Tab 或逗号分隔
            if '\t' in line:
                row = line.split('\t')
            elif ',' in line:
                row = line.split(',')
            else:
                row = [line]
            rows_data.append([v.strip() for v in row])

        # 尝试匹配表头
        tag_to_col = {}
        if rows_data:
            header_row = rows_data[0]
            for tag in self._batch_tags:
                label_lower = self._resolve_label(tag).lower()
                for col_idx, h in enumerate(header_row):
                    h_lower = h.lower()
                    if h_lower == tag.lower() or h_lower == label_lower:
                        tag_to_col[tag] = col_idx
                        break

        data_start = 1 if tag_to_col else 0
        if not tag_to_col:
            for idx, tag in enumerate(self._batch_tags):
                if idx < len(header_row):
                    tag_to_col[tag] = idx

        table = self._batch_table
        table.setRowCount(0)

        for row_data in rows_data[data_start:]:
            if not any(v for v in row_data):
                continue
            row = table.rowCount()
            table.insertRow(row)
            table.setRowHeight(row, 32)
            for tag, col_idx in tag_to_col.items():
                if col_idx < len(row_data):
                    value = row_data[col_idx]
                    if value:
                        tag_col = self._batch_tags.index(tag)
                        item = QTableWidgetItem(value)
                        table.setItem(row, tag_col, item)

        self.status_bar.showMessage(f"已粘贴 {table.rowCount()} 行数据")

    def _on_batch_generate(self):
        """批量生成：每行数据生成一个独立文件"""
        table = self._batch_table
        if not table or table.rowCount() == 0:
            QMessageBox.warning(self, "提示", "批量表格中没有数据行。")
            return

        # 收集所有行数据
        row_contexts = []
        for row in range(table.rowCount()):
            ctx = {}
            has_data = False
            for col in range(table.columnCount()):
                tag = self._batch_tags[col]
                cell_item = table.item(row, col)
                value = cell_item.text().strip() if cell_item else ''
                if value:
                    has_data = True
                ctx[tag] = value
            if has_data:
                row_contexts.append(ctx)

        if not row_contexts:
            QMessageBox.warning(self, "提示", "所有数据行均为空。")
            return

        # 确认
        reply = QMessageBox.question(
            self, "确认批量生成",
            f"将为 {len(row_contexts)} 行数据各生成一个文件，保存到桌面。\n\n确定继续？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        # 逐行生成
        success = 0
        errors = []
        fill_func = fill_xlsx_template if self._template_format == 'xlsx' else fill_template

        for idx, ctx in enumerate(row_contexts):
            try:
                base = os.path.splitext(os.path.basename(self._template_path))[0]
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                # 文件名：模板名_序号_时间戳
                ts_short = datetime.now().strftime('%H%M%S')
                output_path = os.path.join(
                    os.path.expanduser("~"), "Desktop",
                    f"{base}_{idx + 1:03d}_{ts_short}.{self._template_format}"
                )
                fill_func(self._template_path, ctx, output_path)
                success += 1
            except Exception as e:
                errors.append(f"第 {idx + 1} 行: {e}")

        if errors:
            QMessageBox.warning(
                self, "批量生成完成",
                f"成功：{success} 个文件\n失败：{len(errors)} 个\n\n"
                + "\n".join(errors[:5])
            )
        else:
            QMessageBox.information(
                self, "批量生成完成",
                f"已生成 {success} 个文件，保存在桌面上。"
            )
            self.status_bar.showMessage(f"批量生成完成：{success} 个文件 → 桌面")


def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei", 10))
    window = TemplateFillWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
