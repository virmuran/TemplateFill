"""
TemplateFill — 动态章节编辑器

让用户在软件中直接增删章节和子项，而不必修改模板文件。
每个章节 = 一个一级标题，下面可自由添加子标题（H2/H3）、富文本段落、
表格和图片。

数据结构:
    sections = [
        {
            'title': '第一章 概述',
            'blocks': [
                {'type': 'heading', 'level': 2, 'text': '1.1 背景'},
                {'type': 'text', 'html': '<p>正文内容...</p>'},
                {'type': 'table', 'header': True,
                 'data': [['列A', '列B'], ['1', '2']]},
                {'type': 'image', 'path': 'C:/img.png',
                 'width': 12.0, 'caption': '图 1 示意'},
            ]
        },
    ]
"""

import os
import copy

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QTextEdit, QComboBox, QScrollArea, QFrame,
    QSizePolicy, QListWidget, QListWidgetItem, QMessageBox,
    QFileDialog, QSpinBox, QDoubleSpinBox, QCheckBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QApplication,
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QPixmap, QColor


# ── 样式 ────────────────────────────────────────────────

SECTION_EDITOR_STYLE = """
QFrame#blockFrame {
    border: 1px solid #e0e4ea;
    border-radius: 4px;
    background: white;
}
QListWidget#sectionList {
    border: 1px solid #d0d5dd;
    border-radius: 4px;
    background: white;
    font-size: 13px;
    outline: none;
}
QListWidget#sectionList::item {
    padding: 8px 12px;
    border-bottom: 1px solid #eef0f3;
}
QListWidget#sectionList::item:selected {
    background: #e8edf5;
    color: #4a6fa5;
    font-weight: bold;
}
QListWidget#sectionList::item:hover {
    background: #f0f2f5;
}
"""

BTN_BLUE = """
QPushButton {
    background: #3498db; color: white; font-weight: bold;
    border-radius: 4px; padding: 6px 14px; font-size: 12px;
    border: none;
}
QPushButton:hover { background: #2980b9; }
"""

BTN_GREEN = """
QPushButton {
    background: #27ae60; color: white; font-weight: bold;
    border-radius: 4px; padding: 6px 14px; font-size: 12px;
    border: none;
}
QPushButton:hover { background: #219955; }
"""

BTN_RED = """
QPushButton {
    background: #e74c3c; color: white; font-weight: bold;
    border-radius: 4px; padding: 6px 14px; font-size: 12px;
    border: none;
}
QPushButton:hover { background: #c0392b; }
"""

BTN_GRAY = """
QPushButton {
    background: #e8edf2; color: #2c3e50; font-weight: bold;
    border-radius: 4px; padding: 6px 14px; font-size: 12px;
    border: 1px solid #c0c5ce;
}
QPushButton:hover { background: #d0dae5; }
QPushButton:pressed { background: #c0c5ce; }
"""

BTN_TOOLBAR = """
QPushButton {
    background: #f0f2f5; color: #2c3e50;
    border-radius: 3px; padding: 4px 8px; font-size: 12px;
    border: 1px solid #d0d5dd;
    min-width: 28px;
}
QPushButton:hover { background: #e0e4ea; }
QPushButton:checked { background: #4a6fa5; color: white; }
"""


class TextBlockWidget(QFrame):
    """富文本段落编辑块 — 含格式工具栏 + QTextEdit"""

    def __init__(self, block_data, on_change, on_delete, index, parent=None):
        super().__init__(parent)
        self.setObjectName("blockFrame")
        self._block_data = block_data
        self._on_change = on_change
        self._index = index
        self._setup_ui(on_delete)

    def _setup_ui(self, on_delete):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 8)
        layout.setSpacing(4)

        # 头部：标签 + 工具栏 + 删除
        header = QHBoxLayout()
        header.setSpacing(6)

        type_label = QLabel("📄 段落")
        type_label.setStyleSheet("font-size: 12px; font-weight: bold; color: #888;")
        header.addWidget(type_label)

        # 格式工具栏
        self.btn_bold = QPushButton("B")
        self.btn_bold.setCheckable(True)
        self.btn_bold.setStyleSheet(BTN_TOOLBAR + "font-weight: bold;")
        self.btn_bold.setMaximumWidth(36)
        self.btn_bold.setToolTip("加粗 (Ctrl+B)")
        self.btn_bold.clicked.connect(self._toggle_bold)
        header.addWidget(self.btn_bold)

        self.btn_italic = QPushButton("I")
        self.btn_italic.setCheckable(True)
        self.btn_italic.setStyleSheet(BTN_TOOLBAR + "font-style: italic;")
        self.btn_italic.setMaximumWidth(36)
        self.btn_italic.setToolTip("斜体 (Ctrl+I)")
        self.btn_italic.clicked.connect(self._toggle_italic)
        header.addWidget(self.btn_italic)

        self.btn_underline = QPushButton("U")
        self.btn_underline.setCheckable(True)
        self.btn_underline.setStyleSheet(BTN_TOOLBAR + "text-decoration: underline;")
        self.btn_underline.setMaximumWidth(36)
        self.btn_underline.setToolTip("下划线 (Ctrl+U)")
        self.btn_underline.clicked.connect(self._toggle_underline)
        header.addWidget(self.btn_underline)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.VLine)
        sep.setStyleSheet("color: #d0d5dd;")
        header.addWidget(sep)

        self.btn_bullet = QPushButton("• 列表")
        self.btn_bullet.setStyleSheet(BTN_TOOLBAR)
        self.btn_bullet.setToolTip("无序列表")
        self.btn_bullet.clicked.connect(self._insert_bullet_list)
        header.addWidget(self.btn_bullet)

        self.btn_numbered = QPushButton("1. 列表")
        self.btn_numbered.setStyleSheet(BTN_TOOLBAR)
        self.btn_numbered.setToolTip("有序列表")
        self.btn_numbered.clicked.connect(self._insert_numbered_list)
        header.addWidget(self.btn_numbered)

        header.addStretch()

        btn_del = QPushButton("×")
        btn_del.setStyleSheet(BTN_RED)
        btn_del.setMaximumWidth(32)
        btn_del.setToolTip("删除此子项")
        btn_del.clicked.connect(on_delete)
        header.addWidget(btn_del)

        layout.addLayout(header)

        # 编辑器
        self.editor = QTextEdit()
        self.editor.setMinimumHeight(80)
        self.editor.setMaximumHeight(200)
        self.editor.setPlaceholderText("输入段落内容...")
        self.editor.setStyleSheet("font-size: 13px; border: 1px solid #d0d5dd; border-radius: 3px;")
        self.editor.setHtml(self._block_data.get('html', ''))
        self.editor.textChanged.connect(self._on_text_changed)
        self.editor.cursorPositionChanged.connect(self._update_toolbar_state)
        layout.addWidget(self.editor)

    def _toggle_bold(self):
        self.editor.setFontWeight(
            QFont.Weight.Bold if self.editor.fontWeight() == QFont.Weight.Normal
            else QFont.Weight.Normal
        )

    def _toggle_italic(self):
        self.editor.setFontItalic(not self.editor.fontItalic())

    def _toggle_underline(self):
        self.editor.setFontUnderline(not self.editor.fontUnderline())

    def _insert_bullet_list(self):
        cursor = self.editor.textCursor()
        cursor.insertHtml('<ul><li></li></ul><p></p>')

    def _insert_numbered_list(self):
        cursor = self.editor.textCursor()
        cursor.insertHtml('<ol><li></li></ol><p></p>')

    def _on_text_changed(self):
        self._block_data['html'] = self.editor.toHtml()
        self._on_change()

    def _update_toolbar_state(self):
        """根据光标位置更新工具栏按钮状态"""
        self.btn_bold.setChecked(self.editor.fontWeight() == QFont.Weight.Bold)
        self.btn_italic.setChecked(self.editor.fontItalic())
        self.btn_underline.setChecked(self.editor.fontUnderline())


class HeadingBlockWidget(QFrame):
    """子标题编辑块"""

    def __init__(self, block_data, on_change, on_delete, index, parent=None):
        super().__init__(parent)
        self.setObjectName("blockFrame")
        self._block_data = block_data
        self._on_change = on_change
        self._index = index
        self._setup_ui(on_delete)

    def _setup_ui(self, on_delete):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 8)
        layout.setSpacing(4)

        header = QHBoxLayout()
        header.setSpacing(6)

        type_label = QLabel("📝 子标题")
        type_label.setStyleSheet("font-size: 12px; font-weight: bold; color: #888;")
        header.addWidget(type_label)

        level_label = QLabel("级别：")
        level_label.setStyleSheet("font-size: 11px; color: #888;")
        header.addWidget(level_label)

        self.level_combo = QComboBox()
        self.level_combo.addItem("H2", 2)
        self.level_combo.addItem("H3", 3)
        self.level_combo.addItem("H4", 4)
        self.level_combo.setStyleSheet("font-size: 12px;")
        current_level = self._block_data.get('level', 2)
        idx = self.level_combo.findData(current_level)
        if idx >= 0:
            self.level_combo.setCurrentIndex(idx)
        self.level_combo.currentIndexChanged.connect(self._on_level_changed)
        header.addWidget(self.level_combo)

        header.addStretch()

        btn_del = QPushButton("×")
        btn_del.setStyleSheet(BTN_RED)
        btn_del.setMaximumWidth(32)
        btn_del.setToolTip("删除此子项")
        btn_del.clicked.connect(on_delete)
        header.addWidget(btn_del)

        layout.addLayout(header)

        self.editor = QLineEdit()
        self.editor.setPlaceholderText("输入子标题文本...")
        self.editor.setStyleSheet("font-size: 13px;")
        self.editor.setText(self._block_data.get('text', ''))
        self.editor.textChanged.connect(self._on_text_changed)
        layout.addWidget(self.editor)

    def _on_level_changed(self):
        self._block_data['level'] = self.level_combo.currentData()
        self._on_change()

    def _on_text_changed(self):
        self._block_data['text'] = self.editor.text()
        self._on_change()


class TableBlockWidget(QFrame):
    """表格编辑块 — 行列可调 + 单元格直接编辑"""

    def __init__(self, block_data, on_change, on_delete, index, parent=None):
        super().__init__(parent)
        self.setObjectName("blockFrame")
        self._block_data = block_data
        self._on_change = on_change
        self._index = index
        self._loading = True
        self._setup_ui(on_delete)
        self._loading = False
        self._rebuild_table()

    def _setup_ui(self, on_delete):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 8)
        layout.setSpacing(4)

        # 头部
        header = QHBoxLayout()
        header.setSpacing(6)

        type_label = QLabel("📋 表格")
        type_label.setStyleSheet("font-size: 12px; font-weight: bold; color: #888;")
        header.addWidget(type_label)

        header.addStretch()

        btn_del = QPushButton("×")
        btn_del.setStyleSheet(BTN_RED)
        btn_del.setMaximumWidth(32)
        btn_del.setToolTip("删除此子项")
        btn_del.clicked.connect(on_delete)
        header.addWidget(btn_del)
        layout.addLayout(header)

        # 行列控制
        ctrl = QHBoxLayout()
        ctrl.setSpacing(4)

        ctrl.addWidget(QLabel("行："))
        self.rows_spin = QSpinBox()
        self.rows_spin.setRange(1, 50)
        data = self._block_data.get('data') or []
        self.rows_spin.setValue(len(data) if data else 3)
        self.rows_spin.setStyleSheet("font-size: 12px;")
        self.rows_spin.setMinimumWidth(56)
        self.rows_spin.valueChanged.connect(self._on_dims_changed)
        ctrl.addWidget(self.rows_spin)

        ctrl.addWidget(QLabel("列："))
        self.cols_spin = QSpinBox()
        self.cols_spin.setRange(1, 20)
        cols = max((len(r) for r in data), default=0) if data else 0
        self.cols_spin.setValue(cols if cols else 3)
        self.cols_spin.setStyleSheet("font-size: 12px;")
        self.cols_spin.setMinimumWidth(56)
        self.cols_spin.valueChanged.connect(self._on_dims_changed)
        ctrl.addWidget(self.cols_spin)

        self.header_check = QCheckBox("首行为表头")
        self.header_check.setChecked(bool(self._block_data.get('header', True)))
        self.header_check.setStyleSheet("font-size: 12px;")
        self.header_check.stateChanged.connect(self._on_header_changed)
        ctrl.addWidget(self.header_check)

        ctrl.addStretch()
        layout.addLayout(ctrl)

        # 单元格编辑区
        self.table = QTableWidget()
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setDefaultSectionSize(26)
        self.table.setMinimumHeight(100)
        self.table.setMaximumHeight(260)
        self.table.setStyleSheet("""
            QTableWidget { font-size: 12px; border: 1px solid #d0d5dd; border-radius: 3px; }
            QTableWidget::item { padding: 2px 4px; }
        """)
        self.table.cellChanged.connect(self._on_cell_changed)
        layout.addWidget(self.table)

    def _rebuild_table(self):
        """按 block_data['data'] 重建 QTableWidget（行列变化时同步数据）"""
        self._loading = True
        data = self._block_data.setdefault('data', [])
        n_rows = self.rows_spin.value()
        n_cols = self.cols_spin.value()

        # 数据尺寸与 spinbox 对齐（增删行列）
        while len(data) < n_rows:
            data.append([''] * n_cols)
        del data[n_rows:]
        for row in data:
            if len(row) < n_cols:
                row.extend([''] * (n_cols - len(row)))
            del row[n_cols:]

        self.table.setRowCount(n_rows)
        self.table.setColumnCount(n_cols)

        header_bold = self.header_check.isChecked()
        for r in range(n_rows):
            for c in range(n_cols):
                item = QTableWidgetItem(str(data[r][c]))
                if header_bold and r == 0:
                    f = item.font()
                    f.setBold(True)
                    item.setFont(f)
                    item.setBackground(QColor('#e8edf5'))
                self.table.setItem(r, c, item)
        self._loading = False

    def _on_dims_changed(self):
        if self._loading:
            return
        self._rebuild_table()
        self._on_change()

    def _on_header_changed(self):
        if self._loading:
            return
        self._block_data['header'] = self.header_check.isChecked()
        self._rebuild_table()
        self._on_change()

    def _on_cell_changed(self, row, col):
        if self._loading:
            return
        item = self.table.item(row, col)
        if item is None:
            return
        data = self._block_data.setdefault('data', [])
        while len(data) <= row:
            data.append([])
        while len(data[row]) <= col:
            data[row].append('')
        data[row][col] = item.text()
        self._on_change()


class ImageBlockWidget(QFrame):
    """图片编辑块 — 选择本地图片 + 宽度 + 图注"""

    PREVIEW_STYLE = (
        "border: 1px dashed #c0c5ce; border-radius: 3px;"
        "color: #aaa; font-size: 11px; background: #fafbfc;"
    )
    PREVIEW_MISSING_STYLE = (
        "border: 1px dashed #e74c3c; border-radius: 3px;"
        "color: #e74c3c; font-size: 11px; background: #fafbfc;"
    )

    def __init__(self, block_data, on_change, on_delete, index, parent=None):
        super().__init__(parent)
        self.setObjectName("blockFrame")
        self._block_data = block_data
        self._on_change = on_change
        self._index = index
        self._setup_ui(on_delete)

    def _setup_ui(self, on_delete):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 8)
        layout.setSpacing(4)

        # 头部
        header = QHBoxLayout()
        header.setSpacing(6)

        type_label = QLabel("🖼 图片")
        type_label.setStyleSheet("font-size: 12px; font-weight: bold; color: #888;")
        header.addWidget(type_label)

        header.addStretch()

        btn_del = QPushButton("×")
        btn_del.setStyleSheet(BTN_RED)
        btn_del.setMaximumWidth(32)
        btn_del.setToolTip("删除此子项")
        btn_del.clicked.connect(on_delete)
        header.addWidget(btn_del)
        layout.addLayout(header)

        # 主体：缩略图 + 控制
        body = QHBoxLayout()
        body.setSpacing(8)

        self.preview = QLabel("未选择图片")
        self.preview.setFixedSize(140, 105)
        self.preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview.setStyleSheet(self.PREVIEW_STYLE)
        body.addWidget(self.preview)

        controls = QVBoxLayout()
        controls.setSpacing(6)

        btn_pick = QPushButton("📁 重新选择图片...")
        btn_pick.setStyleSheet(BTN_BLUE)
        btn_pick.clicked.connect(self._pick_image)
        controls.addWidget(btn_pick)

        width_row = QHBoxLayout()
        width_row.addWidget(QLabel("宽度："))
        self.width_spin = QDoubleSpinBox()
        self.width_spin.setRange(0.0, 30.0)
        self.width_spin.setDecimals(1)
        self.width_spin.setSuffix(" cm")
        try:
            self.width_spin.setValue(float(self._block_data.get('width', 0) or 0))
        except (TypeError, ValueError):
            self.width_spin.setValue(0.0)
        self.width_spin.setToolTip("0 = 按图片原始大小（最宽 15 cm）")
        self.width_spin.setStyleSheet("font-size: 12px;")
        self.width_spin.valueChanged.connect(self._on_width_changed)
        width_row.addWidget(self.width_spin)
        width_row.addStretch()
        controls.addLayout(width_row)

        cap_label = QLabel("图注（可选，居中显示在图片下方）：")
        cap_label.setStyleSheet("font-size: 11px; color: #888;")
        controls.addWidget(cap_label)

        self.caption_edit = QLineEdit()
        self.caption_edit.setPlaceholderText("例如：图 1 工艺流程示意")
        self.caption_edit.setStyleSheet("font-size: 12px;")
        self.caption_edit.setText(self._block_data.get('caption', ''))
        self.caption_edit.textChanged.connect(self._on_caption_changed)
        controls.addWidget(self.caption_edit)

        controls.addStretch()
        body.addLayout(controls, 1)
        layout.addLayout(body)

        self._update_preview()

    def _pick_image(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择图片", "",
            "图片文件 (*.png *.jpg *.jpeg *.bmp *.gif *.webp);;所有文件 (*)")
        if path:
            self._block_data['path'] = path
            self._update_preview()
            self._on_change()

    def _update_preview(self):
        """根据 path 更新缩略图 / 缺失提示"""
        path = self._block_data.get('path', '')
        self.preview.setPixmap(QPixmap())  # 清空旧图
        if path:
            if os.path.exists(path):
                pix = QPixmap(path)
                if not pix.isNull():
                    self.preview.setStyleSheet(self.PREVIEW_STYLE)
                    self.preview.setPixmap(pix.scaled(
                        138, 103,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation))
                    return
            self.preview.setStyleSheet(self.PREVIEW_MISSING_STYLE)
            self.preview.setText("⚠ 文件不存在")
        else:
            self.preview.setStyleSheet(self.PREVIEW_STYLE)
            self.preview.setText("未选择图片")

    def _on_width_changed(self):
        self._block_data['width'] = self.width_spin.value()
        self._on_change()

    def _on_caption_changed(self):
        self._block_data['caption'] = self.caption_edit.text()
        self._on_change()


class SectionEditor(QWidget):
    """动态章节编辑器

    信号:
        content_changed: 内容变化时发出
    """

    content_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._sections = []
        self._current_index = -1
        self._block_widgets = []
        self._setup_ui()

    def _setup_ui(self):
        self.setStyleSheet(SECTION_EDITOR_STYLE)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # 标题
        header = QLabel("▎动态章节编辑")
        header.setStyleSheet(
            "font-size: 14px; font-weight: bold; color: #8e44ad; padding: 4px 0;"
        )
        layout.addWidget(header)

        hint = QLabel("在模板 {{__BODY__}} 位置插入动态内容，增删章节无需修改模板")
        hint.setStyleSheet("font-size: 11px; color: #888; padding: 2px 0 6px 0;")
        layout.addWidget(hint)

        # 主体：左右分栏
        body = QHBoxLayout()
        body.setSpacing(8)

        # ── 左侧：章节列表 ──
        left = QVBoxLayout()
        left.setSpacing(6)

        left_label = QLabel("章节列表")
        left_label.setStyleSheet("font-size: 12px; font-weight: bold; color: #2c3e50;")
        left.addWidget(left_label)

        self.section_list = QListWidget()
        self.section_list.setObjectName("sectionList")
        self.section_list.setMinimumWidth(200)
        self.section_list.setMaximumWidth(300)
        self.section_list.currentRowChanged.connect(self._on_section_selected)
        left.addWidget(self.section_list)

        # 章节操作按钮
        section_btns = QHBoxLayout()
        section_btns.setSpacing(4)

        btn_add = QPushButton("+ 章节")
        btn_add.setStyleSheet(BTN_GREEN)
        btn_add.clicked.connect(self._add_section)
        section_btns.addWidget(btn_add)

        btn_del = QPushButton("- 删除")
        btn_del.setStyleSheet(BTN_RED)
        btn_del.clicked.connect(self._remove_section)
        section_btns.addWidget(btn_del)

        section_btns.addStretch()

        btn_up = QPushButton("↑")
        btn_up.setStyleSheet(BTN_GRAY)
        btn_up.setMaximumWidth(40)
        btn_up.clicked.connect(self._move_section_up)
        section_btns.addWidget(btn_up)

        btn_down = QPushButton("↓")
        btn_down.setStyleSheet(BTN_GRAY)
        btn_down.setMaximumWidth(40)
        btn_down.clicked.connect(self._move_section_down)
        section_btns.addWidget(btn_down)

        left.addLayout(section_btns)
        body.addLayout(left)

        # ── 右侧：子项编辑器 ──
        right = QVBoxLayout()
        right.setSpacing(6)

        right_label = QLabel("章节内容")
        right_label.setStyleSheet("font-size: 12px; font-weight: bold; color: #2c3e50;")
        right.addWidget(right_label)

        # 章节标题编辑
        title_row = QHBoxLayout()
        title_label = QLabel("章节标题：")
        title_label.setStyleSheet("font-size: 12px; color: #2c3e50;")
        title_row.addWidget(title_label)

        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("例如：第一章 项目概述")
        self.title_edit.setStyleSheet("font-size: 13px;")
        self.title_edit.textChanged.connect(self._on_title_changed)
        title_row.addWidget(self.title_edit, 1)
        right.addLayout(title_row)

        # 子项操作栏
        block_btns = QHBoxLayout()
        block_btns.setSpacing(4)

        btn_add_heading = QPushButton("📝 子标题")
        btn_add_heading.setStyleSheet(BTN_BLUE)
        btn_add_heading.setToolTip("添加二级/三级标题")
        btn_add_heading.clicked.connect(self._add_heading_block)
        block_btns.addWidget(btn_add_heading)

        btn_add_text = QPushButton("📄 段落")
        btn_add_text.setStyleSheet(BTN_BLUE)
        btn_add_text.setToolTip("添加富文本段落")
        btn_add_text.clicked.connect(self._add_text_block)
        block_btns.addWidget(btn_add_text)

        btn_add_table = QPushButton("📋 表格")
        btn_add_table.setStyleSheet(BTN_BLUE)
        btn_add_table.setToolTip("添加表格（行列可调，单元格直接编辑）")
        btn_add_table.clicked.connect(self._add_table_block)
        block_btns.addWidget(btn_add_table)

        btn_add_image = QPushButton("🖼 图片")
        btn_add_image.setStyleSheet(BTN_BLUE)
        btn_add_image.setToolTip("插入本地图片（可设宽度/图注）")
        btn_add_image.clicked.connect(self._add_image_block)
        block_btns.addWidget(btn_add_image)

        block_btns.addStretch()

        btn_block_up = QPushButton("↑ 上移")
        btn_block_up.setStyleSheet(BTN_GRAY)
        btn_block_up.clicked.connect(lambda: self._move_selected_block(-1))
        block_btns.addWidget(btn_block_up)

        btn_block_down = QPushButton("↓ 下移")
        btn_block_down.setStyleSheet(BTN_GRAY)
        btn_block_down.clicked.connect(lambda: self._move_selected_block(1))
        block_btns.addWidget(btn_block_down)

        right.addLayout(block_btns)

        # 子项列表（可滚动）
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")

        self.blocks_container = QWidget()
        self.blocks_layout = QVBoxLayout(self.blocks_container)
        self.blocks_layout.setSpacing(8)
        self.blocks_layout.setContentsMargins(4, 4, 4, 4)
        self.blocks_layout.addStretch()

        self.scroll.setWidget(self.blocks_container)
        right.addWidget(self.scroll, 1)

        # 空状态提示
        self.empty_hint = QLabel("选择或添加一个章节开始编辑")
        self.empty_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_hint.setStyleSheet("color: #aaa; font-size: 13px; padding: 40px;")
        right.addWidget(self.empty_hint)

        body.addLayout(right, 1)
        layout.addLayout(body, 1)

        self._update_ui_state()

    # ── 章节操作 ────────────────────────────────────────────

    def _add_section(self):
        """添加新章节"""
        num = len(self._sections) + 1
        section = {
            'title': f'第{_cn_num(num)}章',
            'blocks': [],
        }
        self._sections.append(section)
        self._refresh_section_list()
        self.section_list.setCurrentRow(len(self._sections) - 1)
        self.content_changed.emit()

    def _remove_section(self):
        """删除当前选中的章节"""
        if self._current_index < 0 or self._current_index >= len(self._sections):
            return
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定删除「{self._sections[self._current_index]['title']}」及其所有子项？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self._sections.pop(self._current_index)
        self._current_index = -1
        self._refresh_section_list()
        self._refresh_blocks()
        self._update_ui_state()
        self.content_changed.emit()

    def _move_section_up(self):
        idx = self._current_index
        if idx > 0:
            self._sections[idx], self._sections[idx - 1] = self._sections[idx - 1], self._sections[idx]
            self._refresh_section_list()
            self.section_list.setCurrentRow(idx - 1)
            self.content_changed.emit()

    def _move_section_down(self):
        idx = self._current_index
        if 0 <= idx < len(self._sections) - 1:
            self._sections[idx], self._sections[idx + 1] = self._sections[idx + 1], self._sections[idx]
            self._refresh_section_list()
            self.section_list.setCurrentRow(idx + 1)
            self.content_changed.emit()

    def _on_section_selected(self, row):
        self._current_index = row
        self._refresh_blocks()
        self._update_ui_state()

    def _on_title_changed(self):
        """章节标题编辑"""
        if self._current_index < 0 or self._current_index >= len(self._sections):
            return
        text = self.title_edit.text()
        self._sections[self._current_index]['title'] = text
        item = self.section_list.item(self._current_index)
        if item:
            item.setText(text)
        self.content_changed.emit()

    # ── 子项操作 ────────────────────────────────────────────

    def _add_heading_block(self):
        if self._current_index < 0:
            QMessageBox.information(self, "提示", "请先选择或添加一个章节。")
            return
        block = {'type': 'heading', 'level': 2, 'text': ''}
        self._sections[self._current_index]['blocks'].append(block)
        self._refresh_blocks()
        self.content_changed.emit()

    def _add_text_block(self):
        if self._current_index < 0:
            QMessageBox.information(self, "提示", "请先选择或添加一个章节。")
            return
        block = {'type': 'text', 'html': ''}
        self._sections[self._current_index]['blocks'].append(block)
        self._refresh_blocks()
        self.content_changed.emit()

    def _add_table_block(self):
        if self._current_index < 0:
            QMessageBox.information(self, "提示", "请先选择或添加一个章节。")
            return
        block = {
            'type': 'table',
            'header': True,
            'data': [['', '', ''], ['', '', ''], ['', '', '']],
        }
        self._sections[self._current_index]['blocks'].append(block)
        self._refresh_blocks()
        self.content_changed.emit()

    def _add_image_block(self):
        if self._current_index < 0:
            QMessageBox.information(self, "提示", "请先选择或添加一个章节。")
            return
        # 添加时直接弹出选图框；取消则不添加
        path, _ = QFileDialog.getOpenFileName(
            self, "选择图片", "",
            "图片文件 (*.png *.jpg *.jpeg *.bmp *.gif *.webp);;所有文件 (*)")
        if not path:
            return
        block = {'type': 'image', 'path': path, 'width': 0, 'caption': ''}
        self._sections[self._current_index]['blocks'].append(block)
        self._refresh_blocks()
        self.content_changed.emit()

    def _remove_block(self, block_index):
        if self._current_index < 0:
            return
        blocks = self._sections[self._current_index]['blocks']
        if 0 <= block_index < len(blocks):
            blocks.pop(block_index)
            self._refresh_blocks()
            self.content_changed.emit()

    def _move_selected_block(self, direction):
        """移动当前焦点所在的子项（-1=上移, +1=下移）"""
        if self._current_index < 0:
            return
        blocks = self._sections[self._current_index]['blocks']
        # 找到包含焦点的块（覆盖段落/表格/图片等所有块类型）
        focused_idx = -1
        focus_w = QApplication.focusWidget()
        if focus_w is not None:
            for i, w in enumerate(self._block_widgets):
                if w is focus_w or w.isAncestorOf(focus_w):
                    focused_idx = i
                    break
        if focused_idx < 0:
            # 没有焦点，移动最后一个块
            if len(blocks) > 0:
                focused_idx = len(blocks) - 1
        new_idx = focused_idx + direction
        if 0 <= new_idx < len(blocks):
            blocks[focused_idx], blocks[new_idx] = blocks[new_idx], blocks[focused_idx]
            self._refresh_blocks()
            self.content_changed.emit()

    # ── 刷新 UI ──────────────────────────────────────────────

    def _refresh_section_list(self):
        self.section_list.clear()
        for i, section in enumerate(self._sections):
            item = QListWidgetItem(section.get('title', f'第{i+1}章'))
            self.section_list.addItem(item)
        if self._current_index >= len(self._sections):
            self._current_index = len(self._sections) - 1 if self._sections else -1

    def _refresh_blocks(self):
        """重建子项编辑区"""
        # 清空旧 widget
        for w in self._block_widgets:
            w.deleteLater()
        self._block_widgets.clear()

        # 清空 layout（保留末尾的 stretch）
        while self.blocks_layout.count() > 1:
            item = self.blocks_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if self._current_index < 0 or self._current_index >= len(self._sections):
            self.empty_hint.show()
            self.title_edit.blockSignals(True)
            self.title_edit.clear()
            self.title_edit.blockSignals(False)
            self.title_edit.setEnabled(False)
            return

        self.empty_hint.hide()
        self.title_edit.setEnabled(True)

        section = self._sections[self._current_index]
        self.title_edit.blockSignals(True)
        self.title_edit.setText(section.get('title', ''))
        self.title_edit.blockSignals(False)

        for i, block in enumerate(section.get('blocks', [])):
            widget = self._create_block_widget(i, block)
            self.blocks_layout.insertWidget(self.blocks_layout.count() - 1, widget)
            self._block_widgets.append(widget)

    def _create_block_widget(self, index, block):
        """创建单个子项的编辑 widget"""
        on_change = self.content_changed.emit
        on_delete = lambda idx=index: self._remove_block(idx)

        if block['type'] == 'heading':
            return HeadingBlockWidget(block, on_change, on_delete, index)
        elif block['type'] == 'table':
            return TableBlockWidget(block, on_change, on_delete, index)
        elif block['type'] == 'image':
            return ImageBlockWidget(block, on_change, on_delete, index)
        else:
            return TextBlockWidget(block, on_change, on_delete, index)

    # ── 状态管理 ────────────────────────────────────────────

    def _update_ui_state(self):
        has_section = self._current_index >= 0
        self.title_edit.setEnabled(has_section)
        if not has_section:
            self.empty_hint.show()

    # ── 数据接口 ────────────────────────────────────────────

    def get_sections(self):
        """获取当前所有章节数据（用于生成文档）"""
        return self._sections

    def set_sections(self, sections):
        """从外部设置章节数据（用于加载项目）— 深拷贝防止数据别名"""
        self._sections = copy.deepcopy(sections) if sections else []
        self._current_index = -1
        self._refresh_section_list()
        self._refresh_blocks()
        self._update_ui_state()

    def clear(self):
        """清空所有章节"""
        self._sections = []
        self._current_index = -1
        self._refresh_section_list()
        self._refresh_blocks()
        self._update_ui_state()


def _cn_num(n):
    """数字转中文（简单版）"""
    cn = '零一二三四五六七八九十'
    if n <= 10:
        return cn[n]
    elif n < 20:
        return '十' + (cn[n - 10] if n > 10 else '')
    else:
        tens, ones = divmod(n, 10)
        result = cn[tens] + '十'
        if ones:
            result += cn[ones]
        return result
