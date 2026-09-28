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
    QStackedWidget, QListWidget, QListWidgetItem, QToolButton, QMenu,
    QRadioButton, QButtonGroup, QCheckBox, QDialog, QSystemTrayIcon,
)
from PySide6.QtCore import Qt, QEvent
from PySide6.QtGui import QFont, QIcon, QAction, QActionGroup

try:
    from version import VERSION       # 打包后 version.py 随 datas 落在 _internal/ 下
except ImportError:                   # 极端情况读不到：不显示版本号，其余功能不受影响
    VERSION = None
from doc_filler import parse_tags, fill_template, create_sample_template, has_body_marker
from field_store import load_state, save_state, clear_state, has_state, state_summary
from field_widgets import AutoGrowTextEdit
from app_settings import (
    CLOSE_ACTIONS, DEFAULT_CLOSE_ACTION,
    get_close_action, set_close_action, get_setting, set_setting,
)


def resource_path(relative_path):
    """获取资源文件的绝对路径（兼容 PyInstaller 打包和直接运行）"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)
from xlsx_filler import (
    parse_tags as parse_xlsx_tags,
    fill_template as fill_xlsx_template,
    create_sample_template as create_xlsx_sample,
)
from batch_manager import BatchManager
from section_editor import SectionEditor
from labels import tag_to_label, load_tag_labels, get_label, get_type, is_long_text_tag

# ── 样式（极简风格）──
ACCENT = "#4a6fa5"          # 唯一强调色
BORDER = "#e3e3e0"
TEXT_MAIN = "#2c2c2a"
TEXT_SUB = "#75756f"
TEXT_HINT = "#9a9a94"

MAIN_STYLE = """
QMainWindow, QWidget { background: #fbfbfa; }
QLabel { color: #2c2c2a; font-size: 13px; background: transparent; }
QLineEdit, QSpinBox, QComboBox {
    border: none; border-bottom: 1px solid #e3e3e0; border-radius: 0;
    padding: 5px 2px; font-size: 13px; color: #2c2c2a;
    background: transparent; selection-background-color: #dfe8f4;
}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus { border-bottom: 1px solid #4a6fa5; }
QLineEdit:disabled { color: #b5b5b0; }
QTextEdit {
    border: 1px solid #e3e3e0; border-radius: 3px;
    padding: 6px 8px; background: #fdfdfc; font-size: 13px;
}
QTextEdit:focus { border: 1px solid #4a6fa5; }
QScrollArea { border: none; background: transparent; }
QScrollBar:vertical { background: transparent; width: 8px; margin: 0; }
QScrollBar::handle:vertical { background: #d5d5d1; border-radius: 4px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #b9b9b4; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal { background: transparent; height: 8px; }
QScrollBar::handle:horizontal { background: #d5d5d1; border-radius: 4px; min-width: 30px; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QStatusBar { background: #fbfbfa; color: #9a9a94; font-size: 12px; }
QStatusBar::item { border: none; }
QToolTip { background: #ffffff; color: #2c2c2a; border: 1px solid #e3e3e0; padding: 4px 6px; }
"""

TOPBAR_STYLE = """
#TopBar { background: #fbfbfa; border-bottom: 1px solid #e3e3e0; }
#AppName { font-size: 14px; font-weight: 500; color: #2c2c2a; }
#TplName { font-size: 12px; color: #9a9a94; }
"""

NAV_STYLE = """
QListWidget#Nav {
    background: #f7f7f5; border: none; border-right: 1px solid #e3e3e0;
    outline: none; padding: 10px 0;
}
QListWidget#Nav::item {
    height: 38px; padding-left: 15px;
    border-left: 3px solid transparent;
    color: #6b6b66; font-size: 13px;
}
QListWidget#Nav::item:hover { background: #efefec; }
QListWidget#Nav::item:selected {
    background: #eceff5; color: #2c2c2a;
    border-left: 3px solid #4a6fa5; font-weight: 500;
}
"""

PAGE_TITLE_STYLE = "font-size: 14px; font-weight: 500; color: #2c2c2a;"
PAGE_HINT_STYLE = "font-size: 12px; color: #9a9a94;"
SECTION_SEP_STYLE = "color: #ecece9;"

# 主操作按钮（生成）
BUTTON_PRIMARY = """
QPushButton {
    background: #4a6fa5; color: #ffffff; border: none;
    border-radius: 4px; padding: 7px 20px; font-size: 13px;
}
QPushButton:hover { background: #3f608f; }
QPushButton:disabled { background: #c3cbd6; }
"""

# 描边按钮（次要动作）
BUTTON_SECONDARY = """
QPushButton {
    background: transparent; color: #4a4a46; border: 1px solid #d8d8d4;
    border-radius: 4px; padding: 6px 14px; font-size: 13px;
}
QPushButton:hover { background: #f2f2ef; border-color: #c3c3bf; }
QPushButton:disabled { color: #b5b5b0; border-color: #e8e8e5; }
"""

# 无边框按钮（工具/菜单）
BUTTON_GHOST = """
QPushButton {
    background: transparent; color: #75756f; border: none;
    border-radius: 4px; padding: 6px 10px; font-size: 13px;
}
QPushButton:hover { background: #f2f2ef; color: #2c2c2a; }
"""

BUTTON_PROJECT = BUTTON_SECONDARY    # 兼容旧调用名
BUTTON_SUCCESS = BUTTON_PRIMARY      # 兼容旧调用名

TEMPLATE_LABEL_STYLE = "#TplName { font-size: 12px; color: #9a9a94; }"
FIELD_LABEL_STYLE = "font-size: 13px; color: #6b6b66;"


class CloseChoiceDialog(QDialog):
    """点窗口「X」时的三选一 —— 最小化到托盘 / 退出 TemplateFill / 取消

    只在「关闭窗口时 = 每次询问」（出厂默认）下弹出。勾选「记住我的选择」后写入
    `~/.TemplateFill/settings.json`，之后点 X 直接照办、不再打扰；想改回来走
    **托盘右键菜单 → 关闭窗口时**（唯一的入口）。

    结果读 `self.choice` 而不是 exec() 返回值 —— 点 X / Esc 关掉对话框时 exec()
    也返回 0，但语义是「取消」，两者不能混为一谈。

    ⚠ 离屏（offscreen）下 exec() 会永久阻塞且不报错，测试切勿触发本对话框；
    请直接调用 `TemplateFillWindow._resolve_close_action()` / `_minimize_to_tray()`。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关闭 TemplateFill")
        self.setMinimumWidth(420)

        #: "tray" / "quit" / "cancel"（点 X 或 Esc 关掉时保持 cancel）
        self.choice = "cancel"
        self.remember = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 16)
        layout.setSpacing(12)

        tip = QLabel("要收进右下角托盘继续在后台运行，还是直接退出 TemplateFill？")
        tip.setWordWrap(True)
        layout.addWidget(tip)

        self.remember_box = QCheckBox("记住我的选择（可在托盘右键菜单「关闭窗口时」改回）")
        layout.addWidget(self.remember_box)

        row = QHBoxLayout()
        row.setSpacing(8)
        row.addStretch(1)
        buttons = (("最小化到托盘", "tray", BUTTON_PRIMARY),
                   ("直接退出", "quit", BUTTON_SECONDARY),
                   ("取消", "cancel", BUTTON_GHOST))
        for text, value, style in buttons:
            btn = QPushButton(text)
            btn.setStyleSheet(style)
            if value == "tray":
                btn.setDefault(True)       # 回车默认收托盘：误关时不会丢现场
            btn.clicked.connect(lambda _=False, v=value: self._choose(v))
            row.addWidget(btn)
        layout.addLayout(row)

    def _choose(self, value):
        self.choice = value
        self.remember = self.remember_box.isChecked()
        self.accept()


class TemplateFillWindow(QMainWindow):
    """主窗口"""

    # 默认模板存放目录
    TEMPLATES_DIR = os.path.join(os.path.expanduser("~"), "TemplateFill", "templates")

    def __init__(self):
        super().__init__()
        self.setWindowTitle("TemplateFill — 模板文档生成器")
        self.setWindowIcon(QIcon(resource_path("TemplateFill.ico")))
        self.setMinimumSize(900, 650)
        self.resize(1000, 700)

        self._template_path = None
        self._template_format = None      # 'docx' or 'xlsx'
        self._project_path = None         # 当前项目文件路径（.tplfill）
        self._last_dir = os.path.expanduser("~\\Desktop")  # 上次打开/保存的目录
        self._field_widgets = {}          # tag_name → (QLineEdit | QTextEdit)
        self._loop_tables = {}            # loop_var → QTableWidget
        self._loop_fields = {}            # loop_var → [field_names]
        self._custom_labels = {}          # user-edited labels (in-app, highest priority)
        self._custom_types = {}           # user-edited widget types (in-app)
        self._json_labels = {}            # labels from .labels.json file (parsed entries)
        self._has_body_marker = False     # 模板是否含 {{__BODY__}} 动态内容标记
        self._section_editor = None       # SectionEditor 实例（仅有 body marker 时创建）
        self._last_output = None          # 最近一次生成的文档路径
        self._field_assign = {}           # tag → 分组名（用户自定义归组）
        self._custom_groups = []          # 自定义分组名（有序）
        self._custom_fields = {}          # tag → {'label','type'}（用户手工新增的字段）
        self._removed_fields = []         # 已从界面隐藏的字段
        self._field_order = []            # 字段显示顺序（空 = 模板解析顺序）
        self._template_tags = []          # 模板解析出的简单变量（不含手工新增）
        self._parsed_tags = set()         # 模板里真实存在的全部变量（提示用）
        self._label_width = self.FIELD_LABEL_MIN_WIDTH
        self._restored_state = {}         # 本次加载恢复的状态（供提示用）

        self.batch = BatchManager(self)   # 批量填充管理器

        # 托盘 / 关闭行为
        self._tray = None                 # QSystemTrayIcon；系统托盘不可用时为 None
        self._tray_menu = None            # 托盘右键菜单（必须长期持有，否则被回收即失效）
        self._close_group = None          # 「关闭窗口时」的单选组
        self._close_actions = {}          # key → QAction
        self._pre_tray_state = None       # 收进托盘前的窗口状态（最大化要还原）
        self._tray_tip_shown = False      # 每次运行只提示一次「程序还在后台」
        self._really_quit = False         # True = 本次关闭是真退出，不再走托盘逻辑

        self._setup_ui()
        self._setup_tray()
        self._auto_load_sample()

    # ── UI 构建 ────────────────────────────────────────────

    # 左侧导航：字段页 + 三个固定页
    INFO_GROUP = "文档信息"
    FIELDS_GROUP = "变量字段"
    NAV_TITLES = ["字段", "循环表格", "动态正文", "生成导出"]
    # 页面骨架：字段页承载全部字段，分组降级为页内小节标题
    HEAD_PAGES = [("fields", "字段")]
    TAIL_PAGES = [("loops", "循环表格"), ("body", "动态正文"), ("export", "生成导出")]
    # 关闭窗口时的行为 key → 显示名（存 ~/.TemplateFill/settings.json）
    CLOSE_LABELS = CLOSE_ACTIONS
    # 默认归入「文档信息」小节的常用字段
    PRIORITY_TAGS = ['title', 'project_name', 'department', 'author',
                     'date', 'doc_number', 'company', 'reviewer', 'phone']
    # 字段行布局：左标题列宽（按页内最长名自适应）与输入框最大宽度
    FIELD_LABEL_MIN_WIDTH = 96
    FIELD_LABEL_MAX_WIDTH = 168
    FIELD_MAX_WIDTH = 760
    FIELD_MAX_LINES = 8

    def _setup_ui(self):
        self.setStyleSheet(MAIN_STYLE)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(0)
        root.setContentsMargins(0, 0, 0, 0)

        # ── 顶栏：应用名 · 模板名 · 主操作 ──
        top = QWidget()
        top.setObjectName("TopBar")
        top.setStyleSheet(TOPBAR_STYLE)
        top.setFixedHeight(52)
        row = QHBoxLayout(top)
        # 上下留边距：控件不再贴满 52px 高度，避免盖住底部分割线，按钮也不会被拉成通栏高
        row.setContentsMargins(20, 9, 20, 10)
        row.setSpacing(10)

        app_name = QLabel("TemplateFill")
        app_name.setObjectName("AppName")
        row.addWidget(app_name)

        self.lbl_template = QLabel("未加载模板")
        self.lbl_template.setObjectName("TplName")
        row.addWidget(self.lbl_template)
        row.addStretch()

        self.btn_open = QPushButton("打开模板")
        self.btn_open.setStyleSheet(BUTTON_SECONDARY)
        self.btn_open.clicked.connect(self._on_open_template)
        row.addWidget(self.btn_open)

        self.btn_batch = QPushButton("批量模式")
        self.btn_batch.setStyleSheet(BUTTON_SECONDARY)
        self.btn_batch.setToolTip("开启后每行数据生成一个独立文件")
        self.btn_batch.clicked.connect(self.batch.toggle)
        row.addWidget(self.btn_batch)

        # 更多菜单：示例模板 / 项目文件
        self.btn_more = QToolButton()
        self.btn_more.setText("⋯")
        self.btn_more.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.btn_more.setStyleSheet("""
            QToolButton { background: transparent; color: #75756f; border: none;
                         border-radius: 4px; padding: 4px 10px; font-size: 15px; }
            QToolButton:hover { background: #f2f2ef; color: #2c2c2a; }
            QToolButton::menu-indicator { image: none; }
        """)
        more_menu = QMenu(self)
        act_word = more_menu.addAction("生成 Word 示例模板")
        act_word.triggered.connect(self._on_create_sample)
        act_excel = more_menu.addAction("生成 Excel 示例模板")
        act_excel.triggered.connect(self._on_create_xlsx_sample)
        more_menu.addSeparator()
        act_open_prj = more_menu.addAction("打开项目...")
        act_open_prj.triggered.connect(self._open_project)
        act_save_prj = more_menu.addAction("保存项目...")
        act_save_prj.triggered.connect(self._save_project)
        more_menu.addSeparator()
        act_clear_vals = more_menu.addAction("清空已填内容")
        act_clear_vals.setToolTip("只清空输入框内容，保留字段分组与显示名设置")
        act_clear_vals.triggered.connect(self._on_clear_values)
        act_forget = more_menu.addAction("清除本模板记忆")
        act_forget.setToolTip("删除本模板的填写记忆（含分组与显示名设置）")
        act_forget.triggered.connect(self._on_forget_template)
        self.btn_more.setMenu(more_menu)
        row.addWidget(self.btn_more)

        self.btn_generate = QPushButton("生成文档")
        self.btn_generate.setStyleSheet(BUTTON_PRIMARY)
        self.btn_generate.setEnabled(False)
        self.btn_generate.clicked.connect(self._on_generate)
        row.addWidget(self.btn_generate)

        root.addWidget(top)

        # ── 主体：左侧导航 + 右侧分区页 ──
        body = QHBoxLayout()
        body.setSpacing(0)
        body.setContentsMargins(0, 0, 0, 0)

        self.nav = QListWidget()
        self.nav.setObjectName("Nav")
        self.nav.setStyleSheet(NAV_STYLE)
        self.nav.setFixedWidth(178)
        self.nav.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.nav.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.nav.currentRowChanged.connect(self._on_nav_changed)
        body.addWidget(self.nav)

        self.stack = QStackedWidget()
        body.addWidget(self.stack, 1)

        # 按当前分组配置生成页面骨架（无自定义分组时为 5 页）
        self._rebuild_pages()

        root.addLayout(body, 1)

        # ── 状态栏 ──
        self.status_bar = QStatusBar()
        self.status_bar.setSizeGripEnabled(False)
        self.status_bar.showMessage("点击「打开模板」开始 — 支持 .docx 与 .xlsx")
        self.setStatusBar(self.status_bar)

        # 初始空状态
        self.nav.setCurrentRow(0)
        self._render_empty_state()

    # ── 分区页辅助 ──────────────────────────────────────────

    def _page_spec(self):
        """页面清单：[(页标识, 标题)]。

        字段不再按分组分页 —— 分组降级为「字段」页内的小节标题，
        这样模板只有几个字段时也不会出现一堆空导航页。
        """
        return list(self.HEAD_PAGES) + list(self.TAIL_PAGES)

    def _rebuild_pages(self):
        """按当前配置重建导航与页面堆栈。

        页面按索引 0..n 排列，与 nav 项一一对应。
        """
        spec = self._page_spec()

        self._page = {}                # 页标识 → 索引
        self._page_titles = {}         # 页标识 → 标题
        self._page_containers = []     # 每页滚动内容容器
        self._page_layouts = []        # 每页内容布局（挂控件用）
        self._page_field_tags = {}     # 页标识 → 该页字段列表

        self.nav.blockSignals(True)
        self.nav.clear()
        while self.stack.count():
            old = self.stack.widget(0)
            self.stack.removeWidget(old)
            old.deleteLater()
        self.nav.blockSignals(False)

        for index, (pid, title) in enumerate(spec):
            self.nav.addItem(QListWidgetItem(title))
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            container = QWidget()
            lay = QVBoxLayout(container)
            lay.setContentsMargins(28, 22, 28, 22)
            lay.setSpacing(12)
            scroll.setWidget(container)
            self.stack.addWidget(scroll)
            self._page[pid] = index
            self._page_titles[pid] = title
            self._page_containers.append(container)
            self._page_layouts.append(lay)

        # 兼容既有调用：form_layout 指向「变量字段」页（批量表格挂载点）
        fi = self._page["fields"]
        self.form_layout = self._page_layouts[fi]
        self.form_container = self._page_containers[fi]
        self.scroll_area = self.stack.widget(fi)

    def _page_index(self, key):
        """页标识或索引 → 索引（越界返回 0）"""
        if isinstance(key, int):
            return key
        return self._page.get(key, 0)

    def _goto_page(self, key):
        """切换分区页（导航与堆栈同步）；key 可为页标识或索引"""
        index = self._page_index(key)
        if 0 <= index < self.stack.count():
            self.nav.setCurrentRow(index)
            self.stack.setCurrentIndex(index)

    def _on_nav_changed(self, index):
        """左侧导航切换"""
        if 0 <= index < self.stack.count():
            self.stack.setCurrentIndex(index)
        if index == self._page.get("export"):
            self._update_summary()

    def _on_section_context_menu(self, name, widget, pos):
        """右键字段页内的小节标题（自定义分组）→ 重命名 / 删除分组"""
        menu = QMenu(self)
        act_rename = menu.addAction("重命名分组...")
        act_delete = menu.addAction("删除分组")
        chosen = menu.exec(widget.mapToGlobal(pos))
        if chosen is act_rename:
            self._rename_group(name)
        elif chosen is act_delete:
            self._delete_group(name)

    def _clear_page(self, key):
        """清空指定分区页的内容布局"""
        index = self._page_index(key)
        if not (0 <= index < len(self._page_layouts)):
            return
        lay = self._page_layouts[index]
        while lay.count():
            item = lay.takeAt(0)
            w = item.widget()
            if w:
                w.setParent(None)
                w.deleteLater()
            elif item.layout():
                self._clear_layout(item.layout())

    def _clear_layout(self, lay):
        while lay.count():
            item = lay.takeAt(0)
            w = item.widget()
            if w:
                w.setParent(None)
                w.deleteLater()
            elif item.layout():
                self._clear_layout(item.layout())

    def _page_header(self, key, title, hint=""):
        """分区页头部：标题 + 说明 + 细分隔线"""
        lay = self._page_layouts[self._page_index(key)]
        head = QHBoxLayout()
        head.setSpacing(10)
        lbl = QLabel(title)
        lbl.setStyleSheet(PAGE_TITLE_STYLE)
        head.addWidget(lbl)
        if hint:
            h = QLabel(hint)
            h.setStyleSheet(PAGE_HINT_STYLE)
            head.addWidget(h)
        head.addStretch()
        lay.addLayout(head)

    def _empty_hint(self, key, text):
        """分区页空状态提示"""
        lay = self._page_layouts[self._page_index(key)]
        lbl = QLabel(text)
        lbl.setStyleSheet("color: #9a9a94; font-size: 13px; padding: 26px 0;")
        lbl.setWordWrap(True)
        lay.addWidget(lbl)

    def _render_empty_state(self):
        """未加载模板时的占位内容"""
        for pid, title in self._page_spec():
            self._clear_page(pid)
            self._page_header(pid, title)
            if pid == "fields":
                self._empty_hint(pid, "尚未加载模板。\n\n点击右上角「打开模板」选择 .docx 或 .xlsx 文件，"
                                      "或在「⋯」菜单里生成示例模板先试用。")
            else:
                self._empty_hint(pid, "尚未加载模板。")
        self._page_field_tags = {}
        self._refresh_nav_counts({})

    def _refresh_nav_counts(self, counts=None):
        """刷新导航项计数。counts: {页标识: 数量}，缺省按当前表单统计"""
        if counts is None:
            counts = {}
            for pid, tags in self._page_field_tags.items():
                counts[pid] = len(tags)
            counts["loops"] = len(self._loop_tables)
            counts["body"] = self._count_body_blocks()
        for pid, index in self._page.items():
            if not (0 <= index < self.nav.count()):
                continue
            title = self._page_titles.get(pid, self.nav.item(index).text())
            n = counts.get(pid, 0)
            self.nav.item(index).setText(f"{title}  ·  {n}" if n else title)

    def _update_nav_counts(self, counts):
        """兼容旧调用：counts 依次对应 字段 / 循环表格 / 动态正文"""
        order = ["fields", "loops", "body"]
        self._refresh_nav_counts({pid: counts[i] for i, pid in enumerate(order)
                                  if i < len(counts)})

    def _on_sections_changed(self):
        """章节编辑器内容变化 → 刷新状态栏与导航计数"""
        self.status_bar.showMessage("章节内容已更新", 2000)
        self._refresh_nav_counts()

    # ── 页 3：动态正文 ─────────────────────────────────────

    def _build_body_page(self):
        """动态正文页（仅 Word 模板含 {{__BODY__}} 时可用）"""
        self._section_editor = None
        if self.batch.mode:
            self._page_header("body", "动态正文", "批量模式不可用")
            self._empty_hint("body", "批量模式下不编辑动态正文，正文内容按模板原样输出。")
            return
        if not self._has_body_marker:
            self._page_header("body", "动态正文", "模板未启用")
            self._empty_hint("body", "当前模板没有 {{__BODY__}} 标记。\n\n"
                                     "把 {{__BODY__}} 写进 Word 模板的正文位置，"
                                     "就能在这里增删章节、自由编排正文，无需改动模板本身；"
                                     "封面、页眉页脚仍按模板原样输出。")
            return
        self._page_header("body", "动态正文",
                          "增删章节与子项 · 生成时插入到 {{__BODY__}} 所在位置")
        self._section_editor = SectionEditor()
        self._section_editor.setMinimumHeight(560)
        self._section_editor.content_changed.connect(self._on_sections_changed)
        self._page_layouts[self._page_index("body")].addWidget(self._section_editor, 1)

    # ── 生成导出页 ─────────────────────────────────────────

    def _build_output_page(self):
        """输出设置 + 生成后动作 + 内容摘要 + 最近输出"""
        self._page_header("export", "生成导出", "设置保存位置与生成后的动作")
        lay = self._page_layouts[self._page_index("export")]

        grid_host = QWidget()
        grid = QGridLayout(grid_host)
        grid.setContentsMargins(0, 4, 0, 0)
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(12)
        grid.setColumnStretch(1, 1)

        lbl_dir = QLabel("保存位置")
        lbl_dir.setFixedWidth(96)
        lbl_dir.setStyleSheet(FIELD_LABEL_STYLE)
        grid.addWidget(lbl_dir, 0, 0)

        self.out_dir_edit = QLineEdit()
        self.out_dir_edit.setPlaceholderText(f"留空用默认：{self._default_output_dir()}")
        grid.addWidget(self.out_dir_edit, 0, 1)

        btn_browse = QPushButton("选择目录")
        btn_browse.setStyleSheet(BUTTON_SECONDARY)
        btn_browse.clicked.connect(self._choose_output_dir)
        grid.addWidget(btn_browse, 0, 2)

        lbl_name = QLabel("文件名称")
        lbl_name.setFixedWidth(96)
        lbl_name.setStyleSheet(FIELD_LABEL_STYLE)
        grid.addWidget(lbl_name, 1, 0)

        self.out_name_edit = QLineEdit()
        self.out_name_edit.setPlaceholderText("留空则用模板名")
        self.out_name_edit.textChanged.connect(self._update_filename_preview)
        grid.addWidget(self.out_name_edit, 1, 1, 1, 2)

        self.lbl_filename_preview = QLabel()
        self.lbl_filename_preview.setStyleSheet(PAGE_HINT_STYLE)
        grid.addWidget(self.lbl_filename_preview, 2, 1, 1, 2)
        lay.addWidget(grid_host)

        act_title = QLabel("生成后")
        act_title.setStyleSheet("font-size: 13px; color: #6b6b66; margin-top: 6px;")
        lay.addWidget(act_title)

        act_row = QHBoxLayout()
        act_row.setSpacing(22)
        self._action_group = QButtonGroup(self)
        for i, (text, key) in enumerate([("仅保存", 'save'),
                                         ("打开文件", 'open_file'),
                                         ("打开所在文件夹", 'open_dir')]):
            rb = QRadioButton(text)
            rb.setStyleSheet("QRadioButton { font-size: 13px; color: #4a4a46; }")
            rb.setProperty("action_key", key)
            self._action_group.addButton(rb, i)
            act_row.addWidget(rb)
            if i == 0:
                rb.setChecked(True)
        act_row.addStretch()
        lay.addLayout(act_row)

        sum_title = QLabel("内容摘要")
        sum_title.setStyleSheet("font-size: 13px; color: #6b6b66; margin-top: 10px;")
        lay.addWidget(sum_title)

        self.lbl_summary = QLabel("—")
        self.lbl_summary.setStyleSheet("font-size: 13px; color: #4a4a46;")
        self.lbl_summary.setWordWrap(True)
        lay.addWidget(self.lbl_summary)

        btn_row = QHBoxLayout()
        self.btn_gen_page = QPushButton("生成文档")
        self.btn_gen_page.setStyleSheet(BUTTON_PRIMARY)
        self.btn_gen_page.setMinimumHeight(34)
        self.btn_gen_page.clicked.connect(self._on_generate)
        btn_row.addWidget(self.btn_gen_page)
        btn_row.addStretch()
        lay.addLayout(btn_row)

        recent_title = QLabel("最近输出")
        recent_title.setStyleSheet("font-size: 13px; color: #6b6b66; margin-top: 10px;")
        lay.addWidget(recent_title)

        recent_box = QWidget()
        recent_lay = QHBoxLayout(recent_box)
        recent_lay.setContentsMargins(0, 0, 0, 0)
        recent_lay.setSpacing(10)
        self.lbl_recent = QLabel("尚未生成文档")
        self.lbl_recent.setStyleSheet(PAGE_HINT_STYLE)
        recent_lay.addWidget(self.lbl_recent, 1)

        self.btn_recent_open = QPushButton("打开文件")
        self.btn_recent_open.setStyleSheet(BUTTON_SECONDARY)
        self.btn_recent_open.setEnabled(False)
        self.btn_recent_open.clicked.connect(lambda: self._open_path(self._last_output))
        recent_lay.addWidget(self.btn_recent_open)

        self.btn_recent_dir = QPushButton("打开所在文件夹")
        self.btn_recent_dir.setStyleSheet(BUTTON_SECONDARY)
        self.btn_recent_dir.setEnabled(False)
        self.btn_recent_dir.clicked.connect(
            lambda: self._open_path(os.path.dirname(self._last_output) if self._last_output else ""))
        recent_lay.addWidget(self.btn_recent_dir)
        lay.addWidget(recent_box)

        lay.addStretch()

        self._update_filename_preview()
        self._update_summary()
        self._update_recent_output()

    def _update_summary(self):
        """刷新内容摘要"""
        if not hasattr(self, 'lbl_summary'):
            return

        if self.batch.mode and self.batch.table:
            t = self.batch.table
            self.lbl_summary.setText(
                f"批量数据　{t.rowCount()} 行 × {t.columnCount()} 列　·　"
                f"将生成 {t.rowCount()} 个文件")
            return

        total = len(self._field_widgets)
        filled = 0
        for widget in self._field_widgets.values():
            try:
                text = widget.toPlainText() if hasattr(widget, 'toPlainText') else widget.text()
            except Exception:
                text = ''
            if text and text.strip():
                filled += 1

        loop_rows = 0
        for table in self._loop_tables.values():
            for r in range(table.rowCount()):
                for c in range(table.columnCount()):
                    item = table.item(r, c)
                    if item and item.text().strip():
                        loop_rows += 1
                        break

        parts = [f"字段　{filled} / {total} 已填"]
        if self._loop_tables:
            parts.append(f"循环数据　{loop_rows} 行")
        if self._has_body_marker:
            parts.append(f"正文章节　{self._count_body_blocks()} 章")
        self.lbl_summary.setText("　·　".join(parts))

    def _update_filename_preview(self):
        """输出文件名预览"""
        if not hasattr(self, 'lbl_filename_preview'):
            return
        if not self._template_path:
            self.lbl_filename_preview.setText("加载模板后显示输出文件名")
            return
        stem = self.out_name_edit.text().strip() or \
            os.path.splitext(os.path.basename(self._template_path))[0] + "_已填写"
        ext = os.path.splitext(self._template_path)[1].lower()
        self.lbl_filename_preview.setText(f"输出文件：{stem}_日期时间{ext}")

    def _update_recent_output(self):
        """最近输出区状态"""
        if not hasattr(self, 'lbl_recent'):
            return
        if self._last_output and os.path.exists(self._last_output):
            self.lbl_recent.setText(os.path.basename(self._last_output))
            self.lbl_recent.setToolTip(self._last_output)
            self.lbl_recent.setStyleSheet("font-size: 13px; color: #2c2c2a;")
            self.btn_recent_open.setEnabled(True)
            self.btn_recent_dir.setEnabled(True)
        else:
            self.lbl_recent.setText("尚未生成文档")
            self.lbl_recent.setToolTip("")
            self.lbl_recent.setStyleSheet(PAGE_HINT_STYLE)
            self.btn_recent_open.setEnabled(False)
            self.btn_recent_dir.setEnabled(False)

    def _choose_output_dir(self):
        """选择输出目录"""
        start = self.out_dir_edit.text().strip() or self._default_output_dir()
        d = QFileDialog.getExistingDirectory(self, "选择保存位置", start)
        if d:
            self.out_dir_edit.setText(d)

    @staticmethod
    def _default_output_dir():
        """默认输出目录：桌面（不存在则用户主目录）"""
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        return desktop if os.path.isdir(desktop) else os.path.expanduser("~")

    def _output_action(self):
        """生成后动作：save / open_file / open_dir"""
        if hasattr(self, '_action_group'):
            btn = self._action_group.checkedButton()
            if btn:
                return btn.property("action_key") or 'save'
        return 'save'

    def _resolve_output_path(self):
        """按输出设置拼出完整路径；返回 None 表示走引擎默认（桌面 + 时间戳）"""
        if not hasattr(self, 'out_dir_edit'):
            return None
        out_dir = self.out_dir_edit.text().strip()
        prefix = self.out_name_edit.text().strip()
        if not out_dir and not prefix:
            return None
        base_dir = out_dir or self._default_output_dir()
        if not os.path.isdir(base_dir):
            try:
                os.makedirs(base_dir, exist_ok=True)
            except Exception:
                base_dir = self._default_output_dir()
        tpl_base = os.path.splitext(os.path.basename(self._template_path or 'output'))[0]
        stem = prefix or f"{tpl_base}_已填写"
        ext = os.path.splitext(self._template_path or '.docx')[1].lower() or '.docx'
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        return os.path.join(base_dir, f"{stem}_{ts}{ext}")

    @staticmethod
    def _open_path(path):
        """用系统默认程序打开文件或文件夹"""
        if not path or not os.path.exists(path):
            return
        try:
            os.startfile(path)
        except Exception:
            pass

    # ── 模板加载 ────────────────────────────────────────────

    def _auto_load_sample(self):
        """启动时检查是否有示例模板可用"""
        sample_path = os.path.join(self.TEMPLATES_DIR, "sample_report.docx")
        if os.path.exists(sample_path):
            self._load_template(sample_path)

    def _on_open_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择模板文件", self._last_dir,
            "Office 模板 (*.docx *.xlsx);;Word 文档 (*.docx);;Excel 工作簿 (*.xlsx);;所有文件 (*.*)"
        )
        if path:
            self._last_dir = os.path.dirname(path)
            self._load_template(path)

    def _on_create_sample(self):
        """生成示例 Word 模板并自动加载"""
        try:
            os.makedirs(self.TEMPLATES_DIR, exist_ok=True)
            sample_path = os.path.join(self.TEMPLATES_DIR, "sample_report.docx")
            create_sample_template(sample_path)
            QMessageBox.information(self, "成功",
                                    f"示例模板已生成：\n{sample_path}\n\n已自动加载，可以开始填写。")
            self._load_template(sample_path)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成示例模板失败：{e}")

    def _on_create_xlsx_sample(self):
        """生成示例 Excel 模板并自动加载"""
        try:
            os.makedirs(self.TEMPLATES_DIR, exist_ok=True)
            sample_path = os.path.join(self.TEMPLATES_DIR, "sample_quote.xlsx")
            create_xlsx_sample(sample_path)
            QMessageBox.information(self, "成功",
                                    f"Excel 示例模板已生成：\n{sample_path}\n\n"
                                    "包含简单变量和循环区域，已自动加载。")
            self._load_template(sample_path)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成 Excel 示例失败：{e}")

    def _load_template(self, path, restore_memory=True):
        """加载模板并生成表单，自动识别 .docx / .xlsx

        restore_memory=True 时，若该模板有填写记忆则自动带出上次内容。
        """
        try:
            # 切换模板前，先记住上一份模板的填写内容
            if self._template_path and \
                    os.path.abspath(self._template_path) != os.path.abspath(path):
                self._remember()

            ext = os.path.splitext(path)[1].lower()

            if ext == '.xlsx':
                self._template_format = 'xlsx'
                result = parse_xlsx_tags(path)
                tags = result['simple']
                loops = result['loops']
                self._json_labels = load_tag_labels(path)
                self._has_body_marker = False
            else:
                self._template_format = 'docx'
                tags = parse_tags(path)
                loops = {}
                self._json_labels = load_tag_labels(path)
                # 检测模板是否含 {{__BODY__}} 动态内容标记
                self._has_body_marker = has_body_marker(path)

            if not tags and not loops:
                QMessageBox.warning(self, "提示",
                                    "模板中未检测到 {{变量}} 占位符。\n\n"
                                    "请在模板中使用 {{变量名}} 格式插入占位符。")
                return

            total_fields = len(tags) + len(loops)
            fmt_name = "Excel" if self._template_format == 'xlsx' else "Word"
            self._template_path = path
            self._project_path = None
            self._section_editor = None
            self.batch.tags = tags
            self.batch.loops = loops
            # 模板里真实存在的变量（用于「模板无此变量」提示）
            self._parsed_tags = set(tags)
            for loop_info in (loops or {}).values():
                self._parsed_tags.update(loop_info.get('fields') or [])
            self.lbl_template.setText(
                f"·  {os.path.basename(path)}  ·  {fmt_name}  ·  {total_fields} 个字段")
            self.lbl_template.setStyleSheet("color: #9a9a94; font-size: 12px;")

            # 读取本模板的填写记忆（分组/显示名需在建设界面前应用）
            state = self._load_previous_state(path) if restore_memory else {}
            self._restored_state = state
            if state:
                self._apply_state_config(state)
            else:
                self._custom_labels = {}
                self._custom_types = {}
                self._custom_groups = []
                self._field_assign = {}
                self._custom_fields = {}
                self._removed_fields = []
                self._field_order = []

            self._rebuild_pages()
            self._build_form(tags, loops)

            restored = 0
            if state:
                restored = self._restore_state(state)
                self._refresh_nav_counts()
                self._update_summary()

            self.btn_generate.setEnabled(True)
            if restored:
                self.status_bar.showMessage(
                    f"已加载{fmt_name}模板并带出上次填写内容（{restored} 个字段）", 6000)
            else:
                self.status_bar.showMessage(f"已加载{fmt_name}模板：{path}")

            self._update_tray_tooltip()      # 托盘悬停提示带上模板名

        except Exception as e:
            QMessageBox.critical(self, "加载失败", str(e))

    # ── 表单构建 ────────────────────────────────────────────

    def _build_form(self, tags, loops=None):
        """构建输入表单。

        tags: 模板解析出的简单变量列表
        loops: 循环区域 dict {loop_var: {fields, item_var, ...}}

        页面：字段（全部字段单列纵向排列，分组降级为小节标题）/ 循环表格 / 动态正文 / 生成导出
        """
        self._field_widgets.clear()
        self._loop_tables.clear()
        self._loop_fields.clear()
        self.batch.table = None
        self._section_editor = None
        self._page_field_tags = {}
        self._template_tags = list(tags or [])
        for pid in list(self._page):
            self._clear_page(pid)

        # ── 批量模式：字段并入字段页的大表格 ──
        if self.batch.mode and tags:
            self._page_header("fields", "批量表格", f"{len(tags)} 列 · 每行数据生成一个独立文件")
            self.batch.build_form(tags, loops)
            if not loops:
                self._page_header("loops", "循环表格")
                self._empty_hint("loops", "该模板没有循环区域。")
            else:
                self._page_header("loops", "循环表格", "批量模式不支持")
                self._empty_hint("loops", "批量模式下不支持循环数据，生成时将跳过。")
            self._build_body_page()
            self._build_output_page()
            self._goto_page("fields")
            self._refresh_nav_counts({"fields": len(tags), "loops": len(loops or {}),
                                      "body": self._count_body_blocks()})
            return

        # ── 字段页（单列 + 小节标题） ──
        self._build_field_page(tags)

        # 循环表格 / 动态正文 / 生成导出
        self._build_loop_page(loops)
        self._build_body_page()
        self._build_output_page()

        self._goto_page("fields")
        self._refresh_nav_counts()

    # ── 字段页：单列排布 + 分组小节 ──────────────────────────

    def _visible_tags(self):
        """界面上实际显示的字段顺序（模板字段 + 手工新增 - 已隐藏，按自定义顺序）"""
        base = [t for t in self._template_tags if t not in self._removed_fields]
        for tag in self._custom_fields:
            if tag not in base and tag not in self._removed_fields:
                base.append(tag)
        if not self._field_order:
            return base
        ordered = [t for t in self._field_order if t in base]
        ordered += [t for t in base if t not in ordered]
        return ordered

    def _group_sections(self, tags):
        """把字段按分组切成小节：[(分组名, [tag, ...])]，空组不出现"""
        buckets = {}
        for tag in tags:
            buckets.setdefault(self._resolve_group(tag), []).append(tag)
        return [(name, buckets[name]) for name in self._group_names() if buckets.get(name)]

    def _label_column_width(self, tags):
        """左标题列宽：按页内最长显示名自适应，限制在 96~168px"""
        fm = self.fontMetrics()
        width = self.FIELD_LABEL_MIN_WIDTH
        for tag in tags:
            width = max(width, fm.horizontalAdvance(self._resolve_label(tag)) + 8)
        return min(width, self.FIELD_LABEL_MAX_WIDTH)

    def _build_field_page(self, tags):
        """字段页：小节标题 + 一行一个字段（左标题 / 右自动增高输入框）"""
        lay = self._page_layouts[self._page_index("fields")]
        all_tags = self._visible_tags()
        self._page_field_tags["fields"] = all_tags
        self._label_width = self._label_column_width(all_tags)

        custom_n = len(self._custom_fields)
        hint = f"{len(all_tags)} 个字段 · 双击标题改名 · 右键插入或删除字段"
        if custom_n:
            hint += f" · 含 {custom_n} 个自定义"
        self._page_header("fields", "字段", hint)

        if not all_tags:
            self._empty_hint("fields", "该模板没有可填字段。\n\n"
                                       "右键任意位置可从「⋯」菜单重新生成示例模板试用。")
            return

        for name, tag_list in self._group_sections(all_tags):
            self._build_section_header(name, len(tag_list))
            for tag in tag_list:
                self._build_field_row(tag)
        lay.addStretch()

    def _build_section_header(self, name, count):
        """字段页内的小节标题（分组）；自定义分组可右键改名/删除"""
        lay = self._page_layouts[self._page_index("fields")]
        wrap = QWidget()
        box = QVBoxLayout(wrap)
        box.setContentsMargins(0, 14, 0, 4)
        box.setSpacing(0)
        lbl = QLabel(f"{name}  ·  {count}")
        lbl.setStyleSheet("font-size: 12px; color: #9a9a94; background: transparent;")
        if name in self._custom_groups:
            lbl.setToolTip("右键可重命名或删除该分组")
            wrap.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            wrap.customContextMenuRequested.connect(
                lambda pos, n=name, w=wrap: self._on_section_context_menu(n, w, pos))
        else:
            lbl.setToolTip("内置分组：双击字段标题可把字段移到其它分组")
        box.addWidget(lbl)
        lay.addWidget(wrap)

    def _build_field_row(self, tag):
        """单个字段行：左标题 + 右输入框（自动增高、最大宽度 760px）"""
        lay = self._page_layouts[self._page_index("fields")]
        display_name = self._resolve_label(tag)
        kind = self._resolve_type(tag)

        row = QWidget()
        outer = QVBoxLayout(row)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(3)

        line = QHBoxLayout()
        line.setContentsMargins(0, 0, 0, 0)
        line.setSpacing(12)

        name = self._make_field_label(tag, display_name)
        name.setFixedWidth(getattr(self, "_label_width", self.FIELD_LABEL_MIN_WIDTH))
        name.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        line.addWidget(name)

        edit = AutoGrowTextEdit(min_lines=3 if kind == 'multi' else 1,
                                max_lines=self.FIELD_MAX_LINES)
        edit.setPlaceholderText(f"请输入{display_name}")
        edit.setMaximumWidth(self.FIELD_MAX_WIDTH)
        edit.setProperty("tag", tag)
        line.addWidget(edit, 1)

        outer.addLayout(line)

        # 变量在模板中不存在 → 标黄提醒（允许创建，但提示不会写入文档）
        if tag not in self._parsed_tags:
            warn = QLabel(f"模板里没有 {{{{{tag}}}}}，这里填写的内容不会写入文档")
            warn.setStyleSheet("font-size: 11px; color: #ba7517; background: transparent;")
            warn.setWordWrap(True)
            wl = QHBoxLayout()
            wl.setContentsMargins(getattr(self, "_label_width", self.FIELD_LABEL_MIN_WIDTH) + 12, 0, 0, 0)
            wl.addWidget(warn)
            wl.addStretch()
            outer.addLayout(wl)

        lay.addWidget(row)
        self._field_widgets[tag] = edit
        return row

    # ── 字段分组 ────────────────────────────────────────────

    def _resolve_group(self, tag):
        """字段所属分组：用户指定优先，否则按内置优先级判断"""
        group = self._field_assign.get(tag)
        if group in self._group_names():
            return group
        return self.INFO_GROUP if tag in self.PRIORITY_TAGS else self.FIELDS_GROUP

    def _group_names(self):
        """全部可用分组名（含两个内置分组）"""
        return [self.INFO_GROUP, self.FIELDS_GROUP] + list(self._custom_groups)

    def _bucket_fields(self, tags):
        """把字段按分组归页，返回 {页标识: [tag, ...]}"""
        buckets = {"info": [], "fields": []}
        for name in self._custom_groups:
            buckets[f"group:{name}"] = []
        for tag in tags:
            group = self._resolve_group(tag)
            if group == self.INFO_GROUP:
                pid = "info"
            elif group == self.FIELDS_GROUP:
                pid = "fields"
            else:
                pid = f"group:{group}"
            if pid not in buckets:      # 分组已不存在 → 回落到默认页
                pid = "info" if tag in self.PRIORITY_TAGS else "fields"
            buckets[pid].append(tag)
        return buckets

    def _set_field_group(self, tag, group):
        """把字段归入指定分组；group 为内置组名或自定义组名"""
        if group in (self.INFO_GROUP, self.FIELDS_GROUP):
            # 内置分组：只在与默认判断不一致时记录，避免状态文件冗余
            default = self.INFO_GROUP if tag in self.PRIORITY_TAGS else self.FIELDS_GROUP
            if group == default:
                self._field_assign.pop(tag, None)
            else:
                self._field_assign[tag] = group
        elif group in self._custom_groups:
            self._field_assign[tag] = group
        else:
            self._field_assign.pop(tag, None)

    def _new_group(self, name):
        """新建自定义分组（重名则忽略）"""
        name = (name or "").strip()
        if not name or name in self._group_names():
            return None
        self._custom_groups.append(name)
        return name

    def _rename_group(self, old_name):
        """重命名自定义分组"""
        new_name, ok = QInputDialog.getText(self, "重命名分组", "分组名称：", text=old_name)
        if not ok:
            return
        new_name = (new_name or "").strip()
        if not new_name or new_name == old_name or new_name in self._group_names():
            return
        self._custom_groups[self._custom_groups.index(old_name)] = new_name
        for tag, group in list(self._field_assign.items()):
            if group == old_name:
                self._field_assign[tag] = new_name
        self._reload_ui(rebuild_pages=True)
        self._remember()
        self.status_bar.showMessage(f"分组已重命名：{old_name} → {new_name}", 4000)

    def _delete_group(self, name):
        """删除自定义分组，组内字段回落到默认页"""
        reply = QMessageBox.question(
            self, "删除分组",
            f"确定删除分组「{name}」？\n\n组内字段会回到默认页，字段内容不受影响。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        if name in self._custom_groups:
            self._custom_groups.remove(name)
        for tag, group in list(self._field_assign.items()):
            if group == name:
                self._field_assign.pop(tag, None)
        self._reload_ui(rebuild_pages=True)
        self._remember()
        self.status_bar.showMessage(f"分组已删除：{name}", 4000)

    def _reload_ui(self, rebuild_pages=False):
        """重建界面（可选重建页面骨架），并保留已填内容"""
        context, loops, sections = ({}, {}, [])
        if self._field_widgets:
            context, loops, sections = self._collect_context()
        if rebuild_pages:
            self._rebuild_pages()
        self._build_form(self.batch.tags, self.batch.loops)
        self._apply_context(context, loops, sections)
        self._update_summary()

    # ── 字段网格 ────────────────────────────────────────────

    def _build_field_grid(self, page_key, tag_list):
        """兼容旧调用：把字段列表渲染成单列（现已统一走字段页）"""
        for tag in tag_list or []:
            self._build_field_row(tag)

    def _make_field_label(self, tag, display_name):
        """字段标题：双击改名，右键插入/删除；模板中不存在的变量标题标黄"""
        missing = tag not in self._parsed_tags
        name = QLabel(display_name)
        name.setStyleSheet(
            f"font-size: 13px; color: {'#ba7517' if missing else '#6b6b66'};"
            " background: transparent; padding-top: 7px;")
        name.setProperty("tag", tag)
        name.setWordWrap(False)
        tip = (f"变量名：{tag}\n所属分组：{self._resolve_group(tag)}\n"
               "双击：改显示名 / 控件类型 / 所属分组\n右键：插入或删除字段")
        if missing:
            tip += f"\n\n模板里没有 {{{{{tag}}}}}，这里填写的内容不会写入文档"
        name.setToolTip(tip)
        name.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        name.customContextMenuRequested.connect(
            lambda pos, t=tag, w=name: self._on_field_context_menu(t, w, pos))
        name.installEventFilter(self)
        return name

    # ── 字段插入 / 删除（右键菜单） ──────────────────────────

    def _on_field_context_menu(self, tag, widget, pos):
        """字段标题右键：编辑 / 上面插入 / 下面插入（级联选标题）/ 删除"""
        menu = QMenu(self)
        act_edit = menu.addAction("编辑字段...")
        menu.addSeparator()
        sub_above = menu.addMenu("在上面插入字段")
        sub_below = menu.addMenu("在下面插入字段")
        self._fill_insert_submenu(sub_above, "above", tag)
        self._fill_insert_submenu(sub_below, "below", tag)
        menu.addSeparator()
        act_delete = menu.addAction("删除此字段")

        chosen = menu.exec(widget.mapToGlobal(pos))
        if chosen is None:
            return
        if chosen is act_edit:
            self._edit_label(tag, widget)
            return
        if chosen is act_delete:
            self._delete_field(tag)
            return
        data = chosen.data()
        if isinstance(data, tuple) and len(data) == 2:
            where, base_tag = data
            self._insert_field_near(tag, base_tag, where)

    def _fill_insert_submenu(self, sub_menu, where, anchor_tag):
        """填充「插入字段」子菜单：空白新字段 + 复用已有标题"""
        act_blank = sub_menu.addAction("（空白新字段）")
        act_blank.setData((where, ""))
        sub_menu.addSeparator()
        for other in self._visible_tags():
            if other == anchor_tag:
                continue
            act = sub_menu.addAction(self._resolve_label(other))
            act.setData((where, other))

    def _suggest_tag(self):
        """为新字段生成一个未被占用的变量名建议"""
        used = set(self._visible_tags()) | set(self._removed_fields) | set(self._parsed_tags)
        i = 1
        while f"field_{i}" in used:
            i += 1
        return f"field_{i}"

    def _insert_field_near(self, anchor_tag, base_tag, where):
        """在 anchor_tag 的上/下方插入新字段；base_tag 非空则复用其显示名与控件类型"""
        from PySide6.QtWidgets import QDialog, QDialogButtonBox, QComboBox, QFormLayout

        base_label = self._resolve_label(base_tag) if base_tag else ''
        base_type = self._resolve_type(base_tag) if base_tag else 'single'
        anchor_label = self._resolve_label(anchor_tag)
        anchor_group = self._resolve_group(anchor_tag)

        dlg = QDialog(self)
        dlg.setWindowTitle("插入字段")
        dlg.setMinimumWidth(420)
        form = QFormLayout(dlg)

        tag_edit = QLineEdit(self._suggest_tag())
        tag_edit.setPlaceholderText("模板里的变量名，如 company")
        form.addRow("变量名：", tag_edit)

        name_edit = QLineEdit(base_label)
        name_edit.setPlaceholderText("显示在界面上的标题")
        form.addRow("显示名称：", name_edit)

        type_combo = QComboBox()
        type_combo.addItem("单行文本", "single")
        type_combo.addItem("多行文本", "multi")
        type_combo.setCurrentIndex(1 if base_type == 'multi' else 0)
        form.addRow("控件类型：", type_combo)

        pos_text = "上面" if where == 'above' else "下面"
        info = QLabel(f"插入位置：「{anchor_label}」{pos_text}　　归属分组：{anchor_group}")
        info.setStyleSheet("font-size: 12px; color: #75756f;")
        form.addRow(info)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                                   | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        form.addRow(buttons)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        tag = (tag_edit.text() or '').strip()
        label = (name_edit.text() or '').strip() or tag
        kind = type_combo.currentData()
        if not tag:
            QMessageBox.information(self, "提示", "变量名不能为空。")
            return
        if tag in self._visible_tags():
            QMessageBox.information(self, "提示",
                                    f"变量「{tag}」已经在填写界面上，不能重复添加。")
            return

        # 若是之前被隐藏的字段，这一步同时把它放回来
        self._removed_fields = [t for t in self._removed_fields if t != tag]
        if tag not in self._template_tags:
            self._custom_fields[tag] = {'label': label, 'type': kind}

        self._custom_labels[tag] = label
        self._custom_types[tag] = kind
        self._json_labels[tag] = {'label': label, 'type': kind}
        self._set_field_group(tag, anchor_group)

        order = self._visible_tags()
        pos = order.index(anchor_tag) if anchor_tag in order else len(order)
        if where == 'below':
            pos += 1
        order.insert(pos, tag)
        self._field_order = order

        self._reload_ui()
        self._remember()
        self._save_labels_json()
        self.status_bar.showMessage(f"已插入字段「{label}」（变量 {tag}）", 4000)

    def _delete_field(self, tag):
        """从填写界面移除字段（不改动模板文件）"""
        if tag not in self._visible_tags():
            return
        label = self._resolve_label(tag)
        reply = QMessageBox.question(
            self, "移除字段",
            f"从填写界面移除「{label}」（变量 {tag}）？\n\n"
            "模板文件不会改动。需要时可以右键其它字段把它插回来，"
            "或用「⋯ → 清除本模板记忆」恢复全部字段。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        self._custom_fields.pop(tag, None)
        if tag not in self._removed_fields:
            self._removed_fields.append(tag)
        self._field_order = [t for t in self._field_order if t != tag]
        self._reload_ui()
        self._remember()
        self.status_bar.showMessage(f"已从界面移除字段「{label}」", 4000)

    def _count_body_blocks(self):
        """统计动态正文章节数（无正文标记时为 0）"""
        if not self._has_body_marker or not self._section_editor:
            return 0
        try:
            return len(self._section_editor.get_sections())
        except Exception:
            return 0

    def _build_loop_page(self, loops):
        """循环表格页"""
        count = len(loops or {})
        if not loops:
            self._page_header("loops", "循环表格", "无循环区域")
            self._empty_hint("loops", "该模板没有 {% for %} 循环区域。\n\n"
                                      "循环区域用于模板中重复出现的表格数据（如明细行），"
                                      "在模板里写成 {% for item in items %} 形式即可被识别。")
            return
        self._page_header("loops", "循环表格", f"{count} 个循环区域 · 每个区域可添加多行数据")
        for loop_var, loop_info in loops.items():
            self._build_loop_section(loop_var, loop_info['fields'], loop_info)
        self._page_layouts[self._page_index("loops")].addStretch()

    def _build_loop_section(self, loop_var, fields, loop_info):
        """为循环区域构建 QTableWidget 输入区"""
        item_var = loop_info.get('item_var', 'item')
        display_name = self._resolve_label(loop_var)

        box = QWidget()
        layout = QVBoxLayout(box)
        layout.setContentsMargins(0, 8, 0, 0)
        layout.setSpacing(8)

        head = QHBoxLayout()
        head.setSpacing(10)
        title = QLabel(display_name)
        title.setStyleSheet("font-size: 13px; font-weight: 500; color: #2c2c2a;")
        head.addWidget(title)
        code = QLabel("{% for " + item_var + " in " + loop_var + " %}")
        code.setStyleSheet("font-size: 11px; color: #b5b5b0;")
        head.addWidget(code)
        head.addStretch()

        btn_add = QPushButton("＋ 添加行")
        btn_add.setStyleSheet(BUTTON_SECONDARY)
        btn_add.clicked.connect(lambda checked=False, lv=loop_var, f=fields: self._add_loop_row(lv, f))
        head.addWidget(btn_add)

        btn_del = QPushButton("− 删除选中行")
        btn_del.setStyleSheet(BUTTON_SECONDARY)
        btn_del.clicked.connect(lambda checked=False, lv=loop_var: self._remove_loop_row(lv))
        head.addWidget(btn_del)
        layout.addLayout(head)

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
                table.setColumnWidth(col_idx, 52)
            else:
                table.setColumnWidth(col_idx, 120)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setAlternatingRowColors(False)
        table.setStyleSheet("""
            QTableWidget { border: 1px solid #e3e3e0; border-radius: 3px;
                          gridline-color: #efefec; font-size: 12.5px; background: #ffffff; }
            QTableWidget::item { padding: 4px 6px; }
            QTableWidget::item:selected { background: #eceff5; color: #2c2c2a; }
            QHeaderView::section { background: #f7f7f5; font-weight: 500; color: #6b6b66;
                                  padding: 6px; border: none;
                                  border-right: 1px solid #efefec; border-bottom: 1px solid #e3e3e0; }
        """)
        table.setMinimumHeight(160)
        table.setMaximumHeight(420)
        table.verticalHeader().setDefaultSectionSize(40)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(table)

        self._loop_tables[loop_var] = table
        self._loop_fields[loop_var] = fields
        self._page_layouts[self._page_index("loops")].addWidget(box)

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
        """双击字段名 → 弹出改名对话框"""
        if event.type() == QEvent.MouseButtonDblClick:
            tag = obj.property("tag")
            if tag:
                self._edit_label(tag, obj)
                return True
        return super().eventFilter(obj, event)

    def _edit_label(self, tag, name_widget):
        """弹出对话框让用户修改显示名、控件类型与所属分组，并自动写回 .labels.json"""
        from PySide6.QtWidgets import QDialog, QDialogButtonBox, QComboBox, QFormLayout

        NEW_GROUP = "＋ 新建分组…"
        current_label = self._resolve_label(tag)
        current_type = self._resolve_type(tag)
        current_group = self._resolve_group(tag)

        dlg = QDialog(self)
        dlg.setWindowTitle("编辑字段")
        dlg.setMinimumWidth(360)
        form = QFormLayout(dlg)

        name_edit = QLineEdit(current_label)
        name_edit.setPlaceholderText(f"变量名：{tag}")
        form.addRow("显示名称：", name_edit)

        type_combo = QComboBox()
        type_combo.addItem("单行文本", "single")
        type_combo.addItem("多行文本", "multi")
        type_combo.setCurrentIndex(0 if current_type == 'single' else 1)
        form.addRow("控件类型：", type_combo)

        group_combo = QComboBox()
        for name in self._group_names():
            label = name
            if name in (self.INFO_GROUP, self.FIELDS_GROUP):
                label = f"{name}（内置）"
            group_combo.addItem(label, name)
        group_combo.addItem(NEW_GROUP, NEW_GROUP)
        idx = group_combo.findData(current_group)
        group_combo.setCurrentIndex(idx if idx >= 0 else 0)
        form.addRow("所属分组：", group_combo)

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

        # 分组：可能是新建
        new_group = group_combo.currentData()
        if new_group == NEW_GROUP:
            text, ok = QInputDialog.getText(self, "新建分组", "分组名称：")
            if not ok or not (text or "").strip():
                new_group = current_group
            else:
                created = self._new_group(text)
                new_group = created or current_group

        self._custom_labels[tag] = new_label
        self._custom_types[tag] = new_type
        group_changed = new_group != current_group
        self._set_field_group(tag, new_group)

        # 写入 _json_labels（新格式）
        self._json_labels[tag] = {
            'label': new_label,
            'type': new_type,
        }

        if name_widget is not None:
            if hasattr(name_widget, 'setTitle'):
                name_widget.setTitle(new_label)
            else:
                name_widget.setText(new_label)

        # 控件类型或分组变化 → 重建界面（保留已填内容）
        if new_type != current_type or group_changed:
            self._reload_ui(rebuild_pages=group_changed)
            self._remember()

        # 自动写回 .labels.json
        self._save_labels_json()

        if group_changed:
            self.status_bar.showMessage(f"「{new_label}」已归入分组：{new_group}", 4000)

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
        """收集当前表单所有数据，返回 (context, loops, sections)"""
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

        # 收集动态章节数据
        sections = []
        if self._section_editor:
            sections = self._section_editor.get_sections()

        return context, loops, sections

    def _apply_context(self, context, loops_data=None, sections=None):
        """把填写数据回填到当前表单（用于打开项目 / 重建界面 / 恢复记忆）"""
        for tag, widget in self._field_widgets.items():
            if tag not in (context or {}):
                continue
            val = context[tag]
            val = '' if val is None else str(val)
            if isinstance(widget, QTextEdit):
                widget.setPlainText(val)
            else:
                widget.setText(val)

        for loop_var, items in (loops_data or {}).items():
            table = self._loop_tables.get(loop_var)
            fields = self._loop_fields.get(loop_var, [])
            if not table or not fields:
                continue
            table.setRowCount(0)
            for item_data in items:
                row = table.rowCount()
                table.insertRow(row)
                table.setRowHeight(row, 40)
                for col, field_name in enumerate(fields):
                    if col >= table.columnCount():
                        break
                    value = item_data.get(field_name, '') if isinstance(item_data, dict) else ''
                    table.setItem(row, col,
                                  QTableWidgetItem('' if value is None else str(value)))

        if sections and self._section_editor:
            self._section_editor.set_sections(sections)

    # ── 填写记忆（同名模板自动带出上次内容） ────────────────

    def _collect_state(self):
        """汇总当前全部可记忆内容"""
        context, loops, sections = self._collect_context()
        state = {
            'values': {k: v for k, v in context.items() if str(v or '').strip()},
            'loops': {k: v for k, v in loops.items() if v},
            'sections': sections,
            'labels': dict(self._custom_labels),
            'types': dict(self._custom_types),
            'groups': list(self._custom_groups),
            'assign': dict(self._field_assign),
            'custom_fields': {k: dict(v) for k, v in self._custom_fields.items()},
            'removed': list(self._removed_fields),
            'order': list(self._field_order),
            'updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }
        if self.batch.mode and self.batch.table:
            rows = []
            for row in range(self.batch.table.rowCount()):
                row_data = {}
                for col in range(self.batch.table.columnCount()):
                    if col >= len(self.batch.tags):
                        continue
                    item = self.batch.table.item(row, col)
                    row_data[self.batch.tags[col]] = item.text().strip() if item else ''
                if any(v for v in row_data.values()):
                    rows.append(row_data)
            state['batch'] = rows
            state['batch_mode'] = True
        return state

    def _remember(self):
        """把当前填写内容写入模板记忆（静默，失败不影响主流程）"""
        if not self._template_path:
            return False
        try:
            return save_state(self._template_path, self._collect_state())
        except Exception:
            return False

    def _load_previous_state(self, template_path):
        """读取模板记忆（异常时返回空 dict）"""
        try:
            return load_state(template_path)
        except Exception:
            return {}

    def _apply_state_config(self, state):
        """恢复与布局有关的状态（显示名 / 控件类型 / 分组 / 自定义字段与顺序）"""
        self._custom_labels = dict(state.get('labels') or {})
        self._custom_types = dict(state.get('types') or {})
        groups = state.get('groups') or []
        self._custom_groups = [g for g in groups
                               if isinstance(g, str) and g.strip()
                               and g not in (self.INFO_GROUP, self.FIELDS_GROUP)]
        assign = state.get('assign') or {}
        self._field_assign = {k: v for k, v in assign.items() if isinstance(v, str)}

        custom = state.get('custom_fields') or {}
        self._custom_fields = {
            k: {'label': str((v or {}).get('label') or k),
                'type': (v or {}).get('type') or 'single'}
            for k, v in custom.items() if isinstance(k, str) and k.strip()}
        removed = state.get('removed') or []
        self._removed_fields = [t for t in removed if isinstance(t, str) and t.strip()]
        order = state.get('order') or []
        self._field_order = [t for t in order if isinstance(t, str) and t.strip()]

    def _restore_state(self, state):
        """恢复填写内容（字段值 / 循环 / 章节 / 批量数据）"""
        if not state:
            return 0
        self._apply_context(state.get('values') or {}, state.get('loops') or {},
                            state.get('sections') or [])

        # 批量数据：仅当当前处于批量模式且表格已建立
        rows = state.get('batch') or []
        if self.batch.mode and self.batch.table and rows:
            table = self.batch.table
            table.setRowCount(0)
            for row_data in rows:
                row = table.rowCount()
                table.insertRow(row)
                table.setRowHeight(row, 32)
                for tag, val in row_data.items():
                    if tag in self.batch.tags:
                        table.setItem(row, self.batch.tags.index(tag),
                                      QTableWidgetItem('' if val is None else str(val)))

        filled, sections, batch_rows = state_summary(state)
        return filled

    def _on_clear_values(self):
        """清空所有输入内容（保留分组与显示名设置）"""
        if not self._template_path:
            return
        reply = QMessageBox.question(
            self, "清空已填内容",
            "确定清空当前模板的所有填写内容？\n\n字段分组与显示名设置会保留。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        for widget in self._field_widgets.values():
            if isinstance(widget, QTextEdit):
                widget.setPlainText('')
            else:
                widget.setText('')
        for table in self._loop_tables.values():
            table.setRowCount(0)
        if self._section_editor:
            self._section_editor.clear()
        if self.batch.mode and self.batch.table:
            self.batch.table.setRowCount(0)
        self._update_summary()
        self._refresh_nav_counts()
        self._remember()
        self.status_bar.showMessage("已清空填写内容", 4000)

    def _on_forget_template(self):
        """删除本模板记忆，并把界面配置恢复为默认"""
        if not self._template_path:
            QMessageBox.information(self, "提示", "请先加载一个模板文件。")
            return
        reply = QMessageBox.question(
            self, "清除本模板记忆",
            f"确定清除「{os.path.basename(self._template_path)}」的填写记忆？\n\n"
            "将删除上次填写内容、自定义分组、显示名设置，"
            "以及手工新增/隐藏的字段，此操作不可撤销。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        clear_state(self._template_path)
        self._restored_state = {}
        self._custom_labels = {}
        self._custom_types = {}
        self._custom_groups = []
        self._field_assign = {}
        self._custom_fields = {}
        self._removed_fields = []
        self._field_order = []
        self._json_labels = load_tag_labels(self._template_path)
        # 重建界面：不保留已填内容，呈现为「刚打开模板」的干净状态
        self._rebuild_pages()
        self._build_form(self.batch.tags, self.batch.loops)
        self._update_summary()
        self._refresh_nav_counts()
        self.status_bar.showMessage("已清除本模板记忆", 4000)

    def _save_project(self):
        """将当前填写数据 + 标签配置保存为 .tplfill 项目文件"""
        if not self._template_path:
            QMessageBox.information(self, "提示", "请先加载一个模板文件。")
            return

        # 先确定保存位置，模板相对路径要以项目文件所在目录为基准
        # （否则项目文件与模板不在同一目录时，打开项目会找不到模板）
        default_name = os.path.splitext(os.path.basename(self._template_path))[0] + '.tplfill'
        if self._project_path:
            default_name = os.path.basename(self._project_path)

        path, _ = QFileDialog.getSaveFileName(
            self, "保存项目", os.path.join(self._last_dir, default_name),
            "TemplateFill 项目 (*.tplfill);;所有文件 (*.*)"
        )
        if not path:
            return

        context, loops, sections = self._collect_context()
        if self.batch.mode and not self._field_widgets:
            # 批量模式下没有独立输入框：沿用记忆里的单人模式内容，
            # 否则项目文件会丢掉单人模式填过的值
            fallback = self._restored_state or \
                self._load_previous_state(self._template_path) or {}
            context = fallback.get('values') or {}
            loops = fallback.get('loops') or {}
            sections = fallback.get('sections') or []

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

        # 模板路径存为相对路径（相对项目文件所在目录，跨盘时退回绝对路径）
        try:
            template_path = os.path.relpath(self._template_path,
                                            os.path.dirname(os.path.abspath(path)))
        except ValueError:
            template_path = self._template_path

        data = {
            'version': 1,
            'template_path': template_path.replace('\\', '/'),
            'template_format': self._template_format,
            'context': context,
            'loops': loops,
            'sections': sections,
            'custom_labels': self._custom_labels,
            'custom_types': self._custom_types,
            'custom_groups': self._custom_groups,
            'field_assign': self._field_assign,
            'batch_mode': self.batch.mode,
            'batch_defaults': self.batch.defaults,
            'batch_data': batch_data,
        }

        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self._project_path = path
            self._last_dir = os.path.dirname(path)
            self._update_title()
            self.status_bar.showMessage(f"项目已保存：{path}")
        except (IOError, PermissionError) as e:
            QMessageBox.critical(self, "保存失败", f"无法写入文件：{e}")

    def _open_project(self):
        """打开 .tplfill 项目文件，恢复编辑状态"""
        path, _ = QFileDialog.getOpenFileName(
            self, "打开项目", self._last_dir,
            "TemplateFill 项目 (*.tplfill);;所有文件 (*.*)"
        )
        if not path:
            return
        self._last_dir = os.path.dirname(path)

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

        # 解析模板路径：优先相对项目文件所在目录，其次原样路径，再次同目录同名文件
        project_dir = os.path.dirname(os.path.abspath(path))
        raw = data['template_path']
        candidates = [
            os.path.normpath(os.path.join(project_dir, raw)),
            raw,
            os.path.join(project_dir, os.path.basename(raw)),
        ]
        template_path = next((c for c in candidates if c and os.path.exists(c)), None)
        if template_path is None:
            template_path = candidates[0]
            reply = QMessageBox.question(
                self, "模板未找到",
                f"项目引用的模板文件不存在：\n{template_path}\n\n是否手动定位模板文件？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                new_path, _ = QFileDialog.getOpenFileName(
                    self, "选择模板文件", self._last_dir,
                    "Office 模板 (*.docx *.xlsx);;所有文件 (*.*)"
                )
                if not new_path:
                    return
                template_path = new_path
            else:
                return

        # 批量模式状态需在加载模板（建表格）之前确定
        want_batch = bool(data.get('batch_mode'))
        if want_batch != self.batch.mode:
            self.batch.mode = want_batch
            self.batch.defaults = data.get('batch_defaults', {}) if want_batch else {}

        # 加载模板（项目数据优先，不用模板记忆覆盖）
        self._load_template(template_path, restore_memory=False)

        # 恢复自定义标签配置与字段分组
        if data.get('custom_labels'):
            self._custom_labels = data['custom_labels']
        if data.get('custom_types'):
            self._custom_types = data['custom_types']
        groups = data.get('custom_groups')
        if groups:
            self._custom_groups = [g for g in groups
                                   if isinstance(g, str) and g.strip()
                                   and g not in (self.INFO_GROUP, self.FIELDS_GROUP)]
        if data.get('field_assign'):
            self._field_assign = {k: v for k, v in data['field_assign'].items()
                                  if isinstance(v, str)}

        # 重建表单以应用自定义显示名与分组，再统一回填数据
        self._rebuild_pages()
        self._build_form(self.batch.tags, self.batch.loops)
        self.batch.sync_buttons()
        self._apply_context(data.get('context', {}), data.get('loops', {}),
                            data.get('sections', []))

        # 暂存项目里的单人模式内容：批量模式下没有独立输入框，
        # 切回单人模式时用这份数据回填（优先于模板记忆）
        self._restored_state = {
            'values': {k: v for k, v in (data.get('context') or {}).items()
                       if str(v or '').strip()},
            'loops': data.get('loops') or {},
            'sections': data.get('sections') or [],
        }

        # 批量数据（表格此时才存在，晚于建表填充才不会被重建冲掉）
        batch_data = data.get('batch_data') or []
        if self.batch.mode and self.batch.table and batch_data:
            table = self.batch.table
            table.setRowCount(0)
            for row_data in batch_data:
                row = table.rowCount()
                table.insertRow(row)
                table.setRowHeight(row, 32)
                for tag, val in row_data.items():
                    if tag in self.batch.tags:
                        table.setItem(row, self.batch.tags.index(tag),
                                      QTableWidgetItem(str(val) if val else ''))

        self._project_path = path
        self._update_title()
        self._update_summary()
        self._refresh_nav_counts()
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

        context, loops_raw, sections = self._collect_context()

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
            out_arg = self._resolve_output_path()
            if self._template_format == 'xlsx':
                out_path = fill_xlsx_template(self._template_path, context, out_arg)
            else:
                # 若模板含 {{__BODY__}} 且有章节数据，传递给 fill_template
                sections_arg = sections if self._has_body_marker else None
                out_path = fill_template(self._template_path, context,
                                         output_path=out_arg, sections=sections_arg)

            self._show_result(out_path)
        except Exception as e:
            QMessageBox.critical(self, "生成失败", str(e))

    def _show_result(self, out_path):
        """生成结果反馈：状态栏 + 最近输出区，并按设置自动打开"""
        self._last_output = out_path
        self._remember()          # 生成成功 → 记住本次填写内容
        self.status_bar.showMessage(f"文档已生成：{out_path}")
        self._update_recent_output()
        self._update_summary()
        self._goto_page("export")

        action = self._output_action()
        if action == 'open_file':
            self._open_path(out_path)
        elif action == 'open_dir':
            self._open_path(os.path.dirname(out_path))

    # ── 托盘 / 关闭行为 ─────────────────────────────────────

    @staticmethod
    def _tray_available():
        """系统托盘是否可用（抽成方法，便于离屏测试替换）"""
        return QSystemTrayIcon.isSystemTrayAvailable()

    def _setup_tray(self):
        """右下角托盘图标 —— 收进托盘后程序继续在后台运行

        关闭行为由 settings.json 的 close_action 决定（ask / tray / quit，出厂 ask）。
        系统托盘不可用的环境里静默跳过，点 X 仍照原样退出 —— 绝不出现
        「窗口不见了、程序还在」这种找不回来的状态。
        """
        if not self._tray_available():
            return None
        try:
            icon = QIcon(resource_path("TemplateFill.ico"))
            tray = QSystemTrayIcon(icon, self)
            tray.setContextMenu(self._build_tray_menu())
            tray.activated.connect(self._on_tray_activated)
            self._tray = tray
            self._update_tray_tooltip()
            tray.show()
            return tray
        except Exception:
            self._tray = None            # 托盘建不起来就当作不可用，走降级路径
            return None

    def _build_tray_menu(self):
        """托盘右键菜单（与「托盘是否可用」解耦，离屏测试可直接取用）

        ⚠ 返回的 QMenu 必须长期持有（self._tray_menu）—— PySide6 里 wrapper
        一被回收就连带删掉底层 C++ 对象，菜单随即失效。
        """
        menu = QMenu(self)
        self._add_menu_action(menu, "显示主窗口", self._restore_from_tray)
        self._add_menu_action(menu, "打开模板...", self._tray_open_template)
        menu.addSeparator()
        # 「关闭窗口时」—— 记住选择之后想改回来，这里是唯一入口
        behavior = menu.addMenu("关闭窗口时")
        self._close_group = QActionGroup(self)
        self._close_group.setExclusive(True)
        self._close_actions = {}
        for key, label in self.CLOSE_LABELS:
            act = QAction(label, self)
            act.setCheckable(True)
            act.setChecked(key == self._resolve_close_action())
            act.triggered.connect(lambda _checked=False, k=key: self._set_close_action(k))
            self._close_group.addAction(act)
            behavior.addAction(act)
            self._close_actions[key] = act
        menu.addSeparator()
        self._add_menu_action(menu, "退出 TemplateFill", self._quit_app)
        self._tray_menu = menu
        return menu

    @staticmethod
    def _add_menu_action(menu, text, slot):
        act = QAction(text, menu)
        act.triggered.connect(slot)
        menu.addAction(act)
        return act

    def _update_tray_tooltip(self):
        """托盘悬停提示：软件名 + 版本 + 当前模板名"""
        if self._tray is None:
            return
        tip = f"TemplateFill v{VERSION}" if VERSION else "TemplateFill"
        if self._template_path:
            tip += f" — {os.path.basename(self._template_path)}"
        else:
            tip += " — 模板文档生成器"
        self._tray.setToolTip(tip)

    def _on_tray_activated(self, reason):
        """单击 / 双击托盘图标 → 唤回主窗口（右键不误唤回）"""
        if reason in (QSystemTrayIcon.ActivationReason.Trigger,
                      QSystemTrayIcon.ActivationReason.DoubleClick):
            self._restore_from_tray()

    def _minimize_to_tray(self):
        """把主窗口收进托盘（窗口对象仍存活，数据落盘）

        收进托盘也算一次落盘 —— 此后即使直接从托盘退出或被强杀，填写内容都不丢。
        """
        if self._tray is None:            # 没有托盘就别把窗口藏起来，否则找不回来
            self._quit_app()              # 退回真退出（同关闭窗口时选「直接退出」）
            return False
        self._pre_tray_state = self.windowState()
        self._save_all()
        self.hide()
        if not self._tray_tip_shown:
            self._tray_tip_shown = True
            self._tray.showMessage(
                "TemplateFill 仍在后台运行",
                "双击右下角图标可重新打开窗口；右键图标 → 退出 TemplateFill 可完全关闭。",
                QSystemTrayIcon.MessageIcon.Information, 5000)
        self.status_bar.showMessage("已收进右下角托盘，程序仍在后台运行", 4000)
        return True

    def _restore_from_tray(self):
        """把主窗口从托盘唤回来（保持收进去之前的大小状态）"""
        self.show()
        if self._pre_tray_state is not None:
            self.setWindowState(self._pre_tray_state)   # 不能用 showNormal()，会丢最大化
        self.raise_()
        self.activateWindow()

    def _tray_open_template(self):
        """托盘菜单「打开模板...」：先把窗口唤回来再弹选择框

        窗口隐藏时弹出的对话框没有可见父窗口，容易挂到屏幕外/被压在下面，
        表现为「点了没反应」—— 所以这里先恢复窗口。
        """
        if self.isHidden():
            self._restore_from_tray()
        self._on_open_template()

    def _resolve_close_action(self):
        """当前生效的关闭行为（文件损坏/取值非法一律退回「每次询问」）"""
        return get_close_action()

    def _set_close_action(self, action):
        """写入关闭行为偏好，并同步托盘菜单里的勾选"""
        if not set_close_action(action):
            return False
        act = self._close_actions.get(action)
        if act is not None and not act.isChecked():
            act.setChecked(True)
        self.status_bar.showMessage(
            f"关闭窗口时：{dict(self.CLOSE_LABELS).get(action, action)}", 4000)
        return True

    def _ask_close_action(self):
        """弹三选一；返回 (action, remember)"""
        dlg = CloseChoiceDialog(self)
        dlg.exec()
        return dlg.choice, dlg.remember

    def _save_all(self):
        """退出/收托盘前的落盘（可重复调用，失败静默）"""
        try:
            self._remember()
        except Exception:
            pass

    def _shutdown(self):
        """退出前收尾：摘掉托盘图标 + 落盘（可重复调用）"""
        if self._tray is not None:
            self._tray.hide()
        self._save_all()

    def _quit_app(self):
        """**结束进程的唯一出口** —— 不再询问、不再收托盘

        三件事缺一不可：① 标记真退出（拦住托盘逻辑）② 收尾落盘并摘掉托盘图标
        ③ `QApplication.quit()` 结束事件循环。

        ⚠ ③ 不能省：main() 里设了 `setQuitOnLastWindowClosed(False)`（收进托盘时
        防止被 Qt 当成「最后一个窗口已关闭」而结束进程），副作用是**光关窗口不会
        结束进程**。只 accept() 关闭事件的话，窗口关了、托盘也摘了，进程却在后台
        活着 —— 窗口和托盘两个入口都没了，用户再也找不回来。
        """
        self._really_quit = True
        self._shutdown()
        QApplication.quit()

    def closeEvent(self, event):
        """点窗口「X」的关闭流程

        出厂默认「每次询问」（本次界面上三选一，可勾选记住）；记住之后点 X 直接
        照办、不再打扰，改回来的入口在托盘右键菜单「关闭窗口时」。收进托盘只是
        hide() —— 窗口对象仍存活，从托盘唤回即可，只有真退出才落盘并结束进程。

        ⚠ 「直接退出」这一支**必须走 `_quit_app()`**，不能只 `event.accept()`：
        main() 设了 `setQuitOnLastWindowClosed(False)`，关掉窗口不会结束进程，
        否则窗口和托盘同时消失、进程却在后台活着，用户两边都找不回来。
        """
        if self._really_quit:             # 已在退出流程中（_quit_app 的重入）
            event.accept()
            return

        action = self._resolve_close_action()
        if action == "ask":
            action, remember = self._ask_close_action()
            if remember and action != "cancel":
                self._set_close_action(action)

        if action == "tray":
            event.ignore()
            self._minimize_to_tray()
            return

        if action == "quit":
            event.accept()                # 先按用户意愿关窗
            self._quit_app()              # 再结束进程（落盘 + 摘托盘 + quit）
            return

        event.ignore()                    # 取消：留在界面上继续用

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("TemplateFill")
    app.setApplicationVersion(VERSION or "")
    app.setFont(QFont("Microsoft YaHei", 10))
    # 主窗口可以被收进托盘（hide()）—— 此时"最后一个窗口关闭"不代表要退出进程，
    # 否则一收托盘程序就没了。真退出统一走 TemplateFillWindow._quit_app()
    app.setQuitOnLastWindowClosed(False)
    window = TemplateFillWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
