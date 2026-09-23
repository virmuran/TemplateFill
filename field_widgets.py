"""TemplateFill — 字段输入控件

AutoGrowTextEdit：内容换行或按回车就自动增高的文本输入框。

设计要点：
  - 继承 QTextEdit：既有代码里的 isinstance(w, QTextEdit) 判断与
    setPlainText / toPlainText 调用无需改动
  - Tab 键切换焦点（tabChangesFocus）：不再插入制表符，便于连续填写
  - 高度随内容在 min_lines..max_lines 之间自适应，超出后框内滚动
  - AcceptRichText=False：粘贴只进纯文本，避免把网页样式带进文档
"""

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QSizePolicy, QTextEdit

DEFAULT_MAX_LINES = 8


class AutoGrowTextEdit(QTextEdit):
    """单行起步、随内容长高的文本输入框"""

    def __init__(self, min_lines=1, max_lines=DEFAULT_MAX_LINES, parent=None):
        super().__init__(parent)
        self._min_lines = max(1, int(min_lines))
        self._max_lines = max(self._min_lines, int(max_lines))
        self._syncing = False
        # 与 MAIN_STYLE 里 QTextEdit 的 padding 保持一致（6px 上下）
        self._vpad = 6

        self.setAcceptRichText(False)
        self.setTabChangesFocus(True)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.document().setDocumentMargin(4)
        self.document().contentsChanged.connect(self._sync_height)
        self._sync_height()

    # ── 对外只读属性（测试与调试用） ──

    @property
    def min_lines(self):
        return self._min_lines

    @property
    def max_lines(self):
        return self._max_lines

    def content_lines(self):
        """当前内容占用的显示行数（含软换行）"""
        return max(1, int(round(self.document().size().height() / self._line_height())))

    # ── 高度自适应 ──

    def _line_height(self):
        return max(1, self.fontMetrics().lineSpacing())

    def _pad(self):
        """边框 + 内边距 + 余量"""
        return 2 * int(self.frameWidth()) + 2 * self._vpad + 2

    def _sync_height(self):
        if self._syncing:
            return
        self._syncing = True
        try:
            doc = self.document()
            width = self.viewport().width()
            if width <= 0:
                width = max(1, self.width() - 2 * int(self.frameWidth()))
            doc.setTextWidth(width)
            need = doc.size().height() + self._pad()
            lo = self._line_height() * self._min_lines + self._pad()
            hi = self._line_height() * self._max_lines + self._pad()
            target = int(max(lo, min(need, hi)))
            if self.height() != target:
                self.setFixedHeight(target)
        finally:
            self._syncing = False

    def resizeEvent(self, event):        # noqa: N802 (Qt 命名)
        super().resizeEvent(event)
        self._sync_height()

    def setPlainText(self, text):        # noqa: N802
        super().setPlainText(text)
        self._sync_height()

    def setHtml(self, text):             # noqa: N802
        # 纯文本控件：HTML 一律按纯文本处理，避免外部样式污染
        self.setPlainText(text)
