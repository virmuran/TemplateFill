"""TemplateFill 字段页改版回归测试（v1.2.0）

覆盖：
  1. AutoGrowTextEdit：单行起步、换行/软换行自动增高、8 行上限、Tab 切焦点、纯文本
  2. 字段页单列布局：一行一个字段、左标题右输入框、限宽、分组小节标题
  3. 右键插入字段：位置（上/下）、继承分组、复用已有标题的显示名与控件类型
  4. 模板中不存在的变量：标题标黄 + 行内提示
  5. 删除字段与插回、顺序与自定义字段的持久化

运行：QT_QPA_PLATFORM=offscreen python test_v12.py
"""

import os
import sys
import shutil
import tempfile

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtWidgets import (  # noqa: E402
    QApplication, QDialog, QLabel, QLineEdit, QMessageBox,
)

app = QApplication([])

import field_store as FS          # noqa: E402
import main as M                  # noqa: E402
from doc_filler import create_sample_template  # noqa: E402
from field_widgets import AutoGrowTextEdit     # noqa: E402

RESULTS = []


def check(name, cond, extra=''):
    RESULTS.append((name, bool(cond), extra))


tmp = tempfile.mkdtemp(prefix='tf_v12_')
FS.STORE_DIR = os.path.join(tmp, 'state')
M.TemplateFillWindow.TEMPLATES_DIR = os.path.join(tmp, 'no_templates')

tpl = os.path.join(tmp, 'demo.docx')
create_sample_template(tpl)

INFO_GROUP = M.TemplateFillWindow.INFO_GROUP

# 模态框替身
QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.StandardButton.Yes)
_msgs = []
QMessageBox.information = staticmethod(
    lambda *a, **k: (_msgs.append(a[2] if len(a) > 2 else ''), QMessageBox.StandardButton.Ok)[1])

_orig_exec = QDialog.exec
_orig_reload = None


def _tag_label(win, tag):
    """字段页上定位某个变量对应的标题标签"""
    for lb in win._page_containers[win._page_index('fields')].findChildren(QLabel):
        if lb.property('tag') == tag:
            return lb
    return None


def _page_labels(win):
    return [lb.text() for lb in
            win._page_containers[win._page_index('fields')].findChildren(QLabel)]


def _drive_insert(new_tag, new_label=None):
    """生成一个对话框替身：填好变量名（可选显示名）后确认"""
    def driver(dlg):
        for le in dlg.findChildren(QLineEdit):
            ph = le.placeholderText()
            if ph.startswith('模板里的变量名'):
                le.setText(new_tag)
            elif new_label is not None and ph.startswith('显示在界面上的标题'):
                le.setText(new_label)
        return QDialog.DialogCode.Accepted
    return driver


# ── 1. AutoGrowTextEdit 行为 ────────────────────────────────
ed = AutoGrowTextEdit(min_lines=1, max_lines=8)
ed.resize(400, ed.height())
ed.show()
app.processEvents()
h_blank = ed.height()
line_h = ed._line_height()

ed.setPlainText('一行内容')
app.processEvents()
h_one = ed.height()
check('空框与单行同高', h_blank == h_one, f'{h_blank} vs {h_one}')
check('单行高度不超过两行加边距', h_one <= line_h * 2 + ed._pad(),
      f'h={h_one} line={line_h} pad={ed._pad()}')

ed.setPlainText('第一行\n第二行\n第三行')
app.processEvents()
h_three = ed.height()
check('按回车换行后自动增高', h_three > h_one, f'{h_one} -> {h_three}')

ed.setPlainText('行')
app.processEvents()
check('删掉内容后缩回一行', ed.height() == h_one, f'{ed.height()} vs {h_one}')

ed.setPlainText('\n'.join(['内容'] * 20))
app.processEvents()
h20 = ed.height()
ed.setPlainText('\n'.join(['内容'] * 30))
app.processEvents()
check('达到上限后不再增高', ed.height() == h20, f'{h20} -> {ed.height()}')
check('超限后改为框内滚动', ed.verticalScrollBar().maximum() > 0,
      f'max={ed.verticalScrollBar().maximum()}')
check('上限为 8 行', ed.max_lines == 8 and h20 <= line_h * 10, f'h20={h20} line={line_h}')

ed2 = AutoGrowTextEdit(min_lines=1, max_lines=8)
ed2.resize(200, ed2.height())
ed2.show()
app.processEvents()
h0 = ed2.height()
ed2.setPlainText('很长的一段中文内容' * 10)
app.processEvents()
h_long = ed2.height()
check('一行放不下时软换行增高', h_long > h0, f'{h0} -> {h_long}')

check('Tab 键切换焦点（不插入制表符）', ed.tabChangesFocus() is True)
check('只接受纯文本', ed.acceptRichText() is False)
ed3 = AutoGrowTextEdit(min_lines=3, max_lines=8)
check('多行类型初始 3 行高', ed3.min_lines == 3)
check('有限制高度（不会被布局拉伸）', ed.minimumHeight() == ed.maximumHeight())

ed.deleteLater()
ed2.deleteLater()
ed3.deleteLater()

# ── 2. 字段页布局 ───────────────────────────────────────────
w = M.TemplateFillWindow()
check('测试环境未自动加载模板', w._template_path is None)
FS.clear_state(tpl)
w._load_template(tpl)
n0 = len(w._field_widgets)
check('字段已构建', n0 > 0, f'{n0} 个')
check('四个导航页', w.nav.count() == 4, f'{w.nav.count()}')

vis = w._visible_tags()
rows = [w._field_widgets[t].parent() for t in vis]
check('每个字段独占一行', len({id(r) for r in rows}) == len(rows),
      f'{len({id(r) for r in rows})} vs {len(rows)}')
check('一行只有一个输入框',
      all(len(r.findChildren(AutoGrowTextEdit)) == 1 for r in rows))


def _inline_pair(row):
    """行内第一个子布局应是「标题 + 输入框」同行"""
    lay = row.layout()
    if lay is None or lay.count() == 0:
        return False
    inner = lay.itemAt(0)
    inner = inner.layout() if inner else None
    if inner is None or inner.count() < 2:
        return False
    first, second = inner.itemAt(0).widget(), inner.itemAt(1).widget()
    return isinstance(first, QLabel) and isinstance(second, AutoGrowTextEdit)


check('左标题与右输入框同行', all(_inline_pair(r) for r in rows))
check('输入框统一限宽', all(w._field_widgets[t].maximumWidth() == w.FIELD_MAX_WIDTH
                            for t in vis), str(w.FIELD_MAX_WIDTH))
check('输入框都是自动增高控件',
      all(isinstance(w._field_widgets[t], AutoGrowTextEdit) for t in vis))
check('标题列宽按最长名自适应',
      w.FIELD_LABEL_MIN_WIDTH <= w._label_width <= w.FIELD_LABEL_MAX_WIDTH,
      f'{w._label_width}')
check('字段页渲染了分组小节标题',
      any(t.split('  ·  ')[0] == INFO_GROUP for t in _page_labels(w)),
      str(_page_labels(w)[:6]))
check('字段顺序即模板解析顺序', vis == list(w.batch.tags), str(vis[:4]))
check('模板中存在的变量不标黄',
      '#ba7517' not in (_tag_label(w, vis[0]).styleSheet() or ''))

# 真实窗口（应用样式）里的单行表现
w.show()
w._goto_page('fields')
app.processEvents()
_first = w._field_widgets[vis[0]]
_first.setPlainText('单行内容')
app.processEvents()
check('单行内容不出现滚动条', _first.verticalScrollBar().maximum() == 0,
      f'max={_first.verticalScrollBar().maximum()} h={_first.height()}')
check('最近填充的字段在导航计数里体现', w.nav.item(0).text().endswith(str(len(vis))),
      w.nav.item(0).text())
_first.setPlainText('\n'.join(['行'] * 6))
app.processEvents()
check('窗口内多行内容同样增高', _first.height() > 38, f'{_first.height()}')

# 真实按 Tab 键：焦点应交给下一个字段的输入框
from PySide6.QtTest import QTest  # noqa: E402
_second = w._field_widgets[vis[1]]
_first.setFocus()
app.processEvents()
QTest.keyClick(_first, Qt.Key.Key_Tab)
app.processEvents()
check('按 Tab 跳到下一个字段', app.focusWidget() is _second,
      type(app.focusWidget()).__name__ if app.focusWidget() else 'None')
QTest.keyClick(_second, Qt.Key.Key_Backtab)
app.processEvents()
check('Shift+Tab 回到上一个字段', app.focusWidget() is _first,
      type(app.focusWidget()).__name__ if app.focusWidget() else 'None')

# ── 3. 右键插入字段：空白新字段 ─────────────────────────────
QDialog.exec = _drive_insert('field_1')
w._insert_field_near('title', '', 'below')
QDialog.exec = _orig_exec

vis = w._visible_tags()
check('新字段已插入到锚点下方', vis.index('field_1') == vis.index('title') + 1,
      str(vis))
check('新字段已出现在界面', 'field_1' in w._field_widgets)
check('字段总数 +1', len(w._field_widgets) == n0 + 1, f'{len(w._field_widgets)}')
check('新字段继承锚点分组',
      w._resolve_group('field_1') == w._resolve_group('title'),
      f"{w._resolve_group('field_1')} vs {w._resolve_group('title')}")
check('模板无此变量 → 标题标黄',
      '#ba7517' in (_tag_label(w, 'field_1').styleSheet() or ''),
      _tag_label(w, 'field_1').styleSheet() if _tag_label(w, 'field_1') else '')
check('模板无此变量 → 行内给出提示',
      any('模板里没有' in t and 'field_1' in t for t in _page_labels(w)),
      str([t for t in _page_labels(w) if '模板里没有' in t]))
check('标黄字段的提示也写明原因',
      '不会写入文档' in (_tag_label(w, 'field_1').toolTip() or ''))

# 上方插入
QDialog.exec = _drive_insert('field_9')
w._insert_field_near('title', '', 'above')
QDialog.exec = _orig_exec
vis = w._visible_tags()
check('新字段已插入到锚点上方', vis.index('field_9') == vis.index('title') - 1,
      str(vis))

# 重复变量名被拒
_msgs.clear()
QDialog.exec = _drive_insert('field_1')
w._insert_field_near('title', '', 'below')
QDialog.exec = _orig_exec
check('重复变量名被拒绝', any('已经在填写界面上' in m for m in _msgs), str(_msgs))
check('拒绝后界面未被改动', len(w._field_widgets) == n0 + 2,
      f'{len(w._field_widgets)}')

# 空变量名被拒
_msgs.clear()
QDialog.exec = _drive_insert('')
w._insert_field_near('title', '', 'below')
QDialog.exec = _orig_exec
check('空变量名被拒绝', any('不能为空' in m for m in _msgs), str(_msgs))

# ── 4. 复用已有标题插入（继承显示名与控件类型） ─────────────
w._custom_types['author'] = 'multi'
author_label = w._resolve_label('author')
QDialog.exec = _drive_insert('field_2')       # 显示名留空 → 用对话框默认值（所选标题）
w._insert_field_near('author', 'author', 'above')
QDialog.exec = _orig_exec

vis = w._visible_tags()
check('复用标题插入到锚点上方', vis.index('field_2') == vis.index('author') - 1,
      str(vis))
check('复用标题：显示名继承',
      w._resolve_label('field_2') == author_label,
      f"{w._resolve_label('field_2')} vs {author_label}")
check('复用标题：控件类型继承', w._resolve_type('field_2') == 'multi',
      w._resolve_type('field_2'))
check('复用多行标题 → 初始 3 行',
      w._field_widgets['field_2'].min_lines == 3,
      str(w._field_widgets['field_2'].min_lines))

# ── 5. 删除字段 / 插回 ──────────────────────────────────────
w._delete_field('field_1')
check('字段已从界面移除', 'field_1' not in w._field_widgets)
check('字段记入隐藏名单', 'field_1' in w._removed_fields, str(w._removed_fields))
check('删除不动模板字段', 'title' in w._field_widgets)
check('剩余字段仍可填', w._field_widgets['title'].toPlainText() == '')

# ── 6. 持久化：顺序 / 自定义字段 / 隐藏字段 ─────────────────
order_before = w._visible_tags()
w._remember()
w2 = M.TemplateFillWindow()
w2._load_template(tpl)
check('重开模板：自定义字段仍在', 'field_2' in w2._field_widgets,
      str(sorted(w2._field_widgets)[:6]))
check('重开模板：隐藏字段保持隐藏', 'field_1' not in w2._field_widgets)
check('重开模板：字段顺序保持', w2._visible_tags() == order_before,
      f'{w2._visible_tags()} vs {order_before}')
check('重开模板：标黄提示仍在',
      '#ba7517' in (_tag_label(w2, 'field_2').styleSheet() or ''))
check('重开模板：自定义字段类型保持', w2._resolve_type('field_2') == 'multi')

# 把隐藏字段插回来
QDialog.exec = _drive_insert('field_1')
w2._insert_field_near('title', '', 'below')
QDialog.exec = _orig_exec
check('隐藏字段可插回界面', 'field_1' in w2._field_widgets)
check('插回后不再算隐藏', 'field_1' not in w2._removed_fields,
      str(w2._removed_fields))

# 字段值不受插入影响
w2._field_widgets['title'].setPlainText('标题内容')
QDialog.exec = _drive_insert('field_3')
w2._insert_field_near('title', '', 'above')
QDialog.exec = _orig_exec
check('插入字段后已填内容保留',
      w2._field_widgets['title'].toPlainText() == '标题内容',
      w2._field_widgets['title'].toPlainText())

# ── 7. 清除记忆后恢复干净状态 ───────────────────────────────
w3 = M.TemplateFillWindow()
w3._load_template(tpl)
check('清除前有自定义字段', 'field_2' in w3._field_widgets)
w3._on_forget_template()
check('清除后自定义字段消失', not any(t.startswith('field_')
                                       for t in w3._field_widgets),
      str(sorted(w3._field_widgets)))
check('清除后字段数回到模板原状', len(w3._field_widgets) == n0,
      f'{len(w3._field_widgets)} vs {n0}')

# ── 输出 ────────────────────────────────────────────────────
QDialog.exec = _orig_exec
passed = sum(1 for _, ok, _ in RESULTS if ok)
total = len(RESULTS)
lines = []
for name, ok, extra in RESULTS:
    mark = 'PASS' if ok else 'FAIL'
    lines.append(f'[{mark}] {name}' + (f'  ({extra})' if extra and not ok else ''))
lines.append('')
lines.append(f'合计 {passed} / {total} 通过')

report = '\n'.join(lines)
with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), '_v12_test_result.txt'),
          'w', encoding='utf-8') as f:
    f.write(report)
print(report)
shutil.rmtree(tmp, ignore_errors=True)
sys.exit(0 if passed == total else 1)
