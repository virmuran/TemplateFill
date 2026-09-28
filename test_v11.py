"""TemplateFill 新增功能回归测试：填写记忆 + 字段自定义分组

覆盖：
  1. field_store 读写/清理/容错
  2. 同名模板再次打开自动带出上次填写内容（字段值/章节/批量数据）
  3. 字段归组：新建分组 → 页内小节 → 字段迁移 → 状态持久化
  4. 分组重命名/删除、清空已填内容、清除模板记忆
  5. 项目文件保存/打开（含分组与批量数据）

运行：QT_QPA_PLATFORM=offscreen python test_v11.py
"""

import os
import sys
import json
import shutil
import tempfile

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import (  # noqa: E402
    QApplication, QComboBox, QDialog, QLabel, QMessageBox, QTableWidgetItem,
)

app = QApplication([])

import field_store as FS          # noqa: E402
import main as M                  # noqa: E402
from doc_filler import create_sample_template                     # noqa: E402
from xlsx_filler import create_sample_template as create_xlsx     # noqa: E402

RESULTS = []


def check(name, cond, extra=''):
    RESULTS.append((name, bool(cond), extra))


tmp = tempfile.mkdtemp(prefix='tf_v11_')
# 隔离：测试期间状态写到临时目录，且不让窗口自动加载真实示例模板
FS.STORE_DIR = os.path.join(tmp, 'state')
M.TemplateFillWindow.TEMPLATES_DIR = os.path.join(tmp, 'no_templates')

tpl = os.path.join(tmp, 'demo.docx')
create_sample_template(tpl)

# 对话框替身（离屏环境不能弹模态框）
_yes = lambda *a, **k: QMessageBox.StandardButton.Yes
QMessageBox.question = staticmethod(_yes)
_orig_info = QMessageBox.information
QMessageBox.information = staticmethod(lambda *a, **k: QMessageBox.StandardButton.Ok)
# 关闭窗口会走「每次询问」对话框，离屏下 exec() 永久阻塞且不报错 → 钉死成"取消"
M.CloseChoiceDialog.exec = lambda self: (setattr(self, 'choice', 'cancel'), 0)[1]


def _page_labels(win):
    """字段页上的全部 QLabel 文本（用于校验小节标题）"""
    return [lb.text() for lb in
            win._page_containers[win._page_index('fields')].findChildren(QLabel)]


# ── 1. field_store 基础 ─────────────────────────────────────
check('状态文件路径稳定', FS.state_file(tpl) == FS.state_file(tpl))
check('路径含模板名', 'demo' in os.path.basename(FS.state_file(tpl)))
check('初始无记忆', FS.has_state(tpl) is False)

state = {'values': {'title': '关于XX的通知', 'author': '沐然'},
         'loops': {}, 'sections': [{'title': '一、背景',
                                    'blocks': [{'type': 'text', 'text': '内容'}]}],
         'labels': {'title': '文件标题'}, 'types': {'title': 'single'},
         'groups': ['买方信息'], 'assign': {'author': '买方信息'}}
check('写入状态成功', FS.save_state(tpl, state) is True)
check('写入后可检测到记忆', FS.has_state(tpl) is True)

loaded = FS.load_state(tpl)
check('字段值可读回', loaded.get('values', {}).get('title') == '关于XX的通知')
check('章节可读回', len(loaded.get('sections') or []) == 1)
check('分组可读回', loaded.get('groups') == ['买方信息'])
check('归组映射可读回', loaded.get('assign', {}).get('author') == '买方信息')
check('版本号已写入', loaded.get('version') == FS.STATE_VERSION)
check('摘要统计正确', FS.state_summary(loaded) == (2, 1, 0))

with open(FS.state_file(tpl), 'w', encoding='utf-8') as f:
    f.write('{ 坏掉的 json')
check('损坏文件不抛异常', FS.load_state(tpl) == {})
FS.save_state(tpl, state)

# ── 2. 自动带出上次填写内容 ─────────────────────────────────
w = M.TemplateFillWindow()
check('测试环境不自动加载模板', w._template_path is None)
check('未加载时导航 4 项', w.nav.count() == 4, f'{w.nav.count()}')

FS.clear_state(tpl)          # 先清空，验证「填写 → 生成 → 再打开」的完整闭环
w._load_template(tpl)
n_fields = len(w._field_widgets)
check('模板已加载', w._template_path == tpl and n_fields > 0, f'{n_fields} 个字段')

for tag, widget in w._field_widgets.items():
    widget.setPlainText(f'内容-{tag}')
w._section_editor.set_sections([{'title': '测试章节',
                                 'blocks': [{'type': 'text', 'text': '段落'}]}])
w._on_generate()             # 生成成功 → 自动记忆

check('生成后已写入记忆', FS.has_state(tpl) is True)
saved = FS.load_state(tpl)
check('记忆含全部字段值', len(saved.get('values') or {}) == n_fields,
      f"{len(saved.get('values') or {})} vs {n_fields}")
check('记忆含章节', len(saved.get('sections') or []) == 1)

w2 = M.TemplateFillWindow()
w2._load_template(tpl)
check('二次打开字段数一致', len(w2._field_widgets) == n_fields)
vals = [wd.toPlainText() for wd in w2._field_widgets.values()]
check('二次打开带出全部填写内容', all(v.startswith('内容-') for v in vals) and vals,
      f'样例值 {vals[:2]}')
check('二次打开带出正文章节',
      len(w2._section_editor.get_sections()) == 1,
      str(len(w2._section_editor.get_sections())))
check('状态栏提示带出内容', '带出上次填写内容' in w2.status_bar.currentMessage(),
      w2.status_bar.currentMessage())

# 清空已填内容：内容清掉、分组设置保留
w2._new_group('保留组')
w2._set_field_group('title', '保留组')
w2._on_clear_values()
empty = [wd.toPlainText() for wd in w2._field_widgets.values()]
check('清空后输入框为空', all(v == '' for v in empty))
check('清空后分组设置保留', w2._custom_groups == ['保留组'])
check('清空后章节为空', len(w2._section_editor.get_sections()) == 0)

# ── 3. 字段自定义分组（页内小节，不再占导航位） ─────────────
FS.clear_state(tpl)          # 从干净状态开始，验证分组新建流程
w3 = M.TemplateFillWindow()
w3._load_template(tpl)
check('默认 4 页（字段页 + 三个固定页）', w3.nav.count() == 4, f'{w3.nav.count()}')

sections = w3._group_sections(w3._visible_tags())
check('字段默认切成小节',
      any(n == M.TemplateFillWindow.INFO_GROUP for n, _ in sections),
      str([n for n, _ in sections]))

w3._new_group('买方信息')
w3._new_group('卖方信息')
w3._set_field_group('title', '买方信息')
w3._reload_ui(rebuild_pages=True)

check('新建分组不新增导航页', w3.nav.count() == 4, f'{w3.nav.count()}')
sections = w3._group_sections(w3._visible_tags())
names = [n for n, _ in sections]
check('新分组成为字段页内小节', '买方信息' in names, str(names))
check('空分组不渲染小节', '卖方信息' not in names, str(names))
check('字段已迁入新分组小节', 'title' in dict(sections).get('买方信息', []),
      str(dict(sections).get('买方信息', [])))
check('字段总数不变', len(w3._visible_tags()) == n_fields,
      f'{len(w3._visible_tags())} vs {n_fields}')
check('页内小节标题带字段计数',
      any(name == '买方信息' and len(tags) == 1 for name, tags in sections),
      str([(n, len(t)) for n, t in sections]))
check('小节标题已渲染到字段页',
      any(t.startswith('买方信息') for t in _page_labels(w3)),
      str(_page_labels(w3)[:8]))

lbl = None
for child in w3._page_containers[w3._page_index('fields')].findChildren(QLabel):
    if child.property('tag') == 'title':
        lbl = child
check('能定位到迁移后的字段标题', lbl is not None)
check('标题提示显示当前分组', lbl is not None and '买方信息' in lbl.toolTip(),
      lbl.toolTip() if lbl else '')

# 归组后仍可正常填写与生成
w3._field_widgets['title'].setPlainText('归组后的标题')
check('归组后字段可填', w3._field_widgets['title'].toPlainText() == '归组后的标题')

# 分组持久化：新建窗口自动恢复分组
w3._remember()
w4 = M.TemplateFillWindow()
w4._load_template(tpl)
check('二次打开恢复自定义分组', w4._custom_groups == ['买方信息', '卖方信息'],
      str(w4._custom_groups))
check('二次打开仍为 4 页', w4.nav.count() == 4, f'{w4.nav.count()}')
check('二次打开恢复归组字段',
      w4._resolve_group('title') == '买方信息', w4._resolve_group('title'))
check('二次打开恢复字段内容',
      w4._field_widgets['title'].toPlainText() == '归组后的标题',
      w4._field_widgets['title'].toPlainText())

# ── 4. 分组重命名 / 删除 ────────────────────────────────────
M.QInputDialog.getText = staticmethod(lambda *a, **k: ('采购方信息', True))
w4._rename_group('买方信息')
check('分组已重命名', '采购方信息' in w4._custom_groups, str(w4._custom_groups))
check('归组字段跟随重命名', w4._resolve_group('title') == '采购方信息',
      w4._resolve_group('title'))
check('页内小节标题同步改名',
      any(t.startswith('采购方信息') for t in _page_labels(w4)),
      str(_page_labels(w4)[:8]))

w4._delete_group('采购方信息')
check('分组已删除', '采购方信息' not in w4._custom_groups, str(w4._custom_groups))
check('删除分组后仍是 4 页', w4.nav.count() == 4, f'{w4.nav.count()}')
check('删除后字段回落文档信息小节',
      w4._resolve_group('title') == M.TemplateFillWindow.INFO_GROUP,
      w4._resolve_group('title'))
check('删除分组不影响字段内容',
      w4._field_widgets['title'].toPlainText() == '归组后的标题')

# ── 5. 编辑字段对话框：归入分组（驱动真实对话框） ────────────
w5 = M.TemplateFillWindow()
w5._load_template(tpl)
w5._new_group('新建测试组')
w5._reload_ui(rebuild_pages=True)


def _drive_dialog(dlg):
    """替身 exec：把「所属分组」下拉切到 新建测试组，然后确认"""
    for combo in dlg.findChildren(QComboBox):
        idx = combo.findData('新建测试组')
        if idx >= 0:
            combo.setCurrentIndex(idx)
    return QDialog.DialogCode.Accepted


_orig_exec = QDialog.exec
QDialog.exec = _drive_dialog
w5._edit_label('author', QLabel('作者'))
QDialog.exec = _orig_exec
check('对话框可把字段归入分组', w5._resolve_group('author') == '新建测试组',
      w5._resolve_group('author'))
check('归组后保留既有填写内容', w5._field_widgets['author'].isEnabled())
w5._field_widgets['author'].setPlainText('作者A')
check('归组后字段仍可编辑', w5._field_widgets['author'].toPlainText() == '作者A')

# ── 6. 批量数据记忆 ─────────────────────────────────────────
xb = os.path.join(tmp, 'demo_batch.xlsx')
create_xlsx(xb)
FS.clear_state(xb)
wb = M.TemplateFillWindow()
wb._load_template(xb)
check('Excel 模板已加载', wb._template_format == 'xlsx')
check('检测到循环区域', len(wb._loop_tables) > 0 or len(wb.batch.loops) > 0,
      f'tables={len(wb._loop_tables)} loops={len(wb.batch.loops)}')

wb.batch.toggle()
check('已进入批量模式', wb.batch.mode is True and wb.batch.table is not None)
wb.batch.table.setItem(0, 0, QTableWidgetItem('第一行值'))
wb.batch.table.setItem(1, 0, QTableWidgetItem('第二行值'))
wb._remember()
saved_batch = FS.load_state(xb).get('batch') or []
check('批量数据已记忆', len(saved_batch) == 2, str(len(saved_batch)))

wb2 = M.TemplateFillWindow()
wb2.batch.mode = True          # 先进入批量模式，再加载模板
wb2._load_template(xb)
rows = wb2.batch.table.rowCount() if wb2.batch.table else 0
first = wb2.batch.table.item(0, 0).text() if rows and wb2.batch.table.item(0, 0) else ''
check('二次打开批量表格行数恢复', rows == 2, f'{rows} 行')
check('二次打开批量首格内容恢复', first == '第一行值', first)

# 循环数据记忆（Excel 模板含循环区域）
loops = wb2.batch.loops or {}
if loops:
    lv = list(loops.keys())[0]
    table = wb2._loop_tables.get(lv)
    if table:
        wb2._add_loop_row(lv, loops[lv]['fields'])
        if table.item(0, 0) is None:
            table.setItem(0, 0, QTableWidgetItem('循环值1'))
        table.item(0, 0).setText('循环值1')
        wb2._remember()
        wb3 = M.TemplateFillWindow()
        wb3._load_template(xb)
        t3 = wb3._loop_tables.get(lv)
        check('二次打开循环数据恢复',
              t3 is not None and t3.rowCount() >= 1 and t3.item(0, 0) is not None
              and t3.item(0, 0).text() == '循环值1',
              f"rows={t3.rowCount() if t3 else 0}")

# ── 7. 清除本模板记忆 ───────────────────────────────────────
w6 = M.TemplateFillWindow()
w6._load_template(tpl)
check('清除前有记忆', FS.has_state(tpl) is True)
w6._on_forget_template()
check('记忆文件已删除', FS.has_state(tpl) is False)
check('分组已重置', w6._custom_groups == [], str(w6._custom_groups))
check('导航保持 4 项', w6.nav.count() == 4, f'{w6.nav.count()}')
vals6 = [wd.toPlainText() for wd in w6._field_widgets.values()]
check('清除后界面内容为空', all(v == '' for v in vals6))

# ── 8. 项目文件保存 / 打开（含分组与批量数据） ───────────────
prj = os.path.join(tmp, 'demo.tplfill')
w7 = M.TemplateFillWindow()
w7._load_template(tpl)
w7._new_group('项目组')
w7._set_field_group('title', '项目组')
w7._reload_ui(rebuild_pages=True)
w7._field_widgets['title'].setPlainText('项目里的标题')
w7._remember()                # 相当于用户已生成过一次，值进入记忆
w7.batch.toggle()             # 真实路径切批量（字段值转入批量默认值）
check('切批量后批量表就绪', w7.batch.mode is True and w7.batch.table is not None)
w7.batch.table.setItem(0, 0, QTableWidgetItem('批量格A'))

_orig_save = M.QFileDialog.getSaveFileName
M.QFileDialog.getSaveFileName = staticmethod(lambda *a, **k: (prj, ''))
w7._save_project()
M.QFileDialog.getSaveFileName = _orig_save
check('项目文件已写出', os.path.exists(prj))

with open(prj, 'r', encoding='utf-8') as f:
    pdata = json.load(f)
check('项目含分组配置', pdata.get('custom_groups') == ['项目组'], str(pdata.get('custom_groups')))
check('项目含归组映射', pdata.get('field_assign', {}).get('title') == '项目组')

M.QFileDialog.getOpenFileName = staticmethod(lambda *a, **k: (prj, ''))
M.QMessageBox.question = staticmethod(_yes)
w8 = M.TemplateFillWindow()
w8._open_project()
M.QFileDialog.getOpenFileName = None
check('已打开项目', w8._project_path == prj, str(w8._project_path))
check('项目分组已恢复', w8._custom_groups == ['项目组'], str(w8._custom_groups))
check('项目批量模式已恢复', w8.batch.mode is True)
bt = w8.batch.table
check('项目批量数据已恢复（修复重建冲掉数据的旧问题）',
      bt is not None and bt.rowCount() >= 1 and bt.item(0, 0) is not None
      and bt.item(0, 0).text() == '批量格A',
      f"rows={bt.rowCount() if bt else 0} cell="
      f"{bt.item(0, 0).text() if bt and bt.item(0, 0) else ''}")
check('批量按钮状态已同步', '批量' in w8.btn_generate.text(), w8.btn_generate.text())
check('批量模式下无独立字段控件', len(w8._field_widgets) == 0)

# 切回单人模式：项目里的字段内容应回到输入框
w8.batch.toggle()
check('切回单人模式已建控件', len(w8._field_widgets) == n_fields, str(len(w8._field_widgets)))
check('切回单人模式恢复项目字段内容',
      w8._field_widgets.get('title') is not None
      and w8._field_widgets['title'].toPlainText() == '项目里的标题',
      w8._field_widgets['title'].toPlainText() if 'title' in w8._field_widgets else '(缺字段)')

# ── 输出 ────────────────────────────────────────────────────
passed = sum(1 for _, ok, _ in RESULTS if ok)
total = len(RESULTS)
lines = []
for name, ok, extra in RESULTS:
    mark = 'PASS' if ok else 'FAIL'
    lines.append(f'[{mark}] {name}' + (f'  ({extra})' if extra and not ok else ''))
lines.append('')
lines.append(f'合计 {passed} / {total} 通过')

report = '\n'.join(lines)
with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), '_v11_test_result.txt'),
          'w', encoding='utf-8') as f:
    f.write(report)
print(report)
shutil.rmtree(tmp, ignore_errors=True)
sys.exit(0 if passed == total else 1)
