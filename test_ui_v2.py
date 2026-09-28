"""TemplateFill 界面（左侧导航 + 字段单页）离屏回归测试

验证：四页骨架、字段单列排布与小节标题、正文页、生成导出页、
输出路径解析、端到端生成、批量模式切换。运行：
    QT_QPA_PLATFORM=offscreen python test_ui_v2.py
"""

import os
import sys
import tempfile

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication  # noqa: E402

app = QApplication([])

import main as M  # noqa: E402
from doc_filler import create_sample_template  # noqa: E402
from field_widgets import AutoGrowTextEdit  # noqa: E402

# 护栏：关闭窗口会走「每次询问」对话框，离屏下 exec() 永久阻塞且不报错
# （整套测试会静默卡住，故直接钉死成"取消"）
M.CloseChoiceDialog.exec = lambda self: (setattr(self, 'choice', 'cancel'), 0)[1]

RESULTS = []


def check(name, cond, extra=''):
    RESULTS.append((name, bool(cond), extra))


tmp = tempfile.mkdtemp(prefix='tf_ui_v2_')
tpl = os.path.join(tmp, 'demo_report.docx')
create_sample_template(tpl)

# ── 1. 骨架 ────────────────────────────────────────────────
w = M.TemplateFillWindow()
check('窗口创建成功', w is not None)
check('导航 4 项', w.nav.count() == 4, f'实际 {w.nav.count()}')
check('分区页 4 个', w.stack.count() == 4, f'实际 {w.stack.count()}')
check('form_layout 指向字段页', w.form_layout is w._page_layouts[0])
check('初始停在字段页', w.stack.currentIndex() == 0)
# 启动会自动加载 ~/TemplateFill/templates/sample_report.docx（若存在）
auto_loaded = w._template_path is not None
check('字段控件数与自动加载状态一致',
      auto_loaded == (len(w._field_widgets) > 0),
      f'auto_loaded={auto_loaded}, fields={len(w._field_widgets)}')
check('生成按钮状态与模板状态一致',
      w.btn_generate.isEnabled() == auto_loaded,
      f'enabled={w.btn_generate.isEnabled()}, auto_loaded={auto_loaded}')
check('顶部按钮已精简', hasattr(w, 'btn_more') and not hasattr(w, 'btn_sample'))

# ── 2. 加载模板 ─────────────────────────────────────────────
w._load_template(tpl)
check('模板路径已记录', w._template_path == tpl)
check('生成按钮已启用', w.btn_generate.isEnabled())
n_fields = len(w._field_widgets)
check('字段控件已构建', n_fields > 0, f'{n_fields} 个')
check('检测到正文标记', w._has_body_marker is True)
check('正文页有章节编辑器', w._section_editor is not None)

nav_texts = [w.nav.item(i).text() for i in range(4)]
check('导航显示计数', any('·' in t for t in nav_texts), str(nav_texts))
check('模板信息显示在顶栏', 'demo_report.docx' in w.lbl_template.text(),
      w.lbl_template.text())

# 字段页：全部字段单列排布 + 分组小节标题
priority = ['title', 'project_name', 'department', 'author', 'date',
            'doc_number', 'company', 'reviewer', 'phone']
info_tags = [t for t in priority if t in w.batch.tags]
other_tags = [t for t in w.batch.tags if t not in info_tags]
check('字段页承载全部字段',
      len(w._page_field_tags.get('fields', [])) == n_fields,
      f"{len(w._page_field_tags.get('fields', []))} vs {n_fields}")
check('字段页导航显示字段总数', w.nav.item(0).text().endswith(str(n_fields)),
      w.nav.item(0).text())
sections = w._group_sections(w._visible_tags())
check('字段按分组切成小节', len(sections) >= 1
      and all(tags for _, tags in sections),
      str([(n, len(t)) for n, t in sections]))
check('文档信息小节包含优先字段',
      len(info_tags) == 0 or any(n == w.INFO_GROUP for n, _ in sections),
      str([n for n, _ in sections]))
check('字段顺序与模板解析顺序一致', w._visible_tags() == list(w.batch.tags),
      str(w._visible_tags()))
check('字段控件为自动增高输入框',
      all(isinstance(x, AutoGrowTextEdit) for x in w._field_widgets.values()))
check('标题列宽在限制区间内',
      w.FIELD_LABEL_MIN_WIDTH <= w._label_width <= w.FIELD_LABEL_MAX_WIDTH,
      f'{w._label_width}')

# ── 3. 填充内容 + 摘要 ──────────────────────────────────────
filled = 0
for tag, widget in w._field_widgets.items():
    if filled >= 3:
        break
    widget.setPlainText(f'测试{tag}')
    filled += 1
w._update_summary()
check('摘要显示已填数量', f'{filled} / {n_fields}' in w.lbl_summary.text(),
      w.lbl_summary.text())

# ── 4. 输出路径解析 ─────────────────────────────────────────
check('未设置时走默认输出路径', w._resolve_output_path() is None)

out_dir = os.path.join(tmp, 'out', 'nested')
w.out_dir_edit.setText(out_dir)
w.out_name_edit.setText('我的报告')
w._update_filename_preview()
out_path = w._resolve_output_path()
check('自定义输出路径已生成', bool(out_path) and out_path.startswith(out_dir),
      str(out_path))
check('目录自动创建', os.path.isdir(out_dir))
check('文件名含前缀', '我的报告' in os.path.basename(out_path or ''))
check('文件名含时间戳', any(ch.isdigit() for ch in os.path.basename(out_path or '')))
check('扩展名正确', (out_path or '').endswith('.docx'))
check('文件名预览已更新', '我的报告' in w.lbl_filename_preview.text(),
      w.lbl_filename_preview.text())

# ── 5. 端到端生成 ───────────────────────────────────────────
w._on_generate()
check('生成后记录了输出路径', bool(w._last_output), str(w._last_output))
check('输出文件已落盘', os.path.exists(w._last_output or ''))
check('生成后跳转到生成导出页', w.stack.currentIndex() == 3,
      f'当前页 {w.stack.currentIndex()}')
check('最近输出区已刷新', os.path.basename(w._last_output or '') == w.lbl_recent.text(),
      w.lbl_recent.text())
check('最近输出按钮已启用', w.btn_recent_open.isEnabled())
check('状态栏有生成提示', '已生成' in w.status_bar.currentMessage(),
      w.status_bar.currentMessage())

# ── 6. 默认路径生成（不设自定义） ────────────────────────────
w.out_dir_edit.setText('')
w.out_name_edit.setText('')
w2_path = w._resolve_output_path()
check('清空设置后回到默认', w2_path is None)

# ── 7. 批量模式切换 ─────────────────────────────────────────
w.batch.toggle()
check('已进入批量模式', w.batch.mode is True)
check('批量表格已建立', w.batch.table is not None)
check('批量表列数 = 字段数', w.batch.table.columnCount() == n_fields,
      f'{w.batch.table.columnCount()} vs {n_fields}')
check('批量表有初始行', w.batch.table.rowCount() >= 3, f'{w.batch.table.rowCount()}')
check('自动跳到字段页', w.stack.currentIndex() == 0)
check('批量按钮文字已切换', '批量' in w.btn_batch.text(), w.btn_batch.text())
check('生成按钮文字已切换', w.btn_generate.text() == '批量生成', w.btn_generate.text())
w._update_summary()
check('批量摘要显示行数', '行' in w.lbl_summary.text(), w.lbl_summary.text())

w.batch.toggle()
check('已切回单人模式', w.batch.mode is False)
check('单人模式重建字段控件', len(w._field_widgets) == n_fields,
      f'{len(w._field_widgets)} vs {n_fields}')
check('单人模式导航恢复', '批量' not in w.btn_generate.text())

# ── 8. Excel 模板路径（无 body marker） ──────────────────────
from xlsx_filler import create_sample_template as create_xlsx_sample  # noqa: E402
xls = os.path.join(tmp, 'demo_sheet.xlsx')
create_xlsx_sample(xls)
w._load_template(xls)
check('Excel 模板已加载', w._template_format == 'xlsx')
check('Excel 无正文标记', w._has_body_marker is False)
check('Excel 正文页提示未启用', w._section_editor is None)
check('Excel 字段可填', len(w._field_widgets) > 0, f'{len(w._field_widgets)} 个')

# ── 9. 空状态与分页切换 ─────────────────────────────────────
w3 = M.TemplateFillWindow()
w3._render_empty_state()
check('空状态：导航无计数',
      all('·' not in w3.nav.item(i).text() for i in range(4)),
      str([w3.nav.item(i).text() for i in range(4)]))
check('空状态：每页都有占位内容',
      all(w3._page_layouts[i].count() > 0 for i in range(4)),
      str([w3._page_layouts[i].count() for i in range(4)]))
w3._goto_page(3)
check('goto_page 同步导航与堆栈',
      w3.stack.currentIndex() == 3 and w3.nav.currentRow() == 3,
      f'stack={w3.stack.currentIndex()} nav={w3.nav.currentRow()}')
w3._goto_page(99)
check('goto_page 越界安全', w3.stack.currentIndex() == 3)

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
out_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_ui_test_result.txt')
with open(out_file, 'w', encoding='utf-8') as f:
    f.write(report)
print(report)
sys.exit(0 if passed == total else 1)
