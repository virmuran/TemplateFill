"""TemplateFill 托盘图标与关闭行为回归测试（v1.3.0）

覆盖：
  1. app_settings：关闭行为偏好的读/写/落盘/非法值拒绝/文件损坏退回默认
  2. 托盘构建：图标创建、右键菜单结构、单选互斥勾选、菜单被长期持有、动作都连着槽
  3. 收托盘：窗口隐藏、进程不退、填写内容已落盘、首次提示只弹一次
  4. 唤回：单击/双击唤回、右键不误唤回、最大化状态不丢
  5. 无托盘降级：不隐藏窗口，转为真退出
  6. closeEvent 各分支：每次询问（取消 / 记住托盘 / 不记住退出）、已记住后不再询问
  7. 源码断言：没有 self.close()、QApplication.quit() 只在 _quit_app 里、
     main() 设了 setQuitOnLastWindowClosed(False)
  8. 真跑事件循环：点 X 选「直接退出」后进程确实结束（无幽灵进程），
     并带一条反向对照证明这个断言真有判别力

运行：QT_QPA_PLATFORM=offscreen python test_v13.py
"""

import ast
import os
import sys
import json
import shutil
import tempfile

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from PySide6.QtCore import Qt, QTimer                  # noqa: E402
from PySide6.QtWidgets import (                        # noqa: E402
    QApplication, QMessageBox, QSystemTrayIcon,
)

app = QApplication([])
# 与 main() 一致的硬前提：收进托盘时"最后一个窗口关闭"不代表退出进程。
# 这里不设的话，「真退出」那条用例会假通过（少一句 quit() 也看不出来）。
app.setQuitOnLastWindowClosed(False)

import app_settings as AS                             # noqa: E402
import field_store as FS                              # noqa: E402

# ── 隔离：绝不碰用户真实的 ~/.TemplateFill ──────────────────
TMP = tempfile.mkdtemp(prefix='tf_v13_')
AS.SETTINGS_DIR = os.path.join(TMP, 'app')
AS.SETTINGS_FILE = os.path.join(AS.SETTINGS_DIR, 'settings.json')
FS.STORE_DIR = os.path.join(TMP, 'state')

import main as M                                      # noqa: E402
from doc_filler import create_sample_template         # noqa: E402

M.TemplateFillWindow.TEMPLATES_DIR = os.path.join(TMP, 'no_templates')

RESULTS = []


def check(name, cond, extra=''):
    RESULTS.append((name, bool(cond), extra))


# ── 护栏：绝不让询问对话框在离屏跑到（exec() 会永久阻塞且不报错）──
M.CloseChoiceDialog.exec = lambda self: (setattr(self, 'choice', 'cancel'), 0)[1]

# 模态框替身（离屏下真弹会卡死）
QMessageBox.critical = staticmethod(lambda *a, **k: print('  [critical]', a[1:3]))
QMessageBox.information = staticmethod(lambda *a, **k: 1)
QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.StandardButton.Yes)

tpl = os.path.join(TMP, 'demo.docx')
create_sample_template(tpl)

CLOSE_KEYS = dict(M.TemplateFillWindow.CLOSE_LABELS)


def _read_settings_file():
    if not os.path.exists(AS.SETTINGS_FILE):
        return {}
    with open(AS.SETTINGS_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)


def _menu_leaves(menu):
    """菜单里的叶子动作文本（不含分隔符与子菜单入口）"""
    return [a.text() for a in menu.actions() if not a.isSeparator() and a.menu() is None]


# ── 源码断言用的 AST 工具（先定义，后面各段落都能用）─────────
SRC = open(os.path.join(HERE, 'main.py'), encoding='utf-8').read()
TREE = ast.parse(SRC)


def _calls_in(func_name):
    """某函数体里出现过的调用表达式

    用 AST 取调用表达式，而不是对源码做子串搜索 —— docstring 里提到的代码
    是字符串字面量，AST 天然不算调用，不会把「注释里写了」误判成「真的写了」。
    """
    for node in ast.walk(TREE):
        if isinstance(node, ast.FunctionDef) and node.name == func_name:
            out = []
            for n in ast.walk(node):
                if isinstance(n, ast.Call):
                    args = ', '.join(ast.unparse(a) for a in n.args)
                    out.append(f'{ast.unparse(n.func)}({args})')
            return out
    return None


# ══ 1. app_settings ═════════════════════════════════════════

check('出厂默认关闭行为是「每次询问」', AS.get_close_action() == 'ask',
      AS.get_close_action())
check('默认时设置文件尚未创建', not os.path.exists(AS.SETTINGS_FILE))

check('写入偏好成功', AS.set_close_action('tray') is True)
check('偏好真的落盘（重读文件核对）',
      _read_settings_file().get('close_action') == 'tray',
      str(_read_settings_file()))

check('非法取值被拒', AS.set_close_action('bogus') is False)
check('非法取值不污染文件', _read_settings_file().get('close_action') == 'tray')

AS.set_close_action('quit')
check('偏好可改为直接退出', AS.get_close_action() == 'quit')
AS.set_close_action('ask')

with open(AS.SETTINGS_FILE, 'w', encoding='utf-8') as f:
    f.write('{ 这不是合法 JSON ')
check('设置文件损坏时退回默认值', AS.get_close_action() == 'ask',
      AS.get_close_action())
check('损坏时 load_settings 返回空 dict', AS.load_settings() == {})

with open(AS.SETTINGS_FILE, 'w', encoding='utf-8') as f:
    json.dump({'close_action': '不认识的取值'}, f, ensure_ascii=False)
check('未知取值退回默认值', AS.get_close_action() == 'ask')

AS.set_close_action('ask')
check('托盘菜单里有三种关闭行为', set(CLOSE_KEYS) == {'ask', 'tray', 'quit'},
      str(CLOSE_KEYS))


# ══ 2. 托盘构建与菜单 ═══════════════════════════════════════

w = M.TemplateFillWindow()
w.show()

check('离屏环境本身没有托盘（降级前提成立）', w._tray is None)
check('无托盘时 _setup_tray() 返回 None', w._setup_tray() is None)

w._tray_available = lambda: True          # 覆写后能造出真实 QSystemTrayIcon
tray = w._setup_tray()
check('覆写可用性后托盘图标创建成功', tray is not None and w._tray is tray)
check('托盘图标挂上了右键菜单', tray is not None and tray.contextMenu() is not None)
check('托盘菜单被长期持有（防 wrapper 回收）',
      w._tray_menu is not None and w._tray_menu is tray.contextMenu())
check('托盘提示含软件名与版本',
      'TemplateFill' in tray.toolTip() and (M.VERSION or '') in tray.toolTip(),
      tray.toolTip())
check('未加载模板时托盘提示为通用文案',
      '模板文档生成器' in tray.toolTip(), tray.toolTip())

menu = w._tray_menu
leaves = _menu_leaves(menu)
check('菜单一级项正确', leaves == ['显示主窗口', '打开模板...', '退出 TemplateFill'],
      str(leaves))
sub_actions = [a for a in menu.actions() if a.menu() is not None]
check('菜单里有「关闭窗口时」子菜单',
      len(sub_actions) == 1 and sub_actions[0].text() == '关闭窗口时',
      str([a.text() for a in sub_actions]))
behavior = sub_actions[0].menu()
check('子菜单列出三种行为',
      [a.text() for a in behavior.actions()]
      == [label for _, label in M.TemplateFillWindow.CLOSE_LABELS],
      str([a.text() for a in behavior.actions()]))
check('子菜单三项都可勾选（单选组）',
      all(a.isCheckable() for a in behavior.actions()))
check('当前偏好对应的项已勾上', w._close_actions['ask'].isChecked())
check('单选互斥成立', sum(1 for a in behavior.actions() if a.isChecked()) == 1)
check('单选组对象在窗口实例上留了引用', w._close_group is not None)

# 每个动作都真的连着槽 —— 直接触发看效果
w._close_actions['tray'].trigger()
check('触发「最小化到托盘」项 → 偏好跟着改', AS.get_close_action() == 'tray',
      AS.get_close_action())
check('触发菜单项后勾选同步', w._close_actions['tray'].isChecked())
check('切换偏好后旧项取消勾选',
      sum(1 for a in behavior.actions() if a.isChecked()) == 1)

called = {}
_orig_quit_app = w._quit_app
w._quit_app = lambda: called.setdefault('quit', True)
menu.actions()[-1].trigger()                       # 「退出 TemplateFill」
check('「退出 TemplateFill」已连到 _quit_app', called.get('quit') is True)
w._quit_app = _orig_quit_app

called.clear()
_orig_open = w._on_open_template
w._on_open_template = lambda: called.setdefault('open', True)
w.hide()
menu.actions()[1].trigger()                        # 「打开模板...」
check('「打开模板...」已连到打开流程', called.get('open') is True)
check('从托盘打开模板会先唤回窗口（避免对话框没落点）', not w.isHidden())
w._on_open_template = _orig_open

w.hide()
menu.actions()[0].trigger()                        # 「显示主窗口」
check('「显示主窗口」已连到唤回', not w.isHidden())


# ══ 3. 收托盘 / 唤回 ════════════════════════════════════════

w._load_template(tpl)
w._field_widgets['title'].setPlainText('托盘测试标题')
w.show()
w._set_close_action('tray')
check('托盘提示已带上模板名', 'demo.docx' in w._tray.toolTip(), w._tray.toolTip())
check('收托盘前窗口是可见的（前置条件）', not w.isHidden())

w.close()                                   # 点 X → 收进托盘
check('收托盘后窗口被隐藏', w.isHidden())
check('收托盘后不是真退出', w._really_quit is False)
check('收托盘后托盘对象仍在（程序继续后台运行）', w._tray is not None)
check('收托盘时把填写内容落了盘', FS.has_state(tpl))
_saved = FS.load_state(tpl)
check('落盘内容就是刚填的值',
      (_saved.get('values') or {}).get('title') == '托盘测试标题',
      str((_saved.get('values') or {}).get('title')))
check('首次收托盘弹了一次提示', w._tray_tip_shown is True)

w._on_tray_activated(QSystemTrayIcon.ActivationReason.Trigger)
check('单击托盘唤回窗口', not w.isHidden())

w.hide()
w._on_tray_activated(QSystemTrayIcon.ActivationReason.Context)
check('右键托盘不误唤回', w.isHidden())

w._on_tray_activated(QSystemTrayIcon.ActivationReason.DoubleClick)
check('双击托盘唤回窗口', not w.isHidden())

w.setWindowState(Qt.WindowState.WindowMaximized)
w._minimize_to_tray()
check('收托盘前记录了窗口状态（最大化不丢）',
      bool(w._pre_tray_state & Qt.WindowState.WindowMaximized),
      str(w._pre_tray_state))
w._restore_from_tray()
check('唤回窗口', not w.isHidden())
check('唤回用 setWindowState 还原、而不是 showNormal',
      'self.setWindowState(self._pre_tray_state)' in (_calls_in('_restore_from_tray') or [])
      and not any('showNormal' in c for c in (_calls_in('_restore_from_tray') or [])),
      str(_calls_in('_restore_from_tray')))


# ══ 4. closeEvent 各分支 ════════════════════════════════════

# ① 每次询问 → 用户选「取消」
wc = M.TemplateFillWindow()
wc.show()
wc._set_close_action('ask')
wc.close()
check('询问后选取消 → 窗口留在界面上', not wc.isHidden())
check('取消不会改动偏好', AS.get_close_action() == 'ask')


def _drive_choice(choice, remember):
    """把询问对话框的 exec 换成直接给答案（离屏下真 exec 会永久阻塞）"""
    def fake_exec(self):
        self.choice = choice
        self.remember = remember
        return 1
    return fake_exec


# ② 每次询问 → 选「最小化到托盘」+ 记住
wc._tray_available = lambda: True
wc._setup_tray()
M.CloseChoiceDialog.exec = _drive_choice('tray', True)
wc.close()
check('询问后选托盘 → 窗口隐藏', wc.isHidden())
check('勾了记住 → 偏好写成 tray', AS.get_close_action() == 'tray')
check('写偏好后菜单勾选同步', wc._close_actions['tray'].isChecked())

# ③ 每次询问 → 选「直接退出」但不记住
ws = M.TemplateFillWindow()
ws.show()
ws._set_close_action('ask')
ws._tray_available = lambda: True
ws._setup_tray()
M.CloseChoiceDialog.exec = _drive_choice('quit', False)
ws.close()
check('询问后选退出 → 判定为真退出', ws._really_quit is True)
check('未记住时偏好不变', AS.get_close_action() == 'ask')

# ④ 已记住「直接退出」→ 点 X 直接走，不再弹询问
wq = M.TemplateFillWindow()
wq.show()
wq._tray_available = lambda: True
wq._setup_tray()
wq._set_close_action('quit')
wq._really_quit = False
wq.close()
check('已记住后点 X 不再询问（询问会改偏好，偏好应保持 quit）',
      AS.get_close_action() == 'quit', AS.get_close_action())
check('已记住后点 X 直接真退出',
      wq._really_quit is True and wq.isHidden())

M.CloseChoiceDialog.exec = lambda self: (setattr(self, 'choice', 'cancel'), 0)[1]


# ══ 5. 无托盘降级 ═══════════════════════════════════════════

wn = M.TemplateFillWindow()               # 离屏下天然没有托盘
wn.show()
check('无托盘环境 _tray 为 None', wn._tray is None)
result = wn._minimize_to_tray()
check('无托盘时不隐藏窗口（否则找不回来）', not wn.isHidden())
check('无托盘时降级为真退出', result is False and wn._really_quit is True)


# ══ 6. 源码断言 ═════════════════════════════════════════════

all_calls = [ast.unparse(n.func) for n in ast.walk(TREE)
             if isinstance(n, ast.Call)]
check('全项目没有 self.close() 调用（会走偏好判定，与「退出」意图相反）',
      'self.close' not in all_calls,
      str([c for c in all_calls if 'close' in c]))

quit_owners = []
for node in ast.walk(TREE):
    if isinstance(node, ast.FunctionDef):
        for n in ast.walk(node):
            if isinstance(n, ast.Call) and ast.unparse(n.func) == 'QApplication.quit':
                quit_owners.append(node.name)
check('结束进程只有 _quit_app 一处出口',
      quit_owners == ['_quit_app'], str(quit_owners))
check('_quit_app 里确实调了 QApplication.quit()',
      'QApplication.quit()' in (_calls_in('_quit_app') or []))
check('_quit_app 里摘掉了托盘图标',
      'self._tray.hide()' in (_calls_in('_shutdown') or []),
      str(_calls_in('_shutdown')))
check('main() 设了 setQuitOnLastWindowClosed(False)',
      'app.setQuitOnLastWindowClosed(False)' in (_calls_in('main') or []),
      str(_calls_in('main')))
close_calls = _calls_in('closeEvent') or []
check('closeEvent 的「直接退出」分支走 _quit_app()',
      'self._quit_app()' in close_calls, str(close_calls))
check('closeEvent 里没有自己另写一份退出逻辑',
      'QApplication.quit()' not in close_calls, str(close_calls))


# ══ 7. 真跑事件循环：点 X 选「直接退出」后进程必须结束 ═══════

# ⚠ 顺序不能反：QApplication.quit() 在 exec() 尚未启动时是空操作。
#    必须「先 exec、再由定时器触发点 X」，否则修没修都会判成幽灵（假阳性）。
def _run_close_zombie_check(win, timeout_ms=1200):
    """在**真事件循环**里点 X，返回是否出现幽灵进程（进程还活着）

    超时器用完必须 stop()：QTimer.singleShot 没法取消，上一轮遗留的超时器
    会在下一轮事件循环里触发，把下一轮提前掐断（实测踩过，反向对照因此假失败）。
    """
    seen = {}
    watchdog = QTimer()
    watchdog.setSingleShot(True)
    watchdog.setInterval(timeout_ms)
    watchdog.timeout.connect(lambda: (seen.setdefault('zombie', True), app.quit()))
    QTimer.singleShot(0, win.close)         # 循环已启动后才点 X
    watchdog.start()
    app.exec()
    watchdog.stop()
    return 'zombie' in seen


wz = M.TemplateFillWindow()
wz._tray_available = lambda: True
wz._setup_tray()
wz.show()
wz._set_close_action('quit')
check('点 X 选「直接退出」后进程真的结束（无幽灵进程）',
      _run_close_zombie_check(wz) is False)

# 反向对照：把 _quit_app 换回旧实现（只收尾、不 quit），本用例必须能判出问题
# —— 证明上面那条断言真有判别力，不是永远为真
wz2 = M.TemplateFillWindow()
wz2._tray_available = lambda: True
wz2._setup_tray()
wz2.show()
wz2._set_close_action('quit')
wz2._quit_app = lambda: (setattr(wz2, '_really_quit', True), wz2._shutdown())
check('反向对照：旧写法（只 accept 不 quit）确实会被判为幽灵进程',
      _run_close_zombie_check(wz2) is True)

AS.set_close_action('ask')      # 收尾，别把测试状态留在设置文件里


# ══ 输出 ════════════════════════════════════════════════════
passed = sum(1 for _, ok, _ in RESULTS if ok)
total = len(RESULTS)
lines = []
for name, ok, extra in RESULTS:
    mark = 'PASS' if ok else 'FAIL'
    lines.append(f'[{mark}] {name}' + (f'  ({extra})' if extra and not ok else ''))
lines.append('')
lines.append(f'合计 {passed} / {total} 通过')

report = '\n'.join(lines)
with open(os.path.join(HERE, '_v13_test_result.txt'), 'w', encoding='utf-8') as f:
    f.write(report)
print(report)
shutil.rmtree(TMP, ignore_errors=True)
sys.exit(0 if passed == total else 1)
