"""真机烟测：托盘图标能不能真的建起来（离屏环境测不到这一步）

离屏平台的 QSystemTrayIcon.isSystemTrayAvailable() 恒为 False，
所以托盘创建、图标落位、菜单挂载只能在本机默认平台下验证。

运行：python smoke_tray.py      （不加 QT_QPA_PLATFORM，窗口会闪 2 秒）
"""

import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from PySide6.QtCore import QTimer                     # noqa: E402
from PySide6.QtWidgets import QApplication, QSystemTrayIcon  # noqa: E402

import app_settings as AS                             # noqa: E402
import field_store as FS                              # noqa: E402

# 隔离：别动用户真实的设置与模板记忆
TMP = tempfile.mkdtemp(prefix='tf_smoke_')
AS.SETTINGS_DIR = os.path.join(TMP, 'app')
AS.SETTINGS_FILE = os.path.join(AS.SETTINGS_DIR, 'settings.json')
FS.STORE_DIR = os.path.join(TMP, 'state')

import main as M                                      # noqa: E402

app = QApplication(sys.argv)
app.setQuitOnLastWindowClosed(False)

print('platformName()          :', app.platformName())
print('isSystemTrayAvailable() :', QSystemTrayIcon.isSystemTrayAvailable())

win = M.TemplateFillWindow()
win.show()

print('_tray 已创建            :', win._tray is not None)
if win._tray is not None:
    print('_tray.isVisible()       :', win._tray.isVisible())
    print('contextMenu() is not None:', win._tray.contextMenu() is not None)
    print('toolTip()               :', win._tray.toolTip())
print('当前关闭行为            :', win._resolve_close_action())
print('菜单一级项              :',
      [a.text() for a in win._tray_menu.actions() if not a.isSeparator()])

RESULT = {}


def _finish():
    RESULT['tray_was_visible'] = bool(win._tray is not None and win._tray.isVisible())
    win._quit_app()
    RESULT['tray_after_quit'] = None if win._tray is None else win._tray.isVisible()


QTimer.singleShot(2000, _finish)
rc = app.exec()

print('---')
print('退出前托盘可见          :', RESULT.get('tray_was_visible'))
print('退出后托盘已摘（应 False）:', RESULT.get('tray_after_quit'))
print('事件循环返回码          :', rc)

ok = (win._tray is not None and RESULT.get('tray_after_quit') is False)
print('烟测结论                :', 'PASS' if ok else 'FAIL')
sys.exit(0 if ok else 1)
