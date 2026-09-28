"""
TemplateFill — 应用级设置存储（与具体模板无关）

存储位置：~/.TemplateFill/settings.json

目前只有一项：
  close_action —— 点窗口「X」时的行为，取值 ask / tray / quit（出厂 ask）

与 field_store（模板级记忆）分开存放：这份是"整台机器上的软件偏好"，
换模板、清模板记忆都不该影响它。

设计原则同 field_store：任何读写异常静默降级，绝不影响主流程。
"""

import os
import json

SETTINGS_DIR = os.path.join(os.path.expanduser("~"), ".TemplateFill")
SETTINGS_FILE = os.path.join(SETTINGS_DIR, "settings.json")

# 关闭窗口时的三种行为（key → 菜单显示名，顺序即菜单顺序）
CLOSE_ACTIONS = (("ask", "每次询问"),
                 ("tray", "最小化到托盘"),
                 ("quit", "直接退出"))
DEFAULT_CLOSE_ACTION = "ask"


# ── 通用读写 ────────────────────────────────────────────────

def load_settings():
    """读取应用设置；文件不存在或损坏时返回空 dict"""
    if not os.path.exists(SETTINGS_FILE):
        return {}
    try:
        with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (IOError, OSError, json.JSONDecodeError, ValueError):
        return {}


def save_settings(settings):
    """写入应用设置；成功返回 True，任何异常返回 False"""
    if not isinstance(settings, dict):
        return False
    try:
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        text = json.dumps(settings, ensure_ascii=False, indent=1)
        tmp = SETTINGS_FILE + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            f.write(text)
        os.replace(tmp, SETTINGS_FILE)      # 原子落盘，避免写一半被中断
        return True
    except (IOError, OSError, TypeError, ValueError):
        return False


def get_setting(key, default=None):
    """取单项设置"""
    return load_settings().get(key, default)


def set_setting(key, value):
    """设置单项（读-改-写）"""
    settings = load_settings()
    settings[key] = value
    return save_settings(settings)


# ── 关闭窗口时的行为 ────────────────────────────────────────

def get_close_action():
    """当前生效的关闭行为；未知取值一律退回「每次询问」"""
    action = load_settings().get('close_action', DEFAULT_CLOSE_ACTION)
    return action if action in dict(CLOSE_ACTIONS) else DEFAULT_CLOSE_ACTION


def set_close_action(action):
    """写入关闭行为偏好；非法取值返回 False 且不改动文件"""
    if action not in dict(CLOSE_ACTIONS):
        return False
    return set_setting('close_action', action)
