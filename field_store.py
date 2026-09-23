"""
TemplateFill — 模板级状态存储

按模板文件保存「上次填写内容」「字段分组」「自定义标签」等，
再次打开同一模板时自动带出，免去重复输入。

存储位置：~/.TemplateFill/state/<模板名>_<路径哈希>.json

设计原则：
  - 任何读写异常都静默降级，绝不影响主流程（模板加载/生成）
  - 状态文件按模板路径哈希隔离，同名不同目录的模板互不干扰
  - 写入前做大包体裁剪，避免图片/超长文本把状态文件撑爆
"""

import os
import json
import hashlib

STORE_DIR = os.path.join(os.path.expanduser("~"), ".TemplateFill", "state")
STATE_VERSION = 1
MAX_STATE_BYTES = 4 * 1024 * 1024      # 单模板状态上限 4MB
MAX_BATCH_ROWS = 500                   # 批量数据最多记忆 500 行（防表格爆炸）


def _slug(template_path):
    """由模板绝对路径生成稳定指纹（文件名可读 + 路径哈希防重名）"""
    abspath = os.path.abspath(template_path)
    stem = os.path.splitext(os.path.basename(abspath))[0]
    # 去掉文件名里对文件系统不友好的字符
    safe = ''.join(ch if (ch.isalnum() or ch in '-_') else '_' for ch in stem)[:40]
    digest = hashlib.md5(abspath.lower().encode('utf-8', 'ignore')).hexdigest()[:10]
    return f"{safe}_{digest}.json"


def state_file(template_path):
    """返回该模板的状态文件路径"""
    return os.path.join(STORE_DIR, _slug(template_path))


def load_state(template_path):
    """读取模板状态；不存在或损坏时返回空 dict"""
    if not template_path:
        return {}
    path = state_file(template_path)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {}
        if data.get('version') != STATE_VERSION:
            # 版本不符：保留可识别的部分，其余忽略
            data = {k: v for k, v in data.items() if k != 'version'}
        return data
    except (IOError, OSError, json.JSONDecodeError, ValueError):
        return {}


def save_state(template_path, state):
    """写入模板状态；成功返回 True，任何异常返回 False"""
    if not template_path or not isinstance(state, dict):
        return False
    payload = dict(state)
    payload['version'] = STATE_VERSION
    payload['template'] = os.path.abspath(template_path)

    # 批量行数裁剪
    batch = payload.get('batch')
    if isinstance(batch, list) and len(batch) > MAX_BATCH_ROWS:
        payload['batch'] = batch[:MAX_BATCH_ROWS]

    try:
        try:
            text = json.dumps(payload, ensure_ascii=False, indent=1)
        except (TypeError, ValueError):
            # 有不可序列化对象（极少见）：退化成字符串
            text = json.dumps(payload, ensure_ascii=False, indent=1, default=str)
        if len(text.encode('utf-8')) > MAX_STATE_BYTES:
            # 体积超限：丢掉最占地方的部分（章节/批量），保住基本字段值
            payload.pop('sections', None)
            payload.pop('batch', None)
            text = json.dumps(payload, ensure_ascii=False, indent=1, default=str)

        os.makedirs(STORE_DIR, exist_ok=True)
        tmp = state_file(template_path) + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            f.write(text)
        os.replace(tmp, state_file(template_path))
        return True
    except (IOError, OSError, TypeError, ValueError):
        return False


def clear_state(template_path):
    """删除该模板的状态文件；成功删除返回 True"""
    if not template_path:
        return False
    path = state_file(template_path)
    try:
        if os.path.exists(path):
            os.remove(path)
            return True
    except (IOError, OSError):
        pass
    return False


def has_state(template_path):
    """该模板是否已有记忆"""
    return bool(template_path) and os.path.exists(state_file(template_path))


def state_summary(state):
    """状态概要：(已填字段数, 章节数, 批量行数)，用于状态栏提示"""
    if not isinstance(state, dict):
        return 0, 0, 0
    values = state.get('values') or {}
    filled = sum(1 for v in values.values() if str(v or '').strip())
    sections = state.get('sections') or []
    batch = state.get('batch') or []
    return filled, len(sections), len(batch)
