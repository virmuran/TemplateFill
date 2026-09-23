"""TemplateFill 版本号 — 唯一来源，所有地方从这里读

版本规范（三段式：主版本.次版本.修订号，如 1.0.0）：

    主版本 MAJOR   不兼容变更 —— 项目文件格式破坏性调整、技术栈更换
    次版本 MINOR   新增功能 —— 新增块类型/新功能，向后兼容
    修订号 PATCH   修 bug、文案、依赖与打包配置

硬性约束（本文件是唯一手写版本号的地方，其余文件由 build_release.py 自动同步）：
    · 必须严格三段纯数字，段内禁止前导零（写 1.0.1，不写 1.0.01）
    · 禁止把日期当版本号（1.0.20260920 这类，修订号位数超过 3 位即判非法）
    · 版本号只增不减；已发布过的号永不复用
"""

VERSION = "1.2.0"

#: 规范版本号：三段，无前导零，主/次 1~2 位、修订 1~3 位
VERSION_PATTERN = r"^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$"

__all__ = [
    "VERSION",
    "VERSION_PATTERN",
    "parse_version",
    "compare_versions",
    "is_valid_version",
]


def parse_version(version_str: str) -> tuple:
    """将版本号字符串解析为可比较的元组（宽容解析，不校验合法性）。"""
    parts = (version_str or "").strip().lstrip("vV").split(".")
    return tuple(int(p) if p.isdigit() else 0 for p in parts)


def compare_versions(current_str: str, latest_str: str) -> int:
    """比较两个版本号：1 = 后者更新，-1 = 后者更旧，0 = 相同。"""
    cur = parse_version(current_str)
    lat = parse_version(latest_str)
    max_len = max(len(cur), len(lat))
    cur = cur + (0,) * (max_len - len(cur))
    lat = lat + (0,) * (max_len - len(lat))
    if lat > cur:
        return 1
    if lat < cur:
        return -1
    return 0


def is_valid_version(version_str: str) -> bool:
    """校验是否符合版本号规范（发版前必须通过）。"""
    import re

    if not version_str or not isinstance(version_str, str):
        return False
    if not re.match(VERSION_PATTERN, version_str):
        return False
    major, minor, patch = (int(x) for x in version_str.split("."))
    if major < 1:
        return False
    if major > 99 or minor > 99:
        return False
    if patch > 999:      # 修订超过三位 = 混进了日期
        return False
    return True
