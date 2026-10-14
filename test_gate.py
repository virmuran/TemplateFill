#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""发版说明闸门的自测 —— 证明它**真会判失败**

一个永远不会失败的检查等于没有检查。本脚本把项目文件复制到临时目录，逐条制造
「该失败」的场景，确认闸门都报了 ✗。基线不动时它必须通过 —— 否则说明闸门坏在
另一半（见不到的问题会被判成没问题）。

在项目根目录运行：
    python test_gate.py
"""
import io
import os
import shutil
import subprocess
import sys
import tempfile

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = os.path.dirname(os.path.abspath(__file__))
GATE = os.path.join("tools", "check_release_notes.py")
FILES = ("README.md", "CHANGELOG.md", "version.py", "RELEASE_NOTES_TEMPLATE.md")

results = []


def check(name, passed, detail=""):
    results.append((name, passed, detail))
    print(f"  {'PASS' if passed else 'FAIL'}  {name}" + (f"   [{detail}]" if detail else ""))


def run_gate(cwd):
    r = subprocess.run([sys.executable, GATE], cwd=cwd, capture_output=True,
                       text=True, encoding="utf-8", errors="replace")
    fails = [ln.strip() for ln in r.stdout.splitlines() if ln.strip().startswith("✗")]
    return r.returncode, fails, r.stdout


def patch(path, old, new):
    text = io.open(path, encoding="utf-8").read()
    assert old in text, f"补丁未命中：{old[:40]!r}"
    io.open(path, "w", encoding="utf-8", newline="\n").write(text.replace(old, new))


def main():
    print("发版说明闸门自测")

    tmp = tempfile.mkdtemp(prefix="tf_gate_")
    os.makedirs(os.path.join(tmp, "tools"))
    for f in FILES:
        shutil.copy(os.path.join(ROOT, f), os.path.join(tmp, f))
    shutil.copy(os.path.join(ROOT, GATE), os.path.join(tmp, GATE))

    readme = os.path.join(tmp, "README.md")
    changelog = os.path.join(tmp, "CHANGELOG.md")

    # ── 基线：不动任何东西，必须通过 ──
    rc, fails, _ = run_gate(tmp)
    check("基线（原样）通过", rc == 0 and not fails, f"rc={rc}")

    # ── ① 版本徽章漂移（本项目真实发生过：徽章停在 0.0.3）──
    patch(readme, "badge/version-1.3.0-green", "badge/version-0.0.3-green")
    rc, fails, _ = run_gate(tmp)
    check("徽章与 version.py 不一致 → 判失败", rc == 1 and any("徽章" in f for f in fails))
    patch(readme, "badge/version-0.0.3-green", "badge/version-1.3.0-green")

    # ── ② README 里的入口链接被删掉 ──
    patch(readme, "[`CHANGELOG.md`](CHANGELOG.md)", "更新日志见仓库里的那个文件")
    rc, fails, _ = run_gate(tmp)
    check("README 没有 CHANGELOG 入口 → 判失败",
          rc == 1 and any("没有指向 CHANGELOG.md 的链接" in f for f in fails),
          f"rc={rc} fails={fails}")
    patch(readme, "更新日志见仓库里的那个文件", "[`CHANGELOG.md`](CHANGELOG.md)")

    # ── ③ CHANGELOG 里缺当前版本段 ──
    patch(changelog, "### v1.3.0 (2026-09-29)", "### v1.2.9 (2026-09-20)")
    rc, fails, _ = run_gate(tmp)
    check("CHANGELOG 缺当前版本段 → 判失败",
          rc == 1 and any("没有 v1.3.0 段" in f for f in fails))
    patch(changelog, "### v1.2.9 (2026-09-20)", "### v1.3.0 (2026-09-29)")

    # ── ④ CHANGELOG 文件整个不见了 ──
    os.remove(changelog)
    rc, fails, _ = run_gate(tmp)
    check("CHANGELOG.md 不存在 → 判失败",
          rc == 1 and any("找不到 CHANGELOG.md" in f for f in fails))

    # ── ⑤ 段边界：v1.3.0 段内自带 `## ` 小标题，不能被误截断 ──
    #     截断会让「段= Release 正文」这个约定悄悄失效（假通过最危险的一种）
    shutil.copy(os.path.join(ROOT, "CHANGELOG.md"), changelog)
    rc, _, out = run_gate(tmp)
    seg_lines = [ln for ln in out.splitlines() if "v1.3.0 段" in ln]
    n = int(seg_lines[0].split("（")[1].split(" 行")[0]) if seg_lines else 0
    check("v1.3.0 段落未被段内小标题截断", n >= 40, f"{n} 行")

    shutil.rmtree(tmp, ignore_errors=True)

    failed = [n for n, p, _ in results if not p]
    print(f"\n合计: {len(results) - len(failed)}/{len(results)} 通过")
    if failed:
        print("失败项: " + "; ".join(failed))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
