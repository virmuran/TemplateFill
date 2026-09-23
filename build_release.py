"""TemplateFill 一键发版脚本（复刻 ChemCal 打包方式）
用法:  .venv-build/Scripts/python.exe build_release.py [--skip-build]
流程:  校验版本号 -> 同步 .iss 版本号 -> PyInstaller onedir -> Inno Setup 安装包 -> 便携 zip
产物:  installer/TemplateFill_<版号>_setup.exe
       dist/TemplateFill_<版号>_portable.zip

升版本号请手改 version.py（本项目暂无 bump_version.py）：
    · 修 bug 升修订号 1.0.0 -> 1.0.1
    · 加功能升次版本 1.0.0 -> 1.1.0
版本规范见 version.py 文件头。
"""
import os
import re
import subprocess
import sys
import zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(ROOT)

ISCC = r"C:\Program Files\Inno Setup 7\ISCC.exe"
BUILD_PY = os.path.join(ROOT, ".venv-build", "Scripts", "python.exe")


def step(msg):
    print(f"\n=== {msg} ===", flush=True)


def run(cmd):
    print("  $", " ".join(cmd), flush=True)
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        print(r.stdout[-2000:])
        print(r.stderr[-2000:])
        sys.exit(f"步骤失败: {cmd[0]}")


def version_gate(ver: str):
    """发版前的版本号闸门 —— 不合规直接中断打包。"""
    step("校验版本号")
    from version import is_valid_version, compare_versions

    if not is_valid_version(ver):
        sys.exit(f"✗ 版本号 '{ver}' 不符合规范（须为三段纯数字 X.Y.Z，禁止日期/四段/前导零）。"
                 f"规范见 version.py 文件头")

    try:
        r = subprocess.run(["git", "tag", "-l", "v*"], cwd=ROOT,
                           capture_output=True, text=True, timeout=15)
        tags = [t.strip().lstrip("v") for t in r.stdout.splitlines() if t.strip()]
        tags = [t for t in tags if t and t[0].isdigit()]
    except Exception:
        tags = []

    if not tags:
        print(f"  v{ver} 格式合法；无历史 tag 可比对")
        return

    latest = sorted(tags, key=lambda t: tuple(int(x) if x.isdigit() else 0
                                              for x in t.split(".")))[-1]
    if ver in tags:
        sys.exit(f"✗ git 中已存在 tag v{ver}（版本号不可复用），请升号再打包")
    if compare_versions(latest, ver) <= 0:
        sys.exit(f"✗ 版本号没有前进：最新 tag 是 v{latest}，当前 version.py 是 v{ver}")
    print(f"  v{ver} 格式合法，且高于最新 tag v{latest}")


def main():
    # 0) 版本号
    sys.path.insert(0, ROOT)
    from version import VERSION  # noqa: E402
    ver = VERSION
    print(f"TemplateFill v{ver}")
    version_gate(ver)

    # 1) 同步 TemplateFill.iss 的 AppVersion（防两处不同步）
    step("同步 TemplateFill.iss 版本号")
    iss = os.path.join(ROOT, "TemplateFill.iss")
    text = open(iss, encoding="utf-8").read()
    new = re.sub(r'#define MyAppVersion "[^"]*"', f'#define MyAppVersion "{ver}"', text)
    if new != text:
        open(iss, "w", encoding="utf-8").write(new)
        print(f"  TemplateFill.iss AppVersion -> {ver}")
    else:
        print(f"  已一致: {ver}")

    # 2) PyInstaller onedir（用独立打包环境）
    if "--skip-build" not in sys.argv:
        step("PyInstaller 打包（约 2~3 分钟）")
        run([BUILD_PY, "-m", "PyInstaller", "TemplateFill.spec", "--noconfirm"])

    # 3) Inno Setup 安装包
    step("Inno Setup 编译安装包")
    run([ISCC, "TemplateFill.iss"])

    # 4) 便携 zip（ZIP_LZMA；勿用 bsdtar 默认 deflate，体积差 3 倍）
    step("制作便携 zip")
    src = os.path.join(ROOT, "dist", "TemplateFill")
    out = os.path.join(ROOT, "dist", f"TemplateFill_{ver}_portable.zip")
    if os.path.exists(out):
        os.remove(out)
    n = 0
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_LZMA, compresslevel=9) as z:
        for root, _dirs, files in os.walk(src):
            for f in files:
                p = os.path.join(root, f)
                z.write(p, os.path.relpath(p, os.path.join(ROOT, "dist")))
                n += 1

    # 5) 汇总
    setup = os.path.join(ROOT, "installer", f"TemplateFill_{ver}_setup.exe")
    mb = lambda p: round(os.path.getsize(p) / 1048576, 1)
    print("\n=== 打包完成 ===")
    print(f"  安装包  {setup}  ({mb(setup)} MB)")
    print(f"  便携包  {out}  ({mb(out)} MB, {n} 文件)")
    print("\n分发话术:")
    print("  · 安装包：双击下一步到底，自动建开始菜单 + 桌面快捷方式")
    print("  · 便携包：解压即用，零联网")


if __name__ == "__main__":
    main()
