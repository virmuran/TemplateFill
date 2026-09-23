@echo off
chcp 65001 >nul
echo Building TemplateFill (onedir + installer + portable zip)...

rem 独立打包环境：不存在则自动创建并装依赖（PySide6-Essentials 精简包）
if not exist ".venv-build\Scripts\python.exe" (
    echo Creating build venv...
    python -m venv .venv-build
    .venv-build\Scripts\python.exe -m pip install -i https://mirrors.aliyun.com/pypi/simple/ --upgrade pip
    .venv-build\Scripts\python.exe -m pip install -i https://mirrors.aliyun.com/pypi/simple/ "PySide6-Essentials>=6.5.0" docxtpl python-docx openpyxl lxml pyinstaller typing_extensions
    if errorlevel 1 (
        echo ERROR: dependency install failed!
        pause
        exit /b 1
    )
)

.venv-build\Scripts\python.exe build_release.py
pause
