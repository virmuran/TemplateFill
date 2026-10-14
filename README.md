<!-- markdownlint-disable -->

<div align="center">

<h1>TemplateFill</h1>

<br>
<div>
    <img alt="Python" src="https://img.shields.io/badge/Python-3.8+-%233776AB?logo=python">
    <img alt="platform" src="https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-blueviolet">
</div>
<div>
    <img alt="license" src="https://img.shields.io/badge/license-MIT-blue">
    <img alt="version" src="https://img.shields.io/badge/version-1.3.0-green">
</div>
<br>

模板驱动的 Office 文档生成器

一键将模板中的 `{{变量}}` 替换为实际内容，生成标准化的 Word / Excel 文档。

</div>

## 下载与安装

```bash
git clone https://github.com/virmuran/TemplateFill.git
cd TemplateFill

# 创建虚拟环境（推荐）
python -m venv .venv
source .venv/Scripts/activate   # Windows
# source .venv/bin/activate     # macOS / Linux

pip install -r requirements.txt
python main.py
```

| 依赖                                                 | 版本    | 用途                  |
| ---------------------------------------------------- | ------- | --------------------- |
| [PySide6](https://pypi.org/project/PySide6/)         | >= 6.5  | Qt6 GUI 桌面界面      |
| [docxtpl](https://pypi.org/project/docxtpl/)         | >= 0.16 | Word 模板 Jinja2 渲染 |
| [python-docx](https://pypi.org/project/python-docx/) | >= 0.8  | Word 文档生成         |
| [openpyxl](https://pypi.org/project/openpyxl/)       | >= 3.1  | Excel 模板读写        |

仓库只含源码、不含示例模板，点击软件里的 **「生成 Word 示例 / Excel 示例」** 会在 `~/TemplateFill/templates` 下按需生成。

## 亮点功能

- **双格式支持**：Word（`.docx`）和 Excel（`.xlsx`）模板均可，自动识别文件类型
- **智能解析**：加载模板即自动提取所有 `{{变量名}}` 占位符，无需手动配置
- **动态表单**：按字段语义自动选用单行 / 多行输入控件，也支持双击手动切换控件类型
- **批量生成**：导入 CSV / Excel 或粘贴数据，一键为每行数据生成独立文档
- **中文友好**：内置 130+ 常用字段的中文显示名映射，`author` 自动显示为「编制人」
- **标签定制**：双击字段标题即可改显示名和控件类型，配置自动存为 `.labels.json`，下次加载自动复用
- **循环区域**（Excel）：支持 `{% for item in items %}...{% endfor %}` 语法，明细表自动展开多行并修正公式引用
- **公式翻译**（Excel）：循环区域展开时，下游公式的行引用自动调整，`=SUM(F8:F8)` 自动变为 `=SUM(F7:F9)`
- **零配置上手**：点击「生成示例模板」即可创建一个完整的演示模板，快速体验全部功能

## 使用说明

### 基本流程

1. 点击 **「打开模板」** 加载 `.docx` 或 `.xlsx` 模板，或点击 **「生成 Word 示例 / Excel 示例」** 创建演示模板
2. 软件自动识别模板中所有 `{{变量}}`，并为每个变量生成对应的输入框
3. 填写各字段内容后，点击 **「生成文档」**，文件保存到桌面（自动带时间戳）

### 模板语法

| 语法     | 示例                                     | 适用格式     | 说明                 |
| -------- | ---------------------------------------- | ------------ | -------------------- |
| 简单变量 | `{{ title }}`                            | Word / Excel | 普通文本替换         |
| 过滤器   | `{{ name \| upper }}`                    | Word         | Jinja2 过滤器        |
| 循环     | `{% for item in items %}...{% endfor %}` | Word / Excel | 循环区域展开         |
| 条件     | `{% if condition %}...{% endif %}`       | Word         | 条件渲染             |
| 富文本   | `{{r content }}`                         | Word         | 保留换行、加粗等格式 |

### 标签显示名

加载模板后，每个字段的标题由以下优先级决定：

| 优先级 | 来源           | 说明                                 |
| ------ | -------------- | ------------------------------------ |
| 1      | 双击改名       | 运行时修改，内存中，当前会话有效     |
| 2      | `.labels.json` | 模板同目录的配套文件，关闭后仍生效   |
| 3      | 内置映射       | `labels.py` 中 130+ 常用字段的中文名 |
| 4      | 自动生成       | 将 `_` 替换为空格并首字母大写        |

双击任意字段标题即可弹出编辑对话框，可修改显示名和控件类型（单行 / 多行）。修改后自动写回 `.labels.json`。

### 批量模式

点击 **「◇ 单人模式」** 按钮切换至批量模式：

- 支持 **CSV / Excel 导入** 或 **直接粘贴** 数据
- 自动匹配表头与模板变量名
- 每行数据生成一个独立文件，文件名带序号和时间戳

### Excel 循环区域

在 Excel 模板中，使用 `{% for item in items %}...{% endfor %}` 定义循环区域：

- `{{item.field}}` 占位符自动替换为数据
- 数据行数量决定展开行数，同时修正上下游公式
- 保留原始单元格样式、格式、边框

## 更新日志

完整更新日志见 [`CHANGELOG.md`](CHANGELOG.md)。

## 加入我们

欢迎提交 [Issue](https://github.com/virmuran/TemplateFill/issues) 反馈 Bug 或建议，也欢迎 Pull Request。

## 致谢

### 开源库

- 模板引擎：[docxtpl](https://github.com/elapouya/python-docx-template)
- Word 文档：[python-docx](https://github.com/python-openxml/python-docx)
- Excel 读写：[openpyxl](https://foss.heptapod.net/openpyxl/openpyxl)
- GUI 框架：[PySide6](https://wiki.qt.io/Qt_for_Python)

## 声明

本软件使用 [MIT License](LICENSE) 开源，仅供学习交流使用。

## 许可证

[MIT License](LICENSE) © 2025-2026 ChemCal Team

联系方式：virmuran@163.com
