# TemplateFill

> 模板驱动的 Office 文档生成工具 —— 加载模板，填写变量，一键生成 Word / Excel 文档。

## 功能特性

- **智能模板解析** —— 自动识别 DOCX 模板中所有 `{{变量名}}` 占位符，无需手动配置
- **动态表单生成** —— 根据模板变量自动生成输入界面，短文本用单行框，长文本用多行编辑区
- **中文化字段映射** —— 内建常用字段的中文名映射（如 `budget` → 预算与资源）
- **一键生成** —— 填写后点击按钮，自动输出到桌面，文件名自带时间戳
- **示例模板** —— 可一键生成项目报告示例模板，快速体验完整流程
- **支持 Jinja2 语法** —— 简单变量、过滤器、循环、条件、富文本均支持

## 安装

### 环境要求

- Python 3.8+
- Windows / macOS / Linux

### 安装依赖

```bash
pip install -r requirements.txt
```

## 使用

### 启动程序

```bash
python main.py
```

### 操作流程

1. 点击 **「打开模板」** 选择已有的 `.docx` 模板文件，或点击 **「生成示例模板」** 创建一个演示模板
2. 在自动生成的表单中填写各字段内容
3. 点击 **「生成文档」**，生成的文档将保存到桌面

### 模板语法

模板使用 Jinja2 风格的占位符，支持以下语法：

| 语法 | 示例 | 说明 |
|------|------|------|
| 简单变量 | `{{ title }}` | 普通文本替换 |
| 过滤器 | `{{ name \| upper }}` | 对变量应用 Jinja2 过滤器 |
| 循环 | `{% for item in items %}` | 遍历列表 |
| 条件 | `{% if condition %}` | 条件判断 |
| 富文本 | `{{r content }}` | 保留原文本格式（换行、加粗等） |

#### 示例模板变量

`create_sample_template()` 生成的示例模板包含以下变量：

| 变量名 | 中文名 | 说明 |
|--------|--------|------|
| `title` | 标题 | 文档标题 |
| `project_name` | 项目名称 | 项目名称 |
| `department` | 部门 | 所属部门 |
| `author` | 作者 | 编制人 |
| `date` | 日期 | 文档日期 |
| `reviewer` | 审核人 | 审核人姓名 |
| `doc_number` | 文档编号 | 文档编号 |
| `company` | 公司名称 | 编制单位 |
| `phone` | 联系电话 | 联系电话 |
| `overview` | 项目概况 | 长篇内容，多行编辑 |
| `technical_plan` | 技术方案 | 长篇内容，多行编辑 |
| `implementation` | 实施计划 | 长篇内容，多行编辑 |
| `budget` | 预算与资源 | 长篇内容，多行编辑 |
| `conclusion` | 结论与建议 | 长篇内容，多行编辑 |

## 项目结构

```
TemplateFill/
├── main.py                  # GUI 主程序 (PySide6)
├── doc_filler.py            # 核心逻辑：模板解析、填充、示例模板生成
├── templates/
│   └── sample_report.docx   # 示例模板（程序自动生成）
├── requirements.txt          # Python 依赖
└── README.md
```

## 依赖

| 库 | 用途 |
|----|------|
| [PySide6](https://pypi.org/project/PySide6/) | Qt6 GUI 框架，构建桌面界面 |
| [docxtpl](https://pypi.org/project/docxtpl/) | Jinja2 模板渲染引擎，将数据填充到 DOCX |
| [python-docx](https://pypi.org/project/python-docx/) | 创建和操作 DOCX 文档（用于生成示例模板） |

## License

MIT
