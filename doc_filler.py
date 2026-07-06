"""
TemplateFill — 模板驱动的表单填写工具

核心逻辑：
  - parse_tags(): 从 DOCX 模板中提取所有 Jinja2 占位符 {{tag}}
  - fill_template(): 将用户输入填充到模板并导出
"""

import re
import os
from docxtpl import DocxTemplate


def parse_tags(docx_path):
    """
    解析 DOCX 模板，提取所有 {{tag_name}} 占位符。
    返回：去重排序后的标签名列表。
    支持：
      - 简单变量 {{name}}
      - 带 filter {{name|upper}}
      - 循环 {% for x in items %}...{% endfor %}
      - 条件 {% if %}...{% endif %}
      - RichText {{r name}} — 富文本标签
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"模板文件不存在: {docx_path}")

    doc = DocxTemplate(docx_path)
    # 获取未声明的变量列表（docxtpl 内置方法）
    try:
        undeclared = doc.get_undeclared_template_variables()
        tags = sorted(set(undeclared))
    except Exception:
        # 回退：手动解析 XML
        tags = _parse_tags_from_xml(docx_path)

    # 过滤掉 Jinja2 内置变量和循环变量
    filtered = []
    for t in tags:
        # 跳过空值和明显的内置名字
        if not t or t.startswith('_') or t in ('l', 'r'):
            continue
        # 跳过 for 循环的内嵌变量（如 x in items 就是 items 需要，x 不需要）
        filtered.append(t)

    return sorted(set(filtered))


def _parse_tags_from_xml(docx_path):
    """回退方案：手动从 DOCX XML 中解析 {{tag}}"""
    import zipfile
    tags = set()
    with zipfile.ZipFile(docx_path, 'r') as z:
        for name in z.namelist():
            if name.startswith('word/') and name.endswith('.xml'):
                content = z.read(name).decode('utf-8', errors='ignore')
                # 匹配 {{ variable_name }}
                found = re.findall(r'\{\{\s*(\w[\w.]*)\s*(?:\|[^}]*)?\}\}', content)
                tags.update(found)
                # 匹配 {% for x in items %} → 提取 items
                for_blocks = re.findall(r'\{%\s*for\s+\w+\s+in\s+(\w+)\s*%\}', content)
                tags.update(for_blocks)
                # 匹配 {% if condition %} → 提取变量名
                if_blocks = re.findall(r'\{%\s*if\s+(\w+)', content)
                tags.update(if_blocks)
    return list(tags)


def fill_template(docx_path, context, output_path=None):
    """
    填充模板并保存。

    参数：
      docx_path:   模板 .docx 文件路径
      context:     dict, 键为标签名, 值为填入内容
      output_path: 输出路径, 默认为桌面 + 模板名_时间戳.docx
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"模板文件不存在: {docx_path}")

    doc = DocxTemplate(docx_path)
    doc.render(context)

    if output_path is None:
        from datetime import datetime
        base = os.path.splitext(os.path.basename(docx_path))[0]
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(
            os.path.expanduser("~"), "Desktop",
            f"{base}_已填写_{ts}.docx"
        )

    doc.save(output_path)
    return output_path


def create_sample_template(output_path):
    """
    创建一个示例 DOCX 模板，包含常用标签。
    用于首次使用时演示功能。
    """
    from docx import Document
    from docx.shared import Pt, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()

    # 设置默认字体
    style = doc.styles['Normal']
    font = style.font
    font.name = '宋体'
    font.size = Pt(11)

    # 标题
    title = doc.add_heading('{{title}}', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 基本信息表格
    doc.add_paragraph()
    table = doc.add_table(rows=4, cols=4, style='Table Grid')
    cells = [
        ('部门', '{{department}}', '日期', '{{date}}'),
        ('编制人', '{{author}}', '审核人', '{{reviewer}}'),
        ('项目名称', '{{project_name}}', '编号', '{{doc_number}}'),
        ('', '', '', ''),
    ]
    for i, row_data in enumerate(cells):
        for j, text in enumerate(row_data):
            if text:
                cell = table.cell(i, j)
                cell.text = text
                # 标签头加粗
                if not text.startswith('{{'):
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.bold = True
                            run.font.size = Pt(10)

    # 正文
    doc.add_paragraph()
    doc.add_heading('一、项目概况', level=2)
    doc.add_paragraph('{{overview}}')

    doc.add_heading('二、技术方案', level=2)
    doc.add_paragraph('{{technical_plan}}')

    doc.add_heading('三、实施计划', level=2)
    doc.add_paragraph('{{implementation}}')

    doc.add_heading('四、预算与资源', level=2)
    doc.add_paragraph('{{budget}}')

    doc.add_heading('五、结论与建议', level=2)
    doc.add_paragraph('{{conclusion}}')

    # 页脚
    doc.add_paragraph()
    footer = doc.add_paragraph('编制单位：{{company}}　　联系电话：{{phone}}')
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.save(output_path)
    return output_path


def load_tag_labels(template_path):
    """
    从模板同目录的 .labels.json 文件加载标签显示名映射。

    文件命名规则：report.docx → report.labels.json

    JSON 格式示例：
    {
        "customer_level": "客户等级",
        "department_header": "页眉_部门"
    }

    返回：dict，若文件不存在或解析失败则返回空 dict。
    """
    import json

    json_path = os.path.splitext(template_path)[0] + '.labels.json'

    if not os.path.exists(json_path):
        return {}

    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except (json.JSONDecodeError, IOError, UnicodeDecodeError):
        pass

    return {}


def create_sample_labels(template_path):
    """
    为指定模板自动生成 .labels.json 骨架文件。
    将模板中所有标签列出，显示名留空等待用户填写。
    如果 .labels.json 已存在则跳过。
    """
    import json

    json_path = os.path.splitext(template_path)[0] + '.labels.json'
    if os.path.exists(json_path):
        return json_path  # 已存在，不覆盖

    tags = parse_tags(template_path)
    skeleton = {t: "" for t in tags}

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(skeleton, f, ensure_ascii=False, indent=2)

    return json_path


if __name__ == '__main__':
    # 测试
    import tempfile
    tmp = os.path.join(tempfile.gettempdir(), 'test_template.docx')
    create_sample_template(tmp)
    print(f'示例模板已创建: {tmp}')

    tags = parse_tags(tmp)
    print(f'检测到标签 ({len(tags)}): {tags}')

    ctx = {t: f'[示例: {t}]' for t in tags}
    out = fill_template(tmp, ctx)
    print(f'已生成: {out}')
