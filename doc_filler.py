"""
TemplateFill — 模板驱动的表单填写工具

核心逻辑：
  - parse_tags(): 从 DOCX 模板中提取所有 Jinja2 占位符 {{tag}}
  - fill_template(): 将用户输入填充到模板并导出
  - has_body_marker(): 检查模板是否含 {{__BODY__}} 动态内容标记
  - fill_template() 支持 sections 参数：在 {{__BODY__}} 位置插入动态章节
"""

import re
import os
from docxtpl import DocxTemplate
from html_to_docx import BODY_MARKER, insert_sections, has_body_marker as _check_body_marker


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


def fill_template(docx_path, context, output_path=None, sections=None):
    """
    填充模板并保存。

    参数：
      docx_path:   模板 .docx 文件路径
      context:     dict, 键为标签名, 值为填入内容
      output_path: 输出路径, 默认为桌面 + 模板名_时间戳.docx
      sections:    动态章节数据（list of dict），若模板含 {{__BODY__}} 标记
                   则在标记位置插入章节内容
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"模板文件不存在: {docx_path}")

    doc = DocxTemplate(docx_path)

    # 如果有动态章节，注入 body marker 用于后续定位
    if sections is not None:
        context = dict(context)  # 不修改原 dict
        context['__BODY__'] = BODY_MARKER

    doc.render(context)

    # 后处理：插入动态章节内容
    if sections is not None:
        insert_sections(doc, sections)

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

    # 正文 — 固定段落 + 动态内容区
    doc.add_paragraph()
    doc.add_heading('一、项目概况', level=2)
    doc.add_paragraph('{{overview}}')

    doc.add_heading('二、技术方案', level=2)
    doc.add_paragraph('{{technical_plan}}')

    # 动态内容插入点（软件中可增删章节）
    doc.add_heading('正文内容', level=2)
    doc.add_paragraph('{{__BODY__}}')

    # 页脚
    doc.add_paragraph()
    footer = doc.add_paragraph('编制单位：{{company}}　　联系电话：{{phone}}')
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.save(output_path)
    return output_path


def has_body_marker(docx_path):
    """检查模板是否包含 {{__BODY__}} 动态内容标记"""
    return _check_body_marker(docx_path)


def load_tag_labels(template_path):
    """Deprecated: 请使用 labels.load_tag_labels 代替"""
    from labels import load_tag_labels as _load
    return _load(template_path)


def create_sample_labels(template_path):
    """Deprecated: 请使用 labels.create_label_skeleton 代替"""
    from labels import create_label_skeleton
    tags = parse_tags(template_path)
    return create_label_skeleton(template_path, tags)


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
