"""
TemplateFill — Excel (.xlsx) 模板填充引擎

核心能力：
  - parse_tags(): 解析 xlsx 模板中的简单变量 {{var}} 和循环区域 {% for %}...{% endfor %}
  - fill_template(): 填充模板，保留格式、公式、合并单元格等
  - 循环区域自动扩展行、公式自动翻译
"""

import re
import os
import json
from collections import OrderedDict, defaultdict
from copy import copy

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.formula.translate import Translator
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side


# ── 正则模式 ────────────────────────────────────────────

# 简单变量：{{var}} 或 {{var|filter}}
RE_SIMPLE = re.compile(r'\{\{\s*(\w[\w.]*)\s*(?:\|[^}]*)?\}\}')

# 循环开始：{% for item in items %}
RE_FOR = re.compile(r'\{%\s*for\s+(\w+)\s+in\s+(\w+)\s*%\}')

# 循环结束：{% endfor %}
RE_ENDFOR = re.compile(r'\{%\s*endfor\s*%\}')

# 条件开始：{% if ... %}
RE_IF = re.compile(r'\{%\s*if\s+(\w[\w.]*)')

# 条件结束：{% endif %}
RE_ENDIF = re.compile(r'\{%\s*endif\s*%\}')

# 循环内字段：{{item.field}}
RE_ITEM_FIELD = re.compile(r'\{\{\s*(\w+)\.(\w+)\s*(?:\|[^}]*)?\}\}')


# ── 标签解析 ────────────────────────────────────────────

def parse_tags(xlsx_path):
    """
    解析 Excel 模板，提取所有占位符。

    返回结构：
    {
        'simple': ['tag1', 'tag2', ...],      # 简单变量列表（去重排序）
        'loops': OrderedDict({                  # 循环区域
            'items': {                          # 循环变量名（{% for item in items %} 中的 items）
                'item_var': 'item',             # 循环项变量名
                'fields': ['name', 'amount'],   # {{item.name}} 中的字段名
                'sheet': 'Sheet1',
                'start_row': 5,                 # 1-indexed，含 for 标记行
                'end_row': 7,                   # 1-indexed，含 endfor 标记行
            }
        })
    }
    """
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"模板文件不存在: {xlsx_path}")

    wb = openpyxl.load_workbook(xlsx_path, data_only=False)

    simple_tags = set()
    loops = OrderedDict()

    for ws in wb.worksheets:
        sheet_loops = _find_loops_in_sheet(ws)
        for loop_var, loop_info in sheet_loops.items():
            if loop_var not in loops:
                loops[loop_var] = loop_info
            else:
                # 同名循环变量跨 sheet（如多个 sheet 都有 {% for item in items %}）
                # 用 sheet 名做后缀区分
                loops[f"{loop_var}_{ws.title}"] = loop_info

        sheet_simple = _find_simple_tags_in_sheet(ws, loops)
        simple_tags.update(sheet_simple)

    wb.close()

    # 过滤：从 simple 中移除属于循环的变量
    loop_vars = set(loops.keys())
    simple_tags = {t for t in simple_tags if t not in loop_vars}

    return {
        'simple': sorted(simple_tags),
        'loops': loops,
    }


def _find_simple_tags_in_sheet(ws, known_loops):
    """在 sheet 中查找简单变量 {{var}}，排除循环区域内的 {{item.field}}"""
    tags = set()
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            text = str(cell.value)

            # 跳过纯循环控制语句
            if RE_FOR.search(text) or RE_ENDFOR.search(text):
                continue

            # 提取简单变量（非 item.field 形式）
            matches = RE_SIMPLE.findall(text)
            for m in matches:
                if '.' not in m:
                    tags.add(m)

    return tags


def _find_loops_in_sheet(ws):
    """在 sheet 中查找 {% for item in items %}...{% endfor %} 循环区域"""
    loops = OrderedDict()
    open_loops = []  # [(row_idx, item_var, loop_var), ...]

    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        row_texts = []
        for cell in row:
            if cell.value is not None:
                row_texts.append(str(cell.value))

        combined = ' '.join(row_texts)

        # 检查 {% for %}
        for m in RE_FOR.finditer(combined):
            item_var = m.group(1)
            loop_var = m.group(2)
            open_loops.append((row_idx, item_var, loop_var))

        # 检查 {% endfor %}
        endfor_matches = list(RE_ENDFOR.finditer(combined))
        for _ in endfor_matches:
            if open_loops:
                start_row, item_var, loop_var = open_loops.pop()
                end_row = row_idx

                # 收集循环区域内的 {{item.field}} 字段
                fields = set()
                template_rows = range(start_row, end_row + 1)
                for tr in template_rows:
                    for cell in ws[tr]:
                        if cell.value is None:
                            continue
                        text = str(cell.value)
                        for fm in RE_ITEM_FIELD.finditer(text):
                            if fm.group(1) == item_var:
                                fields.add(fm.group(2))

                # 按首现顺序排序
                field_order = _field_appearance_order(ws, start_row, end_row, item_var)

                loops[loop_var] = {
                    'item_var': item_var,
                    'fields': field_order if field_order else sorted(fields),
                    'sheet': ws.title,
                    'start_row': start_row,
                    'end_row': end_row,
                }

    return loops


def _field_appearance_order(ws, start_row, end_row, item_var):
    """获取循环区域内字段的首次出现顺序（按列）"""
    seen = set()
    order = []
    for r in range(start_row, end_row + 1):
        for cell in ws[r]:
            if cell.value is None:
                continue
            text = str(cell.value)
            for m in RE_ITEM_FIELD.finditer(text):
                if m.group(1) == item_var:
                    field = m.group(2)
                    if field not in seen:
                        seen.add(field)
                        order.append(field)
    return order


# ── 模板填充 ────────────────────────────────────────────

def fill_template(xlsx_path, context, output_path=None):
    """
    填充 Excel 模板并保存。

    参数：
      xlsx_path:   模板 .xlsx 文件路径
      context:     dict，简单变量映射为字符串，循环变量映射为 list[dict]
                   例：{'title': '报价单', 'items': [{'name': 'A', 'qty': 10}, ...]}
      output_path: 输出路径，默认桌面 + 模板名_时间戳.xlsx

    返回：输出文件的完整路径
    """
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"模板文件不存在: {xlsx_path}")

    wb = openpyxl.load_workbook(xlsx_path)

    # 解析模板结构
    template_info = parse_tags(xlsx_path)

    # 按 sheet 分组处理循环
    loops_by_sheet = defaultdict(list)
    for loop_var, loop_info in template_info['loops'].items():
        loops_by_sheet[loop_info['sheet']].append((loop_var, loop_info))

    # 逐 sheet 处理
    for ws in wb.worksheets:
        sheet_loops = loops_by_sheet.get(ws.title, [])

        if sheet_loops:
            # 按 start_row 降序处理（从底部开始，避免行号偏移）
            sheet_loops.sort(key=lambda x: -x[1]['start_row'])
            for loop_var, loop_info in sheet_loops:
                items = context.get(loop_var, [])
                if not isinstance(items, list):
                    items = []
                _expand_loop(ws, loop_info, items)

        # 处理简单变量
        _fill_simple_vars(ws, context)

    wb.save(_resolve_output(output_path, xlsx_path))
    wb.close()
    return _resolve_output(output_path, xlsx_path)


def _resolve_output(output_path, template_path):
    """解析输出路径"""
    if output_path is not None:
        return output_path
    from datetime import datetime
    base = os.path.splitext(os.path.basename(template_path))[0]
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    return os.path.join(
        os.path.expanduser("~"), "Desktop",
        f"{base}_{ts}.xlsx"
    )


def _expand_loop(ws, loop_info, items):
    """
    展开循环区域：
    1. 找到循环内的数据行（含 {{item.field}} 的行）
    2. 保存数据行的样式和模板内容
    3. 删除整个模板区域
    4. 插入数据行
    5. 逐行填充数据 + 翻译公式
    6. 修复后续公式中引用模板行号的部分

    参数：
      ws:         工作表对象
      loop_info:  {'start_row', 'end_row', 'item_var', 'fields', ...}
      items:      list[dict]，每个 dict 对应一行数据
    """
    start_row = loop_info['start_row']
    end_row = loop_info['end_row']
    item_var = loop_info['item_var']
    template_row_count = end_row - start_row + 1
    item_count = len(items)
    max_col = ws.max_column

    # 阶段一：识别数据行（含 {{item.field}} 但不含 for/endfor 控制语句）
    data_rows = []  # [(row_number, row_height, [cell_info, ...])]
    for src_row in range(start_row, end_row + 1):
        is_data_row = False
        row_data = []
        row_height = ws.row_dimensions[src_row].height if src_row in ws.row_dimensions else None
        for col in range(1, max_col + 1):
            src_cell = ws.cell(row=src_row, column=col)
            val = src_cell.value
            text = str(val) if val is not None else ''
            cell_info = {
                'col': col,
                'value': val,
                'font': copy(src_cell.font),
                'fill': copy(src_cell.fill),
                'border': copy(src_cell.border),
                'alignment': copy(src_cell.alignment),
                'number_format': src_cell.number_format,
            }
            row_data.append(cell_info)
            if RE_ITEM_FIELD.search(text):
                is_data_row = True
        if is_data_row:
            data_rows.append((src_row, row_height, row_data))

    data_rows_per_item = len(data_rows)

    # 阶段一B：保存并清除受影响的合并单元格
    # openpyxl 的 delete_rows/insert_rows 不会自动更新合并单元格范围，
    # 导致保存时 B/C 列数据被当作合并区域清空
    saved_merges = _save_and_clear_merges(ws, start_row)

    # 阶段二：删除模板行
    ws.delete_rows(start_row, template_row_count)

    if data_rows_per_item == 0 or item_count == 0:
        # 无数据行：只恢复合并单元格
        total_data_rows = 0
        _restore_merges(ws, saved_merges, start_row,
                        template_row_count, total_data_rows)
        return

    # 阶段三：插入数据行
    total_data_rows = item_count * data_rows_per_item
    ws.insert_rows(start_row, total_data_rows)

    # 阶段四：填充数据
    for item_idx, item_data in enumerate(items):
        for local_idx, (template_row_num, row_height, row_info) in enumerate(data_rows):
            dest_row = start_row + item_idx * data_rows_per_item + local_idx

            # 行高
            if row_height:
                if dest_row not in ws.row_dimensions:
                    ws.row_dimensions[dest_row] = openpyxl.worksheet.dimensions.RowDimension(ws, index=dest_row)
                ws.row_dimensions[dest_row].height = row_height

            for cell_info in row_info:
                col = cell_info['col']
                dest_cell = ws.cell(row=dest_row, column=col)

                # 复制样式
                dest_cell.font = cell_info['font']
                dest_cell.fill = cell_info['fill']
                dest_cell.border = cell_info['border']
                dest_cell.alignment = cell_info['alignment']
                dest_cell.number_format = cell_info['number_format']

                raw_value = cell_info['value']
                if raw_value is None:
                    continue

                text = str(raw_value)

                # 替换 {{item.field}} — 用已知字段名直接替换
                for field_name in loop_info.get('fields', []):
                    placeholder = f"{{{{{item_var}.{field_name}}}}}"
                    if placeholder in text:
                        val = item_data.get(field_name, '')
                        text = text.replace(placeholder, str(val) if val is not None else '')

                if raw_value is not None and str(raw_value).startswith('='):
                    # 公式：翻译行引用（即使没有占位符也要翻译）
                    formula = text
                    try:
                        origin = f"{get_column_letter(col)}{template_row_num}"
                        trans = Translator(formula, origin)
                        formula = trans.translate_formula(row_delta=dest_row - template_row_num)
                    except Exception:
                        pass
                    dest_cell.value = formula
                elif '{{' not in text:
                    dest_cell.value = text
                else:
                    dest_cell.value = raw_value

    # 阶段五：修复下游公式中对模板数据行的引用
    _fix_post_loop_formulas(ws, loop_info, items, item_count, data_rows_per_item, total_data_rows)

    # 阶段六：恢复合并单元格（按行偏移调整位置）
    _restore_merges(ws, saved_merges, start_row,
                    template_row_count, total_data_rows)


def _save_and_clear_merges(ws, start_row):
    """保存并清除 start_row 及以下所有合并单元格范围。返回 [(min_col, min_row, max_col, max_row), ...]"""
    merges = []
    to_remove = []
    for mc in list(ws.merged_cells.ranges):
        if mc.min_row >= start_row or mc.max_row >= start_row:
            merges.append((mc.min_col, mc.min_row, mc.max_col, mc.max_row))
            to_remove.append(str(mc))
    for mc_str in to_remove:
        ws.unmerge_cells(mc_str)
    return merges


def _restore_merges(ws, saved_merges, start_row, deleted_count, inserted_count):
    """按行偏移恢复合并单元格。shift = inserted_count - deleted_count"""
    row_shift = inserted_count - deleted_count
    for min_col, min_row, max_col, max_row in saved_merges:
        new_min_row = min_row
        new_max_row = max_row
        # 在删除区内的合并单元格：删除（不恢复）
        if min_row >= start_row and max_row < start_row + deleted_count:
            continue
        # 在删除区下方的：偏移
        if min_row >= start_row:
            new_min_row = min_row + row_shift
            new_max_row = max_row + row_shift
        # 跨越删除区的：保留上方部分
        elif max_row >= start_row:
            new_max_row = max_row + row_shift
        try:
            if new_min_row != new_max_row or min_col != max_col:
                from openpyxl.utils import get_column_letter as _gcl
                range_str = f"{_gcl(min_col)}{new_min_row}:{_gcl(max_col)}{new_max_row}"
                ws.merge_cells(range_str)
        except Exception:
            pass


def _fix_post_loop_formulas(ws, loop_info, items, item_count, data_rows_per_item, total_data_rows):
    """
    修复循环展开后，下游公式中对原模板数据行的引用。

    例如模板中 row 10 有 =SUM(F8:F8)，展开后数据在 rows 7-9，
    需改为 =SUM(F7:F9)。
    """
    start_row = loop_info['start_row']
    end_row = loop_info['end_row']
    template_row_count = end_row - start_row + 1

    # 模板中第一个数据行和最后一个数据行的原始行号
    # (需要从 parse 信息中推断，简化处理：用 start_row+1 到 end_row-1 作为数据行范围)
    first_template_data_row = start_row + 1  # 假设第一行是 for，第二行开始是数据
    last_template_data_row = end_row - 1      # 假设最后一行是 endfor

    # 展开后的数据行范围
    final_data_start = start_row
    final_data_end = start_row + total_data_rows - 1

    # 行号偏移：原行 - 模板行数 → 数据区插入后
    # 原来在 end_row+1 的行，现在在 end_row+1 - template_row_count + total_data_rows
    row_shift = total_data_rows - template_row_count

    # 扫描所有公式，修复对模板数据行的引用
    import re as re_mod
    # 匹配形如 XN 的单元格引用（X=字母, N=数字）
    cell_ref_pattern = re_mod.compile(r'([A-Z]+)(\d+)')

    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            val = str(cell.value)
            if not val.startswith('='):
                continue

            # 只修复引用行号在模板数据行范围内的公式
            # 同时也处理行号大于 end_row 的（因为 delete + insert 导致偏移）
            def _fix_ref(match):
                col_letter = match.group(1)
                row_num = int(match.group(2))
                if first_template_data_row <= row_num <= last_template_data_row:
                    # 这个引用指向模板数据行，调整到展开后的范围
                    # 如果只有1个数据行模板，展开后数据均匀分布
                    # 简化：将整个数据行范围扩大
                    new_row = row_num + row_shift
                    return f"{col_letter}{new_row}"
                return match.group(0)

            # 对于 SUM/AVERAGE 等范围函数，处理 F8:F8 这种引用
            range_pattern = re_mod.compile(r'([A-Z]+)(\d+):([A-Z]+)(\d+)')
            new_val = range_pattern.sub(
                lambda m: _fix_range(m, first_template_data_row, last_template_data_row,
                                     final_data_start, final_data_end, row_shift),
                val
            )
            if new_val != val:
                cell.value = new_val


def _fix_range(match, first_tpl, last_tpl, final_start, final_end, row_shift):
    """修复公式中的范围引用"""
    col1, row1 = match.group(1), int(match.group(2))
    col2, row2 = match.group(3), int(match.group(4))

    if first_tpl <= row1 <= last_tpl or first_tpl <= row2 <= last_tpl:
        # 这个范围引用了模板数据行
        return f"{col1}{final_start}:{col2}{final_end}"
    return match.group(0)


def _copy_cell_style(src, dst):
    """复制单元格样式"""
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def _fill_simple_vars(ws, context):
    """替换 sheet 中所有的简单变量 {{var}}"""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            text = str(cell.value)

            # 跳过循环控制语句（可能在上层已被删除，但做个防御）
            if RE_FOR.search(text) or RE_ENDFOR.search(text):
                # 如果还残留，替换为空
                if RE_FOR.search(text) or RE_ENDFOR.search(text):
                    cell.value = ''
                continue

            # 替换 {{var}}
            def _replace_var(match):
                var_name = match.group(1)
                if '.' in var_name:
                    # {{item.field}} 理论上不应该在 simple 阶段出现，跳过
                    return match.group(0)
                val = context.get(var_name, '')
                return str(val) if val is not None else ''

            new_text = RE_SIMPLE.sub(_replace_var, text)

            if '{{' not in new_text:
                # 尝试保留原始数据类型
                # 如果替换后是纯数字，保持为数值
                if new_text != text:
                    try:
                        # 检查原始值是否为数字类型
                        if isinstance(cell.value, (int, float)):
                            pass  # 数字型单元格替换后保持文本，因为模板里嵌入数字不太常见
                    except Exception:
                        pass
                cell.value = new_text


# ── 标签映射工具（与 docx 版本共享接口） ────────────────

def load_tag_labels(template_path):
    """从模板同目录的 .labels.json 加载标签显示名映射"""
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
    """为模板自动生成 .labels.json 骨架（仅包含简单变量）"""
    json_path = os.path.splitext(template_path)[0] + '.labels.json'
    if os.path.exists(json_path):
        return json_path

    result = parse_tags(template_path)
    simple_tags = result['simple']
    skeleton = {t: "" for t in simple_tags}

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(skeleton, f, ensure_ascii=False, indent=2)

    return json_path


# ── 示例模板生成 ────────────────────────────────────────

def create_sample_template(output_path):
    """
    创建一个示例 Excel 模板，包含：
      - 简单变量：抬头信息（标题、客户、日期）
      - 循环区域：产品明细表（支持多行填写）
      - 公式：金额计算和合计
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "报价单"

    # ── 样式定义 ──
    header_font = Font(name='微软雅黑', size=16, bold=True, color='1F4E79')
    label_font = Font(name='微软雅黑', size=11, bold=True, color='333333')
    data_font = Font(name='微软雅黑', size=11)
    table_header_font = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
    table_header_fill = PatternFill(start_color='4A6FA5', end_color='4A6FA5', fill_type='solid')
    table_header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    table_cell_align = Alignment(vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin', color='D0D5DD'),
        right=Side(style='thin', color='D0D5DD'),
        top=Side(style='thin', color='D0D5DD'),
        bottom=Side(style='thin', color='D0D5DD'),
    )
    amount_format = '#,##0.00'

    # ── 标题行 ──
    ws.merge_cells('A1:F1')
    title_cell = ws['A1']
    title_cell.value = '{{title}}'
    title_cell.font = header_font
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 40

    # ── 基本信息行 ──
    ws.merge_cells('A3:B3')
    ws['A3'].value = '客户名称：{{customer}}'
    ws['A3'].font = data_font
    ws.merge_cells('D3:E3')
    ws['D3'].value = '日期：{{date}}'
    ws['D3'].font = data_font
    ws.merge_cells('A4:B4')
    ws['A4'].value = '联系人：{{contact}}'
    ws['A4'].font = data_font
    ws.merge_cells('D4:E4')
    ws['D4'].value = '联系电话：{{phone}}'
    ws['D4'].font = data_font

    # ── 产品明细表头 ──
    header_row = 6
    headers = ['序号', '产品名称', '规格型号', '数量', '单价', '金额']
    col_widths = [8, 22, 16, 10, 12, 14]
    for col_idx, (h, w) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=header_row, column=col_idx)
        cell.value = h
        cell.font = table_header_font
        cell.fill = table_header_fill
        cell.alignment = table_header_align
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    ws.row_dimensions[header_row].height = 28

    # ── 循环区域：{% for item in items %}...{% endfor %} ──
    for_start_row = 7
    ws.cell(row=for_start_row, column=1).value = '{% for item in items %}'

    data_row = 8
    ws.cell(row=data_row, column=1).value = '{{item.seq}}'
    ws.cell(row=data_row, column=2).value = '{{item.product}}'
    ws.cell(row=data_row, column=3).value = '{{item.spec}}'
    ws.cell(row=data_row, column=4).value = '{{item.quantity}}'
    ws.cell(row=data_row, column=5).value = '{{item.price}}'
    # 金额公式（占位符 + 公式结构）
    ws.cell(row=data_row, column=6).value = '=D8*E8'

    for_end_row = 9
    ws.cell(row=for_end_row, column=1).value = '{% endfor %}'

    # 为循环区域模板行设置样式
    for r in [for_start_row, data_row, for_end_row]:
        ws.row_dimensions[r].height = 22
        for c in range(1, 7):
            cell = ws.cell(row=r, column=c)
            cell.font = data_font
            cell.alignment = table_cell_align
            cell.border = thin_border
            if c == 5 or c == 6:  # 单价/金额列
                cell.number_format = amount_format

    # ── 合计行 ──
    total_row = 10
    ws.merge_cells(f'A{total_row}:C{total_row}')
    ws.cell(row=total_row, column=1).value = '合计'
    ws.cell(row=total_row, column=1).font = Font(name='微软雅黑', size=11, bold=True)
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal='right', vertical='center')
    ws.cell(row=total_row, column=6).value = '=SUM(F8:F8)'
    ws.cell(row=total_row, column=6).font = Font(name='微软雅黑', size=11, bold=True)
    ws.cell(row=total_row, column=6).number_format = amount_format
    ws.cell(row=total_row, column=6).border = thin_border

    # ── 备注行 ──
    note_row = 12
    ws.merge_cells(f'A{note_row}:C{note_row}')
    ws.cell(row=note_row, column=1).value = '备注：'
    ws.cell(row=note_row, column=1).font = label_font
    ws.merge_cells(f'A{note_row+1}:F{note_row+1}')
    ws.cell(row=note_row+1, column=1).value = '{{remarks}}'
    ws.cell(row=note_row+1, column=1).font = data_font
    ws.cell(row=note_row+1, column=1).alignment = Alignment(vertical='top', wrap_text=True)
    ws.row_dimensions[note_row+1].height = 50

    # 保存
    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    wb.save(output_path)
    wb.close()
    return output_path


# ── 调试入口 ────────────────────────────────────────────

if __name__ == '__main__':
    import tempfile

    # 创建示例模板
    tmp_dir = tempfile.gettempdir()
    tmp = os.path.join(tmp_dir, 'test_template.xlsx')
    create_sample_template(tmp)
    print(f'示例模板已创建: {tmp}')

    # 解析标签
    result = parse_tags(tmp)
    print(f"简单变量 ({len(result['simple'])}): {result['simple']}")
    print(f"循环区域:")
    for lv, li in result['loops'].items():
        print(f"  {lv}: fields={li['fields']}, range={li['start_row']}-{li['end_row']}")

    # 填充
    ctx = {
        'title': '产品报价单',
        'customer': '张三科技有限公司',
        'date': '2026-07-05',
        'contact': '李四',
        'phone': '13800138000',
        'items': [
            {'seq': 1, 'product': '服务器主机', 'spec': 'Dell R750xs', 'quantity': 2, 'price': 45000},
            {'seq': 2, 'product': '网络交换机', 'spec': 'H3C S5560X', 'quantity': 3, 'price': 8500},
            {'seq': 3, 'product': 'UPS电源', 'spec': 'APC SRT3000', 'quantity': 1, 'price': 12000},
        ],
        'remarks': '以上报价含税，有效期30天。',
    }
    out = fill_template(tmp, ctx)
    print(f'已生成: {out}')
