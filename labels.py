"""
TemplateFill — 共享标签映射工具

集中的标签中文名映射和 .labels.json 读写逻辑，
消除 doc_filler.py 和 xlsx_filler.py 中的重复代码。

.labels.json 文件结构（支持新旧两种格式）：
  新格式（推荐）：
    {
      "author": {"label": "编制人", "type": "single"},
      "budget": {"label": "预算",    "type": "multi"}
    }
  旧格式（仍可读取）：
    {
      "author": "编制人",
      "budget": "预算"
    }
"""

import os
import json

# ── 默认中文显示名映射（硬编码兜底） ────────────────

DEFAULT_LABELS = {
    # ── 文档元信息 ──
    'title': '文档标题',
    'subtitle': '副标题',
    'doc_number': '文档编号',
    'doc_type': '文档类型',
    'doc_title': '文档标题',
    'version': '版本号',
    'revision': '修订版本',
    'status': '状态',
    'page': '页码',
    'total_pages': '总页数',
    'keywords': '关键词',
    'abstract': '摘要',
    'description': '描述',
    'content': '内容',
    'summary': '总结',
    'remarks': '备注',
    'remark': '备注',
    'note': '附注',
    'attachment': '附件',
    'appendix': '附录',
    'reference': '参考文献',

    # ── 日期 ──
    'date': '日期',
    'create_date': '创建日期',
    'submit_date': '提交日期',
    'sign_date': '签署日期',
    'start_date': '开始日期',
    'end_date': '结束日期',
    'deadline': '截止日期',
    'valid_date': '有效日期',
    'expiry_date': '到期日期',
    'approval_date': '审批日期',

    # ── 人员 ──
    'author': '编制人',
    'reviewer': '审核人',
    'approver': '批准人',
    'drafter': '起草人',
    'manager': '负责人',
    'leader': '领导',
    'handler': '经办人',
    'creator': '创建人',
    'modifier': '修改人',
    'contact': '联系人',
    'contact_person': '联系人',

    # ── 单位信息 ──
    'company': '单位名称',
    'company_name': '单位名称',
    'address': '地址',
    'postal_code': '邮编',
    'phone': '联系电话',
    'tel': '电话',
    'fax': '传真',
    'email': '电子邮箱',
    'website': '网站',
    'tax_id': '税号',
    'bank_name': '开户银行',
    'bank_account': '银行账号',
    'legal_rep': '法定代表人',

    # ── 部门 ──
    'department': '部门',
    'dept': '部门',
    'unit': '单位',
    'team': '团队',
    'group': '组别',
    'division': '事业部',

    # ── 项目 ──
    'project_name': '项目名称',
    'project_no': '项目编号',
    'project_manager': '项目经理',
    'milestone': '里程碑',
    'phase': '阶段',
    'progress': '进度',
    'objective': '目标',
    'scope': '范围',
    'background': '背景',
    'method': '方法',
    'standard': '标准',
    'requirement': '需求',
    'deliverable': '交付物',
    'risk': '风险',
    'issue': '问题',

    # ── 财务 ──
    'amount': '金额',
    'total': '合计',
    'subtotal': '小计',
    'grand_total': '总计',
    'tax': '税额',
    'tax_rate': '税率',
    'discount': '折扣',
    'unit_price': '单价',
    'price': '单价',
    'quantity': '数量',
    'qty': '数量',
    'total_price': '总价',
    'amount_in_words': '大写金额',
    'budget': '预算',
    'cost': '成本',
    'fee': '费用',

    # ── 客户 / 客户方 ──
    'customer': '客户名称',
    'customer_name': '客户名称',
    'customer_address': '客户地址',
    'customer_contact': '客户联系人',
    'customer_phone': '客户电话',
    'supplier': '供应商',
    'supplier_name': '供应商名称',

    # ── 合同 ──
    'contract_no': '合同编号',
    'contract_name': '合同名称',
    'party_a': '甲方',
    'party_b': '乙方',
    'party_c': '丙方',
    'clause': '条款',
    'terms': '条款',

    # ── 文档正文 ──
    'overview': '项目概况',
    'technical_plan': '技术方案',
    'implementation': '实施计划',
    'conclusion': '结论与建议',
    'introduction': '引言',
    'preface': '前言',
    'purpose': '目的',
    'principle': '原则',
    'analysis': '分析',
    'evaluation': '评估',
    'suggestion': '建议',
    'proposal': '提案',

    # ── 报告专用 ──
    'report_date': '报告日期',
    'report_period': '报告期间',
    'data_source': '数据来源',
    'indicator': '指标',
    'target': '目标值',
    'actual': '实际值',
    'variance': '偏差',
    'trend': '趋势',
    'ranking': '排名',

    # ── 产品 ──
    'product': '产品名称',
    'product_name': '产品名称',
    'spec': '规格型号',
    'model': '型号',
    'brand': '品牌',
    'unit': '单位',
    'batch_no': '批号',
    'manufacturer': '制造商',

    # ── 序列 / 序号 ──
    'seq': '序号',
    'no': '序号',
    'index': '序号',
    'id': '编号',
    'code': '代码',
    'name': '名称',
    'type': '类型',
    'category': '分类',
    'level': '等级',
    'grade': '级别',
    'priority': '优先级',
}


def tag_to_label(tag):
    """将标签名转为友好的中文显示名（硬编码映射兜底）"""
    return DEFAULT_LABELS.get(tag, tag.replace('_', ' ').title())


# ── 长文本变量名（多行控件兜底判断） ────────────────

LONG_TEXT_TAGS = {
    'overview', 'technical_plan', 'implementation', 'budget', 'conclusion',
    'content', 'description', 'summary', 'abstract', 'details', 'remarks',
    'remark', 'note', 'note_text',
}


def is_long_text_tag(tag):
    """判断变量名默认应使用多行控件（兜底逻辑）"""
    if tag in LONG_TEXT_TAGS:
        return True
    return tag.endswith(('_text', '_plan', '_desc', '_content', '_details'))


# ── 标签数据结构（统一处理新旧格式） ────────────────

def parse_entry(value, tag=None):
    """将 .labels.json 中的值解析为统一格式。

    输入 → 输出：
      "编制人"           → {"label": "编制人", "type": "single"|"multi"}
      {"label": "..."}   → {"label": "...", "type": "..."}
      "" (空字符串)      → {"label": "", "type": "single"|"multi"}
    """
    if isinstance(value, dict):
        return {
            'label': value.get('label', '') or '',
            'type': value.get('type', 'multi' if is_long_text_tag(tag or '') else 'single'),
        }
    # 字符串 / 旧格式
    label = value if isinstance(value, str) else ''
    return {
        'label': label,
        'type': 'multi' if is_long_text_tag(tag or '') else 'single',
    }


def get_label(entries, tag):
    """从 entries 字典取 tag 的显示名（空字符串返回 None）"""
    entry = entries.get(tag)
    if entry is None:
        return None
    parsed = parse_entry(entry, tag)
    label = parsed['label']
    return label if label else None


def get_type(entries, tag):
    """从 entries 字典取 tag 的控件类型（'single' | 'multi'）"""
    entry = entries.get(tag)
    if entry is None:
        return 'multi' if is_long_text_tag(tag) else 'single'
    return parse_entry(entry, tag)['type']


# ── .labels.json 读写 ──────────────────────────────

def load_tag_labels(template_path):
    """从模板同目录的 .labels.json 加载标签显示名映射。

    文件命名规则：report.docx → report.labels.json

    返回 dict，若文件不存在或解析失败则返回空 dict。
    """
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


def create_label_skeleton(template_path, tags):
    """为模板生成 .labels.json 骨架（标签名→空字符串）。

    若文件已存在则跳过。
    """
    json_path = os.path.splitext(template_path)[0] + '.labels.json'
    if os.path.exists(json_path):
        return json_path

    skeleton = {t: "" for t in tags}
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(skeleton, f, ensure_ascii=False, indent=2)
    return json_path
