"""
TemplateFill — 批量填充管理器

将批量模式的表格构建、CSV/剪贴板导入、批量生成等逻辑
从主窗口拆分为独立模块，保持 TemplateFillWindow 精简。
"""

import os
import csv

from PySide6.QtWidgets import (
    QLabel, QPushButton, QHBoxLayout,
    QFileDialog, QMessageBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QApplication, QSizePolicy,
)
from PySide6.QtCore import Qt


class BatchManager:
    """批量填充模式管理器。

    持有批量模式专用状态，通过 window 引用操作主窗口 UI。
    """

    def __init__(self, window):
        self.window = window          # TemplateFillWindow 引用
        self.table = None             # QTableWidget
        self.tags = []                # 当前模板的简单标签列表
        self.loops = {}               # 当前模板的循环区域
        self.defaults = {}            # 从单人模式带入的默认值
        self.mode = False             # 是否在批量模式

    # ── 模式切换 ────────────────────────────────────────────

    def toggle(self):
        """切换单人 / 批量模式"""
        if not self.window._template_path:
            QMessageBox.information(self.window, "提示", "请先加载一个模板文件。")
            return
        self.mode = not self.mode
        if self.mode:
            self._capture_defaults()
        self.sync_buttons()
        self.window._build_form(self.tags, self.loops)
        if not self.mode:
            # 切回单人模式：用模板记忆把上次填写内容带回来
            self._restore_single_values()
        self.window.status_bar.showMessage(
            "已切换到批量模式：每行数据生成一个独立文件" if self.mode else "已切回单人模式", 4000)

    def _restore_single_values(self):
        """从模板记忆 / 项目数据恢复单人模式的字段内容（批量模式下没有独立输入框）"""
        try:
            state = self.window._restored_state or \
                self.window._load_previous_state(self.window._template_path) or {}
            if not state:
                return
            self.window._apply_context(state.get('values') or {},
                                       state.get('loops') or [],
                                       state.get('sections') or [])
            self.window._update_summary()
        except Exception:
            pass

    def sync_buttons(self):
        """按钮状态与当前模式保持同步（进入/退出批量、打开项目后调用）"""
        style_on = """
            QPushButton { background: #eceff5; color: #4a6fa5; border: 1px solid #c8d4e4;
                         border-radius: 4px; padding: 6px 14px; font-size: 13px; }
            QPushButton:hover { background: #e2e8f2; }
        """
        style_off = """
            QPushButton { background: transparent; color: #4a4a46; border: 1px solid #d8d8d4;
                         border-radius: 4px; padding: 6px 14px; font-size: 13px; }
            QPushButton:hover { background: #f2f2ef; border-color: #c3c3bf; }
        """
        if self.mode:
            self.window.btn_batch.setText("批量模式 ✓")
            self.window.btn_batch.setStyleSheet(style_on)
            self.window.btn_generate.setText("批量生成")
            self.window.btn_generate.setToolTip("每行数据生成一个独立文件")
        else:
            self.window.btn_batch.setText("批量模式")
            self.window.btn_batch.setStyleSheet(style_off)
            self.window.btn_generate.setText("生成文档")
            self.window.btn_generate.setToolTip("")

    def _capture_defaults(self):
        """从单人模式收集当前填值作为批量默认值"""
        self.defaults = {}
        for tag, widget in self.window._field_widgets.items():
            if isinstance(widget, QLabel):
                continue
            val = widget.toPlainText() if hasattr(widget, 'toPlainText') else widget.text()
            if val.strip():
                self.defaults[tag] = val.strip()

    # ── 表单构建 ────────────────────────────────────────────

    def build_form(self, tags, loops):
        """批量模式：构建大数据表格，每行 = 一个输出文件"""
        fl = self.window.form_layout

        # 提示（页头已说明模式，这里只补默认值来源）
        if self.defaults:
            hint = QLabel(f"已从单人模式带入 {len(self.defaults)} 个默认值")
            hint.setStyleSheet("font-size: 12px; color: #9a9a94;")
            fl.addWidget(hint)

        btn_style = """
            QPushButton { background: transparent; color: #4a4a46; border: 1px solid #d8d8d4;
                         border-radius: 4px; padding: 6px 14px; font-size: 13px; }
            QPushButton:hover { background: #f2f2ef; border-color: #c3c3bf; }
        """
        btn_danger = """
            QPushButton { background: transparent; color: #a33a3a; border: 1px solid #e5c6c6;
                         border-radius: 4px; padding: 6px 14px; font-size: 13px; }
            QPushButton:hover { background: #fdf1f1; border-color: #d9b3b3; }
        """

        # 操作栏
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        btn_csv = QPushButton("导入 CSV / Excel")
        btn_csv.setStyleSheet(btn_style)
        btn_csv.clicked.connect(self._import_csv)
        btn_row.addWidget(btn_csv)

        btn_paste = QPushButton("粘贴数据")
        btn_paste.setStyleSheet(btn_style)
        btn_paste.setToolTip("从剪贴板粘贴表格数据，首行可作列名")
        btn_paste.clicked.connect(self._paste_clipboard)
        btn_row.addWidget(btn_paste)

        btn_row.addSpacing(8)

        btn_add = QPushButton("＋ 添加行")
        btn_add.setStyleSheet(btn_style)
        btn_add.clicked.connect(self._add_batch_row)
        btn_row.addWidget(btn_add)

        btn_del = QPushButton("− 删除选中行")
        btn_del.setStyleSheet(btn_danger)
        btn_del.clicked.connect(self._remove_batch_row)
        btn_row.addWidget(btn_del)

        if self.defaults:
            btn_apply = QPushButton("应用默认值")
            btn_apply.setStyleSheet(btn_style)
            btn_apply.clicked.connect(self._apply_defaults)
            btn_row.addWidget(btn_apply)

        btn_row.addStretch()
        fl.addLayout(btn_row)

        # 批量表格
        table = QTableWidget()
        table.setColumnCount(len(tags))
        headers = [self.window._resolve_label(t) for t in tags]
        table.setHorizontalHeaderLabels(headers)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setMinimumSectionSize(60)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setAlternatingRowColors(False)
        table.setStyleSheet("""
            QTableWidget { border: 1px solid #e3e3e0; border-radius: 3px;
                          gridline-color: #efefec; font-size: 12.5px; background: #ffffff; }
            QTableWidget::item { padding: 4px 6px; }
            QTableWidget::item:selected { background: #eceff5; color: #2c2c2a; }
            QHeaderView::section { background: #f7f7f5; font-weight: 500; color: #6b6b66;
                                  padding: 6px; border: none;
                                  border-right: 1px solid #efefec; border-bottom: 1px solid #e3e3e0; }
        """)
        table.verticalHeader().setDefaultSectionSize(32)
        table.setMinimumHeight(240)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        # 初始 3 个空行（含默认值）
        for _ in range(3):
            self._add_batch_row_to(table, self.defaults)

        self.table = table
        fl.addWidget(table)

        # 循环区域提示（批量模式暂不支持）
        if loops:
            loop_warn = QLabel(
                "该模板含循环区域（{% for %}），批量模式不支持循环数据，生成时将跳过。")
            loop_warn.setStyleSheet("font-size: 12px; color: #a33a3a;")
            loop_warn.setWordWrap(True)
            fl.addWidget(loop_warn)

        fl.addStretch()

    # ── 表格操作 ────────────────────────────────────────────

    def _add_batch_row_to(self, table, defaults=None):
        """向指定表格添加一行，可选填入默认值"""
        row = table.rowCount()
        table.insertRow(row)
        table.setRowHeight(row, 32)
        if defaults:
            for tag, val in defaults.items():
                if tag in self.tags:
                    col = self.tags.index(tag)
                    item = QTableWidgetItem(val)
                    table.setItem(row, col, item)

    def _add_batch_row(self):
        if self.table:
            self._add_batch_row_to(self.table, self.defaults)

    def _apply_defaults(self):
        """将默认值应用到所有行"""
        if not self.defaults or not self.table:
            return
        table = self.table
        for row in range(table.rowCount()):
            for tag, val in self.defaults.items():
                if tag in self.tags:
                    col = self.tags.index(tag)
                    existing = table.item(row, col)
                    if not existing or not existing.text().strip():
                        item = QTableWidgetItem(val)
                        table.setItem(row, col, item)
        self.window.status_bar.showMessage(f"已将 {len(self.defaults)} 个默认值应用到所有空单元格")

    def _remove_batch_row(self):
        if not self.table:
            return
        rows = set()
        for item in self.table.selectedItems():
            rows.add(item.row())
        for row in sorted(rows, reverse=True):
            self.table.removeRow(row)

    # ── 数据导入 — 表头匹配公共逻辑 ─────────────────────────

    def _map_headers_to_tags(self, header_row):
        """将导入数据的表头行映射到模板标签的列索引。

        匹配策略：
        1. 优先按标签名/显示名精确匹配
        2. 未匹配的回退为按列顺序一一对应

        返回 (tag_to_col: dict, data_start: int)
        """
        tag_to_col = {}
        for tag in self.tags:
            label_lower = self.window._resolve_label(tag).lower()
            for col_idx, h in enumerate(header_row):
                h_clean = h.strip().lower()
                h_normalized = h_clean.replace(' ', '_').replace('（', '').replace('）', '')
                if h_normalized == tag.lower() or h_clean == label_lower:
                    tag_to_col[tag] = col_idx
                    break

        if not tag_to_col:
            # 表头不匹配，按列顺序对应
            for idx, tag in enumerate(self.tags):
                if idx < len(header_row):
                    tag_to_col[tag] = idx

        data_start = 1 if tag_to_col else 0
        return tag_to_col, data_start

    # ── CSV / Excel 导入 ─────────────────────────────────────

    def _import_csv(self):
        """从 CSV 或 Excel 文件导入批量数据"""
        path, _ = QFileDialog.getOpenFileName(
            self.window, "导入数据文件", self.window._last_dir,
            "表格文件 (*.csv *.xlsx *.xls);;CSV 文件 (*.csv);;Excel 工作簿 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        if not path:
            return
        self.window._last_dir = os.path.dirname(path)

        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in ('.xlsx', '.xls'):
                import openpyxl
                wb = openpyxl.load_workbook(path, data_only=True)
                ws = wb.active
                rows_data = []
                for row in ws.iter_rows(values_only=True):
                    rows_data.append([str(v) if v is not None else '' for v in row])
                wb.close()
            else:
                rows_data = []
                for encoding in ['utf-8-sig', 'utf-8', 'gbk', 'gb2312']:
                    try:
                        with open(path, 'r', encoding=encoding) as f:
                            reader = csv.reader(f)
                            rows_data = [list(row) for row in reader]
                        break
                    except (UnicodeDecodeError, Exception):
                        continue
        except Exception as e:
            QMessageBox.critical(self.window, "导入失败", f"无法读取文件：{e}")
            return

        if not rows_data:
            QMessageBox.warning(self.window, "提示", "文件中没有数据。")
            return

        tag_to_col, data_start = self._map_headers_to_tags(rows_data[0])
        self._populate_table(rows_data[data_start:], tag_to_col)
        self.window.status_bar.showMessage(f"已导入 {self.table.rowCount()} 行数据")

    # ── 剪贴板粘贴 ──────────────────────────────────────────

    def _paste_clipboard(self):
        """从剪贴板粘贴表格数据（Tab/逗号分隔）"""
        clip = QApplication.clipboard()
        text = clip.text()
        if not text:
            return

        lines = text.strip().split('\n')
        rows_data = []
        for line in lines:
            if '\t' in line:
                row = line.split('\t')
            elif ',' in line:
                row = line.split(',')
            else:
                row = [line]
            rows_data.append([v.strip() for v in row])

        if not rows_data:
            return

        tag_to_col, data_start = self._map_headers_to_tags(rows_data[0])
        self._populate_table(rows_data[data_start:], tag_to_col)
        self.window.status_bar.showMessage(f"已粘贴 {self.table.rowCount()} 行数据")

    def _populate_table(self, rows_data, tag_to_col):
        """用 rows_data 填充批量表格"""
        table = self.table
        table.setRowCount(0)

        for row_data in rows_data:
            if not any(v.strip() for v in row_data):
                continue
            row = table.rowCount()
            table.insertRow(row)
            table.setRowHeight(row, 32)
            for tag, col_idx in tag_to_col.items():
                if col_idx < len(row_data):
                    value = row_data[col_idx].strip()
                    if value:
                        tag_col = self.tags.index(tag)
                        item = QTableWidgetItem(value)
                        table.setItem(row, tag_col, item)

    # ── 批量生成 ────────────────────────────────────────────

    def _on_batch_generate(self):
        """批量生成：每行数据生成一个独立文件"""
        table = self.table
        if not table or table.rowCount() == 0:
            QMessageBox.warning(self.window, "提示", "批量表格中没有数据行。")
            return

        row_contexts = []
        for row in range(table.rowCount()):
            ctx = {}
            has_data = False
            for col in range(table.columnCount()):
                tag = self.tags[col]
                cell_item = table.item(row, col)
                value = cell_item.text().strip() if cell_item else ''
                if value:
                    has_data = True
                ctx[tag] = value
            if has_data:
                row_contexts.append(ctx)

        if not row_contexts:
            QMessageBox.warning(self.window, "提示", "所有数据行均为空。")
            return

        # 输出位置：沿用「生成导出」页设置，留空则桌面
        out_dir = ''
        prefix = ''
        try:
            out_dir = self.window.out_dir_edit.text().strip()
            prefix = self.window.out_name_edit.text().strip()
        except AttributeError:
            pass
        if not out_dir or not os.path.isdir(out_dir):
            out_dir = os.path.join(os.path.expanduser("~"), "Desktop")

        reply = QMessageBox.question(
            self.window, "确认批量生成",
            f"将为 {len(row_contexts)} 行数据各生成一个文件。\n\n保存位置：{out_dir}\n\n确定继续？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        from datetime import datetime
        from doc_filler import fill_template
        from xlsx_filler import fill_template as fill_xlsx_template

        fill_func = fill_xlsx_template if self.window._template_format == 'xlsx' else fill_template

        success = 0
        errors = []
        ts_base = datetime.now()
        ts_short = ts_base.strftime('%H%M%S')
        base = prefix or os.path.splitext(os.path.basename(self.window._template_path))[0]

        for idx, ctx in enumerate(row_contexts):
            try:
                ext = self.window._template_format or 'docx'
                output_path = os.path.join(out_dir, f"{base}_{idx + 1:03d}_{ts_short}.{ext}")
                fill_func(self.window._template_path, ctx, output_path)
                success += 1
            except Exception as e:
                errors.append(f"第 {idx + 1} 行: {e}")

        # 批量数据也写入模板记忆，下次打开自动带出
        self.window._remember()

        if errors:
            QMessageBox.warning(
                self.window, "批量生成完成",
                f"成功：{success} 个文件\n失败：{len(errors)} 个\n\n"
                + "\n".join(errors[:5])
            )
        else:
            QMessageBox.information(
                self.window, "批量生成完成",
                f"已生成 {success} 个文件，保存在：\n{out_dir}"
            )
            self.window.status_bar.showMessage(f"批量生成完成：{success} 个文件 → {out_dir}")
