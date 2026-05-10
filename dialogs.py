# tsukasamiyashita/miyagantt/MiyaGantt-537e3273fbb3bd08e66a6a82ad13431ee8f2e49c/dialogs.py
import os
import sys
import calendar
import re
import markdown
from datetime import datetime, timedelta
from PySide6.QtWidgets import (QDialog, QFormLayout, QDateEdit, QPushButton, 
                               QVBoxLayout, QHBoxLayout, QScrollArea, QWidget, QLabel, 
                               QGridLayout, QTabWidget, QTableWidget, 
                               QAbstractItemView, QTableWidgetItem, QTextBrowser,
                               QListWidget, QListWidgetItem, QDialogButtonBox, QComboBox, QFileDialog, QMessageBox, QSplitter)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor, QFont, QPageLayout
from PySide6.QtPrintSupport import QPrinterInfo, QPrintPreviewWidget, QPageSetupDialog, QPrintDialog, QPrinter

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("編集期間")
        self.layout = QFormLayout(self)
        
        self.start_date_edit = QDateEdit(self)
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDate(parent.min_date.date() if parent else QDate.currentDate())
        
        self.end_date_edit = QDateEdit(self)
        self.end_date_edit.setCalendarPopup(True)
        if parent:
            end_date = parent.min_date + timedelta(days=parent.display_days - 1)
            self.end_date_edit.setDate(end_date.date())
        else:
            self.end_date_edit.setDate(QDate.currentDate().addDays(30))
        
        self.layout.addRow("開始日:", self.start_date_edit)
        self.layout.addRow("終了日:", self.end_date_edit)
        
        self.btn_ok = QPushButton("OK", self)
        self.btn_ok.clicked.connect(self.accept)
        self.layout.addRow(self.btn_ok)

class ColorGridDialog(QDialog):
    def __init__(self, color_groups, parent=None):
        super().__init__(parent)
        self.setWindowTitle("色を選択")
        self.selected_color = None
        
        main_layout = QVBoxLayout(self)
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        container = QWidget()
        container_layout = QVBoxLayout(container)
        
        for group_name, colors in color_groups:
            group_label = QLabel(group_name)
            group_label.setStyleSheet("font-weight: bold; margin-top: 10px; color: #555;")
            container_layout.addWidget(group_label)
            
            grid = QGridLayout()
            grid.setSpacing(4)
            cols = 4 
            for i, (name, code) in enumerate(colors):
                btn = QPushButton()
                btn.setFixedSize(80, 30)
                btn.setToolTip(name)
                btn.setStyleSheet(f"background-color: {code}; border: 1px solid #ccc;")
                btn.clicked.connect(lambda checked, c=code: self.select_color(c))
                
                grid.addWidget(btn, i // cols, i % cols)
            container_layout.addLayout(grid)
            
        scroll.setWidget(container)
        main_layout.addWidget(scroll)
        
        self.resize(380, 500)

    def select_color(self, color):
        self.selected_color = color
        self.accept()

class SummaryDialog(QDialog):
    def __init__(self, app, tasks, parent=None):
        super().__init__(parent)
        self.app = app
        self.setWindowTitle("グループ別集計レポート")
        self.resize(1000, 600)
        self.tasks = tasks
        
        layout = QVBoxLayout(self)
        self.tabs = QTabWidget(self)
        
        self.weekly_table = QTableWidget()
        self.monthly_table = QTableWidget()
        self.yearly_table = QTableWidget()
        
        for table in [self.weekly_table, self.monthly_table, self.yearly_table]:
            table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            table.setAlternatingRowColors(True)
            table.setStyleSheet("QTableWidget { background-color: #ffffff; gridline-color: #e0e0e0; }")
        
        self.tabs.addTab(self.weekly_table, "週間集計")
        self.tabs.addTab(self.monthly_table, "月間集計")
        self.tabs.addTab(self.yearly_table, "年間集計")
        self.tabs.setCurrentIndex(1) 
        
        layout.addWidget(self.tabs)
        
        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        
        self.refresh_data()

    def refresh_data(self):
        group_data = []
        current_group_list = []
        current_group_name = "未分類"
        
        for t in self.tasks:
            if t.get('is_group'):
                if current_group_list:
                    group_data.append({'name': current_group_name, 'tasks': current_group_list})
                current_group_name = t.get('name', '無題グループ')
                current_group_list = []
            else:
                current_group_list.append(t)
        
        if current_group_list:
            group_data.append({'name': current_group_name, 'tasks': current_group_list})
            
        if not group_data:
            return

        all_dates = []
        for g in group_data:
            for t in g['tasks']:
                for p in t.get('periods', []):
                    if p.get('start_date'): all_dates.append(datetime.strptime(p['start_date'], "%Y-%m-%d"))
                    if p.get('end_date'): all_dates.append(datetime.strptime(p['end_date'], "%Y-%m-%d"))
        
        if not all_dates:
            start_range = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            end_range = start_range + timedelta(days=30)
        else:
            start_range = min(all_dates)
            end_range = max(all_dates)

        self.fill_table(self.weekly_table, group_data, start_range, end_range, 'week')
        self.fill_table(self.monthly_table, group_data, start_range, end_range, 'month')
        self.fill_table(self.yearly_table, group_data, start_range, end_range, 'year')

    def fill_table(self, table, group_data, start_range, end_range, unit):
        headers = []
        curr = start_range
        if unit == 'week':
            curr = curr - timedelta(days=curr.weekday())
            while curr <= end_range:
                headers.append((curr, curr + timedelta(days=6), curr.strftime("%m/%d~")))
                curr += timedelta(days=7)
        elif unit == 'month':
            curr = curr.replace(day=1)
            while curr <= end_range:
                last_day = calendar.monthrange(curr.year, curr.month)[1]
                headers.append((curr, curr.replace(day=last_day), curr.strftime("%Y/%m")))
                m = curr.month + 1
                y = curr.year
                if m > 12: m = 1; y += 1
                curr = datetime(y, m, 1)
        elif unit == 'year':
            curr = curr.replace(month=1, day=1)
            while curr <= end_range:
                headers.append((curr, curr.replace(month=12, day=31), curr.strftime("%Y年")))
                curr = curr.replace(year=curr.year + 1)

        table.setColumnCount(len(headers) + 1)
        table.setRowCount(len(group_data) + 1)
        table.setHorizontalHeaderLabels(["グループ名"] + [h[2] for h in headers])
        
        total_period_maps = [{} for _ in range(len(headers))]
        
        for r, g in enumerate(group_data):
            table.setItem(r, 0, QTableWidgetItem(g['name']))
            for c, (h_start, h_end, label) in enumerate(headers):
                color_map = {}
                for t in g['tasks']:
                    if t.get('mode') in ['auto', 'memo']:
                        continue
                    t_color = t.get('color', '#0000ff')
                    hc = float(t.get('headcount', 1.0)) * float(t.get('efficiency', 1.0))

                    for p in t.get('periods', []):
                        if not p.get('start_date') or not p.get('end_date'): continue
                        try:
                            psd = datetime.strptime(p['start_date'], "%Y-%m-%d")
                            ped = datetime.strptime(p['end_date'], "%Y-%m-%d")
                            calc_start = max(psd, h_start)
                            calc_end = min(ped, h_end)
                            if calc_start <= calc_end:
                                overlap = (calc_end - calc_start).days + 1
                                if overlap > 0:
                                    p_color = p.get('color')
                                    if p_color and p_color.lower() != t_color.lower():
                                        continue
                                    val = overlap * hc
                                    color_map[t_color] = color_map.get(t_color, 0) + val
                                    total_period_maps[c][t_color] = total_period_maps[c].get(t_color, 0) + val
                        except ValueError:
                            pass
                
                if not color_map:
                    item = QTableWidgetItem("-")
                else:
                    text = self.app.format_summary_workload(color_map)
                    item = QTableWidgetItem(text)
                
                item.setTextAlignment(Qt.AlignCenter)
                if color_map:
                    item.setForeground(QColor(0, 0, 255))
                    f = item.font(); f.setBold(True); item.setFont(f)
                table.setItem(r, c + 1, item)
        
        total_row_idx = len(group_data)
        total_label_item = QTableWidgetItem("全体合計")
        total_label_item.setBackground(QColor(240, 248, 255))
        f = total_label_item.font(); f.setBold(True); total_label_item.setFont(f)
        table.setItem(total_row_idx, 0, total_label_item)
        
        for c, color_map in enumerate(total_period_maps):
            text = self.app.format_summary_workload(color_map) if color_map else "-"
            item = QTableWidgetItem(text)
            item.setTextAlignment(Qt.AlignCenter)
            item.setBackground(QColor(240, 248, 255))
            f = item.font(); f.setBold(True); item.setFont(f)
            table.setItem(total_row_idx, c + 1, item)
        
        table.resizeColumnsToContents()

class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("ヘルプ - MiyaGantt")
        self.resize(800, 700)
        
        self.setStyleSheet("QDialog { background-color: #ffffff; }")
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 15)
        layout.setSpacing(10)
        
        self.browser = QTextBrowser()
        self.browser.setOpenExternalLinks(True) 
        self.browser.setStyleSheet("""
            QTextBrowser {
                border: none;
                background-color: #ffffff;
            }
        """)
        layout.addWidget(self.browser)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_close = QPushButton("閉じる")
        btn_close.setMinimumWidth(120)
        btn_close.setMinimumHeight(35)
        btn_close.setStyleSheet("""
            QPushButton {
                background-color: #005fb8;
                border: none;
                border-radius: 4px;
                font-weight: bold;
                color: #ffffff;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #004a98;
            }
        """)
        btn_close.clicked.connect(self.accept)
        btn_layout.addWidget(btn_close)
        layout.addLayout(btn_layout)

        self.load_readme()

    def load_readme(self):
        readme_path = self.get_readme_path()
        if not os.path.exists(readme_path):
            self.browser.setHtml(f"<p>README.md が見つかりませんでした。<br>検索パス: {readme_path}</p>")
            return

        try:
            with open(readme_path, 'r', encoding='utf-8') as f:
                content = f.read()
                
            html_body = self.convert_markdown_to_styled_html(content)
            self.browser.setHtml(html_body)
                
        except Exception as e:
            self.browser.setHtml(f"<p>README.md の読み込みに失敗しました: {e}</p>")

    def convert_markdown_to_styled_html(self, md_text):
        html = md_text
        try:
            html = markdown.markdown(md_text, extensions=['tables', 'fenced_code'])
        except ImportError:
            html = re.sub(r'^### (.*)', r'<h3>\1</h3>', html, flags=re.MULTILINE)
            html = re.sub(r'^## (.*)', r'<h2>\1</h2>', html, flags=re.MULTILINE)
            html = re.sub(r'^# (.*)', r'<h1>\1</h1>', html, flags=re.MULTILINE)
            html = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', html)
            html = re.sub(r'^- (.*)', r'<li>\1</li>', html, flags=re.MULTILINE)
            html = html.replace('</li>\n<li>', '</li><li>')
            html = re.sub(r'(<li>.*</li>)', r'<ul>\1</ul>', html, flags=re.DOTALL)
            html = html.replace('\n\n', '<br><br>')

        styled_html = f"""
        <html>
        <head>
        <style>
            body {{
                font-family: 'Segoe UI', Meiryo, sans-serif;
                font-size: 14.5px;
                color: #323130;
                line-height: 1.7;
                background-color: #ffffff;
                margin: 0;
                padding: 10px;
            }}
            h1 {{
                font-size: 26px;
                color: #005fb8;
                border-bottom: 2px solid #edebe9;
                padding-bottom: 10px;
                margin-top: 5px;
                margin-bottom: 25px;
                font-weight: normal;
            }}
            h2 {{
                font-size: 18px;
                color: #201f1e;
                background-color: #f8f8f8;
                border-left: 5px solid #005fb8;
                padding: 8px 15px;
                margin-top: 35px;
                margin-bottom: 20px;
                font-weight: 600;
            }}
            h3 {{
                font-size: 16px;
                color: #323130;
                margin-top: 25px;
                margin-bottom: 10px;
                font-weight: 600;
            }}
            ul {{
                margin-top: 5px;
                margin-bottom: 20px;
                padding-left: 25px;
            }}
            li {{
                margin-bottom: 10px;
            }}
            p {{
                margin-top: 0;
                margin-bottom: 15px;
            }}
            strong {{
                font-weight: bold;
                color: #201f1e;
            }}
            blockquote {{
                background-color: #fff9e6;
                border-left: 4px solid #ffcc00;
                padding: 12px 20px;
                margin: 20px 0;
                color: #323130;
            }}
        </style>
        </head>
        <body>
            {html}
        </body>
        </html>
        """
        return styled_html

    def get_readme_path(self):
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, 'README.md')
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'README.md')

class PrintSettingsDialog(QDialog):
    def __init__(self, parent=None, tasks_info=None, min_date=None, max_date=None):
        super().__init__(parent)
        self.setWindowTitle("印刷対象の設定")
        self.resize(600, 500)
        self.tasks_info = tasks_info or []
        self.parent_table = parent.table if parent and hasattr(parent, 'table') else None
        
        layout = QVBoxLayout(self)
        
        # 期間設定
        date_layout = QHBoxLayout()
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        if min_date:
            self.start_date_edit.setDate(QDate(min_date.year, min_date.month, min_date.day))
            
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        if max_date:
            self.end_date_edit.setDate(QDate(max_date.year, max_date.month, max_date.day))
            
        date_layout.addWidget(QLabel("印刷期間:"))
        date_layout.addWidget(self.start_date_edit)
        date_layout.addWidget(QLabel("～"))
        date_layout.addWidget(self.end_date_edit)
        layout.addLayout(date_layout)
        
        # 行と列のリストを横に並べるためのレイアウト
        lists_layout = QHBoxLayout()
        
        # タスク(行)選択
        row_layout = QVBoxLayout()
        row_layout.addWidget(QLabel("印刷対象の行:"))
        self.task_list = QListWidget()
        for i, info in enumerate(self.tasks_info):
            t = info['task']
            name = t.get('name', '無題')
            indent = "  " * info.get('indent', 0)
            item = QListWidgetItem(f"{indent}{name}")
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            is_hidden = self.parent_table.isRowHidden(i) if self.parent_table else False
            item.setCheckState(Qt.Unchecked if is_hidden else Qt.Checked)
            item.setData(Qt.UserRole, i)
            self.task_list.addItem(item)
        row_layout.addWidget(self.task_list)
        
        btn_row_layout = QHBoxLayout()
        btn_row_all = QPushButton("全選択")
        btn_row_clear = QPushButton("全解除")
        btn_row_all.clicked.connect(lambda: self.set_all_checked(self.task_list, Qt.Checked))
        btn_row_clear.clicked.connect(lambda: self.set_all_checked(self.task_list, Qt.Unchecked))
        btn_row_layout.addWidget(btn_row_all)
        btn_row_layout.addWidget(btn_row_clear)
        row_layout.addLayout(btn_row_layout)
        
        lists_layout.addLayout(row_layout)
        
        # 列選択
        col_layout = QVBoxLayout()
        col_layout.addWidget(QLabel("印刷対象の列:"))
        self.col_list = QListWidget()
        if self.parent_table:
            for i in range(self.parent_table.columnCount()):
                header_item = self.parent_table.horizontalHeaderItem(i)
                header_text = header_item.text().strip() if header_item and header_item.text().strip() else ""
                if not header_text:
                    if i == 0: header_text = "選択マーク"
                    elif i == 1: header_text = "開閉ボタン"
                    else: header_text = f"列 {i}"
                    
                item = QListWidgetItem(header_text)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                is_hidden = self.parent_table.isColumnHidden(i)
                item.setCheckState(Qt.Unchecked if is_hidden else Qt.Checked)
                item.setData(Qt.UserRole, i)
                self.col_list.addItem(item)
        col_layout.addWidget(self.col_list)
        
        btn_col_layout = QHBoxLayout()
        btn_col_all = QPushButton("全選択")
        btn_col_clear = QPushButton("全解除")
        btn_col_all.clicked.connect(lambda: self.set_all_checked(self.col_list, Qt.Checked))
        btn_col_clear.clicked.connect(lambda: self.set_all_checked(self.col_list, Qt.Unchecked))
        btn_col_layout.addWidget(btn_col_all)
        btn_col_layout.addWidget(btn_col_clear)
        col_layout.addLayout(btn_col_layout)
        
        lists_layout.addLayout(col_layout)
        layout.addLayout(lists_layout)
        
        # ボタン
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
    def set_all_checked(self, list_widget, state):
        for i in range(list_widget.count()):
            list_widget.item(i).setCheckState(state)
            
    def get_settings(self):
        sd_qdate = self.start_date_edit.date()
        ed_qdate = self.end_date_edit.date()
        sd = datetime(sd_qdate.year(), sd_qdate.month(), sd_qdate.day())
        ed = datetime(ed_qdate.year(), ed_qdate.month(), ed_qdate.day())
        
        selected_row_indices = []
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            if item.checkState() == Qt.Checked:
                selected_row_indices.append(item.data(Qt.UserRole))
                
        selected_col_indices = []
        for i in range(self.col_list.count()):
            item = self.col_list.item(i)
            if item.checkState() == Qt.Checked:
                selected_col_indices.append(item.data(Qt.UserRole))
                
        return sd, ed, selected_row_indices, selected_col_indices

class CustomPrintPreviewDialog(QDialog):
    def __init__(self, printer, render_func, parent=None):
        super().__init__(parent)
        self.setWindowTitle("印刷プレビュー・プリンター設定")
        self.setWindowState(Qt.WindowMaximized)
        self.printer = printer
        self.render_func = render_func
        
        layout = QVBoxLayout(self)
        
        ctrl_layout = QHBoxLayout()
        
        ctrl_layout.addWidget(QLabel("プリンター:"))
        self.printer_combo = QComboBox()
        for p in QPrinterInfo.availablePrinters():
            self.printer_combo.addItem(p.printerName(), p)
        idx = self.printer_combo.findText(printer.printerName())
        if idx >= 0:
            self.printer_combo.setCurrentIndex(idx)
        self.printer_combo.currentIndexChanged.connect(self.change_printer)
        ctrl_layout.addWidget(self.printer_combo)
        
        btn_page_setup = QPushButton("用紙/余白の設定")
        btn_page_setup.clicked.connect(self.open_page_setup)
        ctrl_layout.addWidget(btn_page_setup)
        
        btn_fit_width = QPushButton("幅に合わせる")
        btn_fit_height = QPushButton("高さに合わせる")
        btn_fit_page = QPushButton("ページ全体")
        
        ctrl_layout.addWidget(btn_fit_width)
        ctrl_layout.addWidget(btn_fit_height)
        ctrl_layout.addWidget(btn_fit_page)
        
        ctrl_layout.addStretch()

        btn_pdf = QPushButton("📥 PDF出力")
        btn_pdf.clicked.connect(self.export_pdf)
        ctrl_layout.addWidget(btn_pdf)
        
        btn_print = QPushButton("🖨 印刷実行")
        btn_print.setStyleSheet("background-color: #0078d4; color: white; font-weight: bold; padding: 6px 15px;")
        btn_print.clicked.connect(self.direct_print)
        ctrl_layout.addWidget(btn_print)
        
        layout.addLayout(ctrl_layout)
        
        self.preview_widget = QPrintPreviewWidget(self.printer, self)
        self.preview_widget.paintRequested.connect(self.render_func)
        layout.addWidget(self.preview_widget)
        
        btn_fit_width.clicked.connect(lambda: self.preview_widget.fitToWidth())
        btn_fit_height.clicked.connect(lambda: self.preview_widget.setZoomMode(QPrintPreviewWidget.FitToHeight) if hasattr(QPrintPreviewWidget, 'FitToHeight') else None)
        btn_fit_page.clicked.connect(lambda: self.preview_widget.fitInView())
        
    def change_printer(self):
        p_info = self.printer_combo.currentData()
        if p_info:
            self.printer.setPrinterName(p_info.printerName())
            self.preview_widget.updatePreview()
            
    def open_page_setup(self):
        dlg = QPageSetupDialog(self.printer, self)
        if dlg.exec():
            self.preview_widget.updatePreview()

    def export_pdf(self):
        filename, _ = QFileDialog.getSaveFileName(self, "PDFとして保存", "", "PDF Files (*.pdf)")
        if filename:
            original_printer_name = self.printer.printerName()
            original_output_format = self.printer.outputFormat()
            original_output_file = self.printer.outputFileName()

            self.printer.setOutputFormat(QPrinter.PdfFormat)
            self.printer.setOutputFileName(filename)
            
            self.render_func(self.printer)
            
            self.printer.setOutputFormat(original_output_format)
            self.printer.setOutputFileName(original_output_file)
            self.printer.setPrinterName(original_printer_name)
            
            QMessageBox.information(self, "成功", f"PDFを出力しました:\n{filename}")
            self.accept()
            
    def direct_print(self):
        print_dlg = QPrintDialog(self.printer, self)
        if print_dlg.exec():
            self.render_func(self.printer)
            self.accept()