import sys
import os
import calendar
from datetime import datetime, timedelta
import jpholiday
import shiboken6
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *

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
            cols = 4 # 1行に4つ並べる
            for i, (name, code) in enumerate(colors):
                btn = QPushButton()
                btn.setFixedSize(80, 30)
                btn.setToolTip(name)
                btn.setStyleSheet(f"background-color: {code}; border: 1px solid #ccc;")
                btn.clicked.connect(lambda checked, c=code: self.select_color(c))
                
                # 色名も表示したい場合は重ねるか下に置く
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
        self.tabs.setCurrentIndex(1) # デフォルトを月間集計にする
        
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
        
        # 期間ごとの全グループ合計を保持するリスト
        total_period_maps = [{} for _ in range(len(headers))]
        
        for r, g in enumerate(group_data):
            table.setItem(r, 0, QTableWidgetItem(g['name']))
            for c, (h_start, h_end, label) in enumerate(headers):
                color_map = {}
                for t in g['tasks']:
                    for p in t.get('periods', []):
                        if not p.get('start_date') or not p.get('end_date'): continue
                        psd = datetime.strptime(p['start_date'], "%Y-%m-%d")
                        ped = datetime.strptime(p['end_date'], "%Y-%m-%d")
                        overlap = (min(ped, h_end) - max(psd, h_start)).days + 1
                        if overlap > 0:
                            c_code = t.get('color', '#0078d4')
                            # 人数を考慮した集計
                            val = overlap * t.get('person_count', 1)
                            color_map[c_code] = color_map.get(c_code, 0) + val
                            # 全体合計に加算
                            total_period_maps[c][c_code] = total_period_maps[c].get(c_code, 0) + val
                
                if not color_map:
                    item = QTableWidgetItem("-")
                else:
                    text = self.app.format_total_days(color_map)
                    item = QTableWidgetItem(text)
                
                item.setTextAlignment(Qt.AlignCenter)
                if color_map:
                    item.setForeground(QColor(0, 120, 212))
                    f = item.font(); f.setBold(True); item.setFont(f)
                table.setItem(r, c + 1, item)
        
        # 全体合計行の作成
        total_row_idx = len(group_data)
        total_label_item = QTableWidgetItem("全体合計")
        total_label_item.setBackground(QColor(240, 248, 255))
        f = total_label_item.font(); f.setBold(True); total_label_item.setFont(f)
        table.setItem(total_row_idx, 0, total_label_item)
        
        for c, color_map in enumerate(total_period_maps):
            text = self.app.format_total_days(color_map) if color_map else "-"
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
        self.resize(850, 650)
        layout = QVBoxLayout(self)
        
        self.browser = QTextBrowser()
        self.browser.setOpenExternalLinks(True) # リンクをブラウザで開けるようにする
        
        # README.md を読み込む
        readme_path = self.get_readme_path()
        if os.path.exists(readme_path):
            try:
                with open(readme_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                self.browser.setMarkdown(content)
            except Exception as e:
                self.browser.setText(f"README.md の読み込みに失敗しました: {e}")
        else:
            self.browser.setText(f"README.md が見つかりませんでした。\n検索パス: {readme_path}")
            
        layout.addWidget(self.browser)
        
        btn_close = QPushButton("閉じる")
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close)

    def get_readme_path(self):
        if hasattr(sys, '_MEIPASS'):
            # PyInstallerでパッケージ化された場合
            return os.path.join(sys._MEIPASS, 'README.md')
        # ソースから実行された場合
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'README.md')

