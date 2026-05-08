# tsukasamiyashita/miyagantt/MiyaGantt-46a1664b6d1737cb32f1dd17429ce06cca8dc678/dialogs.py
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                               QDateEdit, QGridLayout, QPushButton, QTableWidget, 
                               QTableWidgetItem, QHeaderView, QTextEdit, QListWidget, QListWidgetItem, QDialogButtonBox)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor, QFont
from datetime import datetime, timedelta

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("カレンダー編集期間の設定")
        
        layout = QVBoxLayout(self)
        
        form_layout = QHBoxLayout()
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDate(QDate(parent.min_date.year, parent.min_date.month, parent.min_date.day))
        
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        if hasattr(parent, 'max_date') and parent.max_date:
            self.end_date_edit.setDate(QDate(parent.max_date.year, parent.max_date.month, parent.max_date.day))
        else:
            display_days = parent.display_days if hasattr(parent, 'display_days') else 180
            ed = parent.min_date + timedelta(days=display_days - 1)
            self.end_date_edit.setDate(QDate(ed.year, ed.month, ed.day))
            
        form_layout.addWidget(QLabel("開始日:"))
        form_layout.addWidget(self.start_date_edit)
        form_layout.addWidget(QLabel("～ 終了日:"))
        form_layout.addWidget(self.end_date_edit)
        
        layout.addLayout(form_layout)
        
        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("OK")
        btn_cancel = QPushButton("キャンセル")
        btn_ok.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        
        layout.addLayout(btn_layout)

class ColorGridDialog(QDialog):
    def __init__(self, color_groups, parent=None):
        super().__init__(parent)
        self.setWindowTitle("色の選択")
        self.selected_color = None
        
        layout = QVBoxLayout(self)
        grid = QGridLayout()
        grid.setSpacing(8)
        
        row = 0
        for group_name, colors in color_groups:
            label = QLabel(group_name)
            label.setFont(QFont("Segoe UI", 9, QFont.Bold))
            grid.addWidget(label, row, 0, 1, 6)
            row += 1
            
            col = 0
            for name, hex_code in colors:
                btn = QPushButton()
                btn.setFixedSize(30, 30)
                btn.setStyleSheet(f"background-color: {hex_code}; border: 1px solid #999; border-radius: 4px;")
                btn.setToolTip(name)
                btn.clicked.connect(lambda checked, c=hex_code: self.select_color(c))
                grid.addWidget(btn, row, col)
                col += 1
                if col > 5:
                    col = 0
                    row += 1
            if col != 0:
                row += 1
                
        layout.addLayout(grid)
        
    def select_color(self, hex_code):
        self.selected_color = hex_code
        self.accept()

class SummaryDialog(QDialog):
    def __init__(self, parent=None, tasks=None, app=None):
        super().__init__(parent)
        self.setWindowTitle("工数集計")
        self.resize(600, 400)
        self.tasks = tasks or []
        self.app = app
        
        layout = QVBoxLayout(self)
        
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["色 / グループ", "名前", "合計工数"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 120)
        self.table.setColumnWidth(2, 100)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.table)
        
        btn_close = QPushButton("閉じる")
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close)
        
        self.calculate_summary()
        
    def calculate_summary(self):
        summary = {}
        
        for t in self.tasks:
            if t.get('is_group') or t.get('mode') in ['auto', 'memo']:
                continue
                
            hc = t.get('headcount', 1.0) * t.get('efficiency', 1.0)
            t_color = t.get('color', '#0000ff')
            
            for p in t.get('periods', []):
                if not p.get('start_date') or not p.get('end_date'): continue
                
                try:
                    sd = datetime.strptime(p['start_date'], "%Y-%m-%d")
                    ed = datetime.strptime(p['end_date'], "%Y-%m-%d")
                    days = (ed - sd).days + 1
                    
                    if days > 0:
                        c = p.get('color')
                        if c and c.lower() != t_color.lower():
                            continue
                        summary[t_color] = summary.get(t_color, 0) + (days * hc)
                except ValueError:
                    pass
                    
        self.table.setRowCount(len(summary))
        for row, (color_code, total) in enumerate(summary.items()):
            color_name = self.app.get_color_name(color_code) if self.app else color_code
            
            item_c = QTableWidgetItem(color_code)
            item_c.setBackground(QColor(color_code))
            item_c.setForeground(Qt.white if QColor(color_code).lightness() < 128 else Qt.black)
            
            item_n = QTableWidgetItem(color_name)
            item_v = QTableWidgetItem(f"{total:g} 工数")
            item_v.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            
            self.table.setItem(row, 0, item_c)
            self.table.setItem(row, 1, item_n)
            self.table.setItem(row, 2, item_v)

class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("ヘルプ・使い方")
        self.resize(700, 500)
        
        layout = QVBoxLayout(self)
        
        text = QTextEdit()
        text.setReadOnly(True)
        text.setHtml("""
        <h2>MiyaGantt 使い方ガイド</h2>
        
        <h3>1. 基本的な操作</h3>
        <ul>
            <li><b>タスク追加:</b> ツールバーの「追加」ボタンで新規タスクを作成します。</li>
            <li><b>モードの選択:</b> 追加時にデフォルト設定から「人員(手動)」「案件(自動)」「メモ」を選べます。追加後もテーブルから変更可能です。</li>
            <li><b>期間の描画:</b> チャート上の空いている領域をドラッグすると、期間（バー）を簡単に作成できます。</li>
            <li><b>バーの移動・伸縮:</b> バーの端をドラッグで伸縮、中央をドラッグで移動できます。</li>
        </ul>
        
        <h3>2. 各モードの特徴</h3>
        <ul>
            <li><b>👤 人員モード (手動):</b> 人を中心としたスケジュール管理。指定した人数のリソースを提供します。バーの色を変えることで、案件モードの特定タスクにリソースを紐付けられます。</li>
            <li><b>⚡ 案件モード (自動):</b> 案件（タスク）中心の管理。開始日と必要な「合計工数」を設定すると、同じ色の「人員モード」で確保されたリソースを自動的に消化し、バーが自動描画されます。</li>
            <li><b>📝 メモモード:</b> 工数計算に影響を与えない、純粋なスケジュールやメモを表示します。</li>
        </ul>
        
        <h3>3. その他の便利機能</h3>
        <ul>
            <li><b>休日の設定:</b> カレンダーのヘッダー（日付部分）をクリックすると、その日を休日に設定・解除できます。</li>
            <li><b>色による紐付け:</b> 人員モードのバーの色と、案件モードのタスクの色を同じにすることで、自動的にリソースが割り当てられます。テーブルの「色」列をダブルクリック、またはバーを右クリックして変更できます。</li>
            <li><b>コメント:</b> チャート上を右クリックして「コメントを追加」できます。</li>
        </ul>
        """)
        
        layout.addWidget(text)
        
        btn_close = QPushButton("閉じる")
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close)

class PrintSettingsDialog(QDialog):
    def __init__(self, parent=None, tasks_info=None, min_date=None, max_date=None):
        super().__init__(parent)
        self.setWindowTitle("印刷設定")
        self.resize(400, 500)
        self.tasks_info = tasks_info or []
        
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
        
        # タスク選択
        layout.addWidget(QLabel("印刷対象の行:"))
        self.task_list = QListWidget()
        for i, info in enumerate(self.tasks_info):
            t = info['task']
            name = t.get('name', '無題')
            indent = "  " * info.get('indent', 0)
            item = QListWidgetItem(f"{indent}{name}")
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            # 現在の表示状態を初期値にする
            is_hidden = parent.table.isRowHidden(i) if parent and hasattr(parent, 'table') else False
            item.setCheckState(Qt.Unchecked if is_hidden else Qt.Checked)
            item.setData(Qt.UserRole, i)
            self.task_list.addItem(item)
        layout.addWidget(self.task_list)
        
        # 全選択/全解除
        btn_layout = QHBoxLayout()
        btn_all = QPushButton("全選択")
        btn_clear = QPushButton("全解除")
        btn_all.clicked.connect(lambda: self.set_all_checked(Qt.Checked))
        btn_clear.clicked.connect(lambda: self.set_all_checked(Qt.Unchecked))
        btn_layout.addWidget(btn_all)
        btn_layout.addWidget(btn_clear)
        layout.addLayout(btn_layout)
        
        # ボタン
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
    def set_all_checked(self, state):
        for i in range(self.task_list.count()):
            self.task_list.item(i).setCheckState(state)
            
    def get_settings(self):
        sd_qdate = self.start_date_edit.date()
        ed_qdate = self.end_date_edit.date()
        sd = datetime(sd_qdate.year(), sd_qdate.month(), sd_qdate.day())
        ed = datetime(ed_qdate.year(), ed_qdate.month(), ed_qdate.day())
        selected_indices = []
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            if item.checkState() == Qt.Checked:
                selected_indices.append(item.data(Qt.UserRole))
        return sd, ed, selected_indices