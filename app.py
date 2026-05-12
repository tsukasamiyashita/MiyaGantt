# tsukasamiyashita/miyagantt/MiyaGantt-46a1664b6d1737cb32f1dd17429ce06cca8dc678/app.py
import sys
import os
import calendar
import json
from datetime import datetime, timedelta

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QTableWidgetItem, 
                               QSplitter, QGraphicsView, QMessageBox, 
                               QLabel, QSpinBox, QComboBox, QHeaderView, QTableWidget,
                               QLineEdit, QAbstractItemView)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QBrush, QColor, QFont, QIcon, QPainter

from dialogs import SettingsDialog, ColorGridDialog, SummaryDialog, HelpDialog
from gantt_items import HeaderScene, ChartScene
from task_table import TaskTable, HeadcountDelegate, ModeDelegate, EfficiencyDelegate
from chart_renderer import ChartRenderer

from history_manager import HistoryManagerMixin
from file_manager import FileManagerMixin
from print_manager import PrintManagerMixin
from task_manager import TaskManagerMixin

class GanttApp(QMainWindow, HistoryManagerMixin, FileManagerMixin, PrintManagerMixin, TaskManagerMixin):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MiyaGantt - Professional Gantt Chart")
        self.resize(1380, 850)
        self.project_title = ""
        self.tasks = []
        self.day_width = 40
        self.row_height = 40
        self.header_height = 70
        self.min_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        self.display_unit = 1
        self.display_count = 6
        self.zoom_unit = 1
        self.zoom_count = 1
        self.update_display_days()
        self.month_label_items = []
        self.visible_tasks_info = []
        self.custom_holidays = {}
        self.summary_visible = True
        self.last_summary_base_key = None
        self.max_date = self.min_date + timedelta(days=180)
        self.last_path = ""
        
        self.renderer = ChartRenderer(self)
        
        self.setWindowIcon(QIcon(self.get_icon_path()))
        
        self.snap_timer = QTimer()
        self.snap_timer.setSingleShot(True)
        self.snap_timer.timeout.connect(self.snap_horizontal_scroll)
        
        self.init_history()
        self.init_ui()
        self.apply_styles()
        self.update_history_buttons()
        
        # 起動時点の状態を保存済みとして記録
        self.saved_snapshot = self.get_current_data_snapshot()

    def get_current_data_snapshot(self):
        """現在のデータの状態をJSON文字列として取得（変更検知用）"""
        data = {
            'project_title': getattr(self, 'project_title', ''),
            'min_date': getattr(self, 'min_date', datetime.now()).strftime("%Y-%m-%d"),
            'max_date': getattr(self, 'max_date', datetime.now()).strftime("%Y-%m-%d"),
            'tasks': getattr(self, 'tasks', [])
        }
        return json.dumps(data, ensure_ascii=False, sort_keys=True)

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #f0f0f0; }
            QWidget { background-color: #f0f0f0; color: #333333; }
            QTableWidget { background-color: #ffffff; gridline-color: #e0e0e0; border: 1px solid #cccccc; color: #333333; outline: 0; }
            QTableWidget::item:selected { background-color: transparent; color: #333333; }
            QHeaderView::section { background-color: #e8e8e8; color: #333333; border: 1px solid #cccccc; padding: 4px; font-weight: bold; }
            QPushButton { background-color: #ffffff; border: 1px solid #cccccc; padding: 6px 12px; border-radius: 4px; }
            QPushButton:hover:!disabled { background-color: #e8e8e8; }
            QPushButton:disabled { background-color: #f5f5f5; color: #aaaaaa; border: 1px solid #e0e0e0; }
        """)

    def init_ui(self):
        mw = QWidget()
        self.setCentralWidget(mw)
        ml = QVBoxLayout(mw)
        
        title_layout = QHBoxLayout()
        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("プロジェクトのタイトルを入力...")
        self.title_edit.setFont(QFont("Segoe UI", 16, QFont.Bold))
        self.title_edit.setStyleSheet("""
            QLineEdit { border: 1px solid transparent; background: transparent; padding: 4px; border-radius: 4px; color: #333333; }
            QLineEdit:hover { border: 1px solid #cccccc; }
            QLineEdit:focus { border: 1px solid #0078d4; background: #ffffff; }
        """)
        self.title_edit.setText(self.project_title)
        self.title_edit.textChanged.connect(self.on_title_changed)
        self.title_edit.editingFinished.connect(self.save_state_if_changed)
        title_layout.addWidget(self.title_edit)
        ml.addLayout(title_layout)
        
        tl = QHBoxLayout()
        
        self.btn_add = QPushButton("追加")
        self.btn_group = QPushButton("グループ")
        self.btn_up = QPushButton("↑")
        self.btn_down = QPushButton("↓")
        self.btn_del = QPushButton("削除")
        self.btn_undo = QPushButton("↶ 戻る")
        self.btn_redo = QPushButton("↷ 進む")
        self.btn_undo.setToolTip("元に戻す (Ctrl+Z)")
        self.btn_redo.setToolTip("やり直し (Ctrl+Y)")
        
        self.btn_add.clicked.connect(self.add_task)
        self.btn_group.clicked.connect(self.add_group)
        self.btn_up.clicked.connect(self.move_row_up)
        self.btn_down.clicked.connect(self.move_row_down)
        self.btn_del.clicked.connect(self.delete_task)
        self.btn_undo.clicked.connect(self.undo)
        self.btn_redo.clicked.connect(self.redo)
        self.btn_undo.setShortcut("Ctrl+Z")
        self.btn_redo.setShortcut("Ctrl+Y")
        
        tl.addWidget(self.btn_add)
        tl.addWidget(self.btn_group)
        tl.addWidget(self.btn_up)
        tl.addWidget(self.btn_down)
        tl.addWidget(self.btn_del)
        tl.addWidget(self.btn_undo)
        tl.addWidget(self.btn_redo)
        
        tl.addWidget(QLabel(" ｜ デフォルト設定:"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["人員モード (手動)", "案件モード (自動)", "メモモード", "見出し"])
        tl.addWidget(self.mode_combo)
        
        tl.addWidget(QLabel(" ｜ 案件表示:"))
        self.auto_disp_combo = QComboBox()
        self.auto_disp_combo.addItems(["日別/累積", "日別のみ", "累積のみ", "非表示"])
        self.auto_disp_combo.currentIndexChanged.connect(self.on_auto_disp_changed)
        tl.addWidget(self.auto_disp_combo)
        self.auto_disp_mode = 0
        
        tl.addStretch()
        
        tl.addWidget(QLabel("1画面の表示枠:"))
        self.zoom_unit_combo = QComboBox()
        self.zoom_unit_combo.addItems(["週間", "月間", "年間"])
        self.zoom_unit_combo.setCurrentIndex(self.zoom_unit)
        self.zoom_unit_combo.currentIndexChanged.connect(self.on_zoom_changed)
        tl.addWidget(self.zoom_unit_combo)
        
        tl.addWidget(QLabel("枠に収める数:"))
        self.zoom_count_spin = QSpinBox()
        self.zoom_count_spin.setRange(1, 100)
        self.zoom_count_spin.setValue(self.zoom_count)
        self.zoom_count_spin.valueChanged.connect(self.on_zoom_changed)
        tl.addWidget(self.zoom_count_spin)
        
        self.btn_new = QPushButton("新規作成")
        self.btn_new.setShortcut("Ctrl+N")
        self.btn_new.setToolTip("新規作成 (Ctrl+N)")
        self.btn_load = QPushButton("読込")
        self.btn_save = QPushButton("上書き保存")
        self.btn_save_as = QPushButton("名前を付けて保存")
        self.btn_save.setShortcut("Ctrl+S")
        self.btn_save_as.setShortcut("Ctrl+Shift+S")
        self.btn_save.setToolTip("上書き保存 (Ctrl+S)")
        self.btn_save_as.setToolTip("名前を付けて保存 (Ctrl+Shift+S)")
        self.btn_settings = QPushButton("⚙ 編集期間")
        self.btn_summary = QPushButton("📊 集計")
        self.btn_print = QPushButton("🖨 印刷")
        self.btn_save_config = QPushButton("💾 設定保存")
        self.btn_help = QPushButton("❓ ヘルプ")
        self.btn_new.clicked.connect(self.new_project)
        self.btn_load.clicked.connect(self.load_data)
        self.btn_save.clicked.connect(self.save_data)
        self.btn_save_as.clicked.connect(self.save_data_as)
        self.btn_settings.clicked.connect(self.open_settings)
        self.btn_summary.clicked.connect(self.open_summary)
        self.btn_print.clicked.connect(self.print_gantt)
        self.btn_save_config.clicked.connect(self.save_app_config)
        self.btn_help.clicked.connect(self.open_help)
        
        tl.addWidget(self.btn_new)
        tl.addWidget(self.btn_load)
        tl.addWidget(self.btn_save)
        tl.addWidget(self.btn_save_as)
        tl.addWidget(self.btn_settings)
        tl.addWidget(self.btn_summary)
        tl.addWidget(self.btn_print)
        tl.addWidget(self.btn_save_config)
        tl.addWidget(self.btn_help)
        tl.addStretch()
        
        ml.addLayout(tl)
        
        tl2 = QHBoxLayout()
        tl2.addWidget(QLabel("カレンダー移動:"))
        
        self.btn_prev_y = QPushButton("≪年")
        self.btn_prev_m = QPushButton("＜月")
        self.btn_prev_w = QPushButton("＜週")
        self.btn_prev_d = QPushButton("＜日")
        self.btn_today = QPushButton("📅 今日")
        self.btn_next_d = QPushButton("日＞")
        self.btn_next_w = QPushButton("週＞")
        self.btn_next_m = QPushButton("月＞")
        self.btn_next_y = QPushButton("年≫")
        
        self.btn_prev_y.clicked.connect(lambda: self.scroll_by_unit('year', -1))
        self.btn_prev_m.clicked.connect(lambda: self.scroll_by_unit('month', -1))
        self.btn_prev_w.clicked.connect(lambda: self.scroll_by_unit('week', -1))
        self.btn_prev_d.clicked.connect(lambda: self.scroll_by_unit('day', -1))
        self.btn_today.clicked.connect(self.scroll_to_today)
        self.btn_next_d.clicked.connect(lambda: self.scroll_by_unit('day', 1))
        self.btn_next_w.clicked.connect(lambda: self.scroll_by_unit('week', 1))
        self.btn_next_m.clicked.connect(lambda: self.scroll_by_unit('month', 1))
        self.btn_next_y.clicked.connect(lambda: self.scroll_by_unit('year', 1))
        
        nav_btns = [self.btn_today, 
                    self.btn_prev_d, self.btn_next_d,
                    self.btn_prev_w, self.btn_next_w,
                    self.btn_prev_m, self.btn_next_m,
                    self.btn_prev_y, self.btn_next_y]
        
        for b in nav_btns:
            if b != self.btn_today:
                b.setFixedWidth(45)
            tl2.addWidget(b)
            
        tl2.addStretch()
        ml.addLayout(tl2)
        
        self.splitter = QSplitter(Qt.Horizontal)
        ml.addWidget(self.splitter)
        
        self.left_container = QWidget()
        self.left_layout = QVBoxLayout(self.left_container)
        self.left_layout.setContentsMargins(0, 0, 0, 0)
        self.left_layout.setSpacing(2)
        
        self.col_toggle_layout = QHBoxLayout()
        self.col_toggle_layout.setContentsMargins(2, 2, 2, 2)
        self.col_toggle_layout.setSpacing(2)
        self.col_actions = {}
        col_info = [
            (0, "マーク"), (1, "開閉"), (3, "モード"), (4, "人数"), 
            (5, "補正"), (6, "期間"), (7, "色"), (8, "合計")
        ]
        for idx, name in col_info:
            btn = QPushButton(name)
            btn.setCheckable(True)
            btn.setChecked(True)
            btn.setFixedHeight(24)
            btn.setStyleSheet("""
                QPushButton { background-color: #f8f8f8; border: 1px solid #ddd; border-radius: 4px; color: #666; font-size: 10px; padding: 0 5px; }
                QPushButton:checked { background-color: #e1f0ff; border: 1px solid #0000ff; color: #0000ff; font-weight: bold; }
                QPushButton:hover { background-color: #eeeeee; }
            """)
            btn.clicked.connect(lambda checked, i=idx: self.toggle_column_visibility(i, checked))
            self.col_actions[idx] = btn
            self.col_toggle_layout.addWidget(btn)
        self.col_toggle_layout.addStretch()
        self.left_layout.addLayout(self.col_toggle_layout)

        self.table = TaskTable(0, 8)
        self.table.setHorizontalHeaderLabels(["", "", "タスク名", "モード", "人数", "工数補正", "期間/開始日", "色"])
        self.table.setColumnWidth(0, 25)
        self.table.setColumnWidth(1, 35)
        self.table.setColumnWidth(2, 170)
        self.table.setColumnWidth(3, 45)
        self.table.setColumnWidth(4, 45)
        self.table.setColumnWidth(5, 65)
        self.table.setColumnWidth(6, 145)
        self.table.setColumnWidth(7, 40)
        
        self.table.horizontalHeader().setFixedHeight(self.header_height)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.table.verticalHeader().setDefaultSectionSize(self.row_height)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.table.setItemDelegateForColumn(5, EfficiencyDelegate(self.table))
        self.table.setItemDelegateForColumn(4, HeadcountDelegate(self.table))
        self.table.setItemDelegateForColumn(3, ModeDelegate(self.table))
        
        self.table.itemChanged.connect(self.on_table_item_changed)
        self.table.cellClicked.connect(self.on_table_cell_clicked)
        self.table.cellDoubleClicked.connect(self.on_table_cell_double_clicked)
        self.table.currentCellChanged.connect(self.update_selection_mark)
        
        self.left_layout.addWidget(self.table)
        self.splitter.addWidget(self.left_container)
        
        rc = QWidget()
        rcl = QVBoxLayout(rc); rcl.setContentsMargins(0,0,0,0); rcl.setSpacing(0)
        
        spacer = QWidget()
        spacer.setFixedHeight(28)
        rcl.addWidget(spacer)
        
        self.hs = HeaderScene(self)
        self.cs = ChartScene(self)
        self.hv = QGraphicsView(self.hs)
        self.chart_view = QGraphicsView(self.cs)
        
        for v in [self.hv, self.chart_view]:
            v.setRenderHint(QPainter.Antialiasing)
            v.setAlignment(Qt.AlignLeft | Qt.AlignTop)
            v.setBackgroundBrush(QBrush(Qt.white))
            v.setStyleSheet("QGraphicsView { border: none; border-left: 1px solid #cccccc; }")
            
        self.hv.setFixedHeight(self.header_height)
        self.hv.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.hv.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.chart_view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        
        rcl.addWidget(self.hv)
        rcl.addWidget(self.chart_view)
        self.splitter.addWidget(rc)
        self.splitter.setSizes([450, 930])
        
        self.load_app_config()
        
        self.chart_view.horizontalScrollBar().valueChanged.connect(self.on_horizontal_scroll)
        self.chart_view.horizontalScrollBar().sliderReleased.connect(self.snap_horizontal_scroll)
        self.table.verticalScrollBar().valueChanged.connect(self.chart_view.verticalScrollBar().setValue)
        self.chart_view.verticalScrollBar().valueChanged.connect(self.table.verticalScrollBar().setValue)

    def on_title_changed(self, text):
        self.project_title = text
        title_str = f"{text} - MiyaGantt" if text else "MiyaGantt - Professional Gantt Chart"
        self.setWindowTitle(title_str)

    def on_horizontal_scroll(self, v):
        try:
            self.hv.horizontalScrollBar().setValue(v)
            self.update_month_labels_pos()
            
            if not self.chart_view.horizontalScrollBar().isSliderDown():
                self.snap_timer.start(300)
            
            if self.day_width <= 0: return
            days_scrolled = v / self.day_width
            visible_start = self.min_date + timedelta(days=days_scrolled)
            threshold_date = self.get_threshold_date(visible_start)
            
            if self.display_unit == 0:
                base_key = (threshold_date - timedelta(days=threshold_date.weekday())).strftime("%Y-%W")
            elif self.display_unit == 1:
                base_key = threshold_date.strftime("%Y-%m")
            else:
                base_key = threshold_date.strftime("%Y")
                
            if base_key != self.last_summary_base_key:
                self.last_summary_base_key = base_key
                self.sync_summary_to_scroll(threshold_date)
        except RuntimeError:
            pass

    def snap_horizontal_scroll(self):
        if self.day_width > 0:
            v = self.chart_view.horizontalScrollBar().value()
            snapped = round(v / self.day_width) * self.day_width
            self.chart_view.horizontalScrollBar().setValue(int(snapped))

    def update_month_labels_pos(self):
        try:
            view_left = self.hv.mapToScene(0, 0).x()
            for start_x, end_x, item in self.month_label_items:
                if item.scene():
                    tw = item.boundingRect().width()
                    new_x = max(start_x + 5, min(view_left + 5, end_x - tw - 5))
                    item.setPos(new_x, item.pos().y())
        except RuntimeError:
            pass

    def calculate_day_width(self):
        if self.zoom_unit == 0:
            days = self.zoom_count * 7
        elif self.zoom_unit == 1:
            days = self.zoom_count * 30.416
        else:
            days = self.zoom_count * 365.25
            
        view_width = self.chart_view.viewport().width() if hasattr(self, 'chart_view') else 1000
        if view_width < 100:
            view_width = 1000
            
        self.day_width = max(1.0, view_width / max(1.0, days))
        
        if hasattr(self, 'chart_view'):
            self.chart_view.horizontalScrollBar().setSingleStep(max(1, int(self.day_width)))

    def on_zoom_changed(self, *_):
        self.zoom_unit = self.zoom_unit_combo.currentIndex()
        self.zoom_count = self.zoom_count_spin.value()
        self.display_unit = self.zoom_unit
        self.update_display_days()
        self.calculate_day_width()
        self.update_ui()

    def on_auto_disp_changed(self, idx):
        self.auto_disp_mode = idx
        self.draw_chart()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'chart_view'):
            self.calculate_day_width()
            self.update_ui()

    def update_display_range(self):
        self.zoom_unit_combo.blockSignals(True)
        self.zoom_unit_combo.setCurrentIndex(self.display_unit)
        self.zoom_unit_combo.blockSignals(False)
        self.zoom_unit = self.display_unit
        
        self.update_display_days()
        self.calculate_day_width()
        self.update_ui()

    def update_display_days(self):
        if hasattr(self, 'max_date') and self.max_date:
            self.display_days = max(1, (self.max_date - self.min_date).days + 1)
            return

        if self.display_unit == 0:
            self.display_days = self.display_count * 7
        elif self.display_unit == 1:
            m = self.min_date.month - 1 + self.display_count
            y = self.min_date.year + m // 12
            m = m % 12 + 1
            last_day = calendar.monthrange(y, m)[1]
            d = min(self.min_date.day, last_day)
            end_date = datetime(y, m, d)
            self.display_days = max(1, (end_date - self.min_date).days)
        elif self.display_unit == 2:
            count_months = self.display_count * 12
            m = self.min_date.month - 1 + count_months
            y = self.min_date.year + m // 12
            m = m % 12 + 1
            last_day = calendar.monthrange(y, m)[1]
            d = min(self.min_date.day, last_day)
            end_date = datetime(y, m, d)
            self.display_days = max(1, (end_date - self.min_date).days)
        
        self.max_date = self.min_date + timedelta(days=self.display_days - 1)

    def update_selection_mark(self, *args):
        self.table.blockSignals(True)
        curr = self.table.currentRow()
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it:
                it.setText("●" if r == curr else "")
                it.setTextAlignment(Qt.AlignCenter)
                it.setForeground(QColor(0, 120, 212))
        self.table.blockSignals(False)

    def get_threshold_date(self, visible_start):
        if self.display_unit == 0:
            return visible_start + timedelta(days=2)
        elif self.display_unit == 1:
            return visible_start + timedelta(days=7)
        else:
            return visible_start + timedelta(days=30)

    def get_summary_headers(self, base_date=None, count=None):
        if base_date is None: base_date = self.min_date
        base_date = base_date.replace(hour=0, minute=0, second=0, microsecond=0)
        if count is None:
            if self.zoom_unit == 0:
                visible_days = self.zoom_count * 7
            elif self.zoom_unit == 1:
                visible_days = self.zoom_count * 30.416
            else:
                visible_days = self.zoom_count * 365.25
                
            if self.display_unit == 0:
                count = max(1, round(visible_days / 7))
            elif self.display_unit == 1:
                count = max(1, round(visible_days / 30.416))
            else:
                count = max(1, round(visible_days / 365.25))
        
        headers = []
        curr = base_date
        unit_type = ['week', 'month', 'year'][self.display_unit]
        
        if unit_type == 'week':
            curr = curr - timedelta(days=curr.weekday())
            for _ in range(count):
                end_d = curr + timedelta(days=6)
                label = f"{curr.strftime('%m/%d')}~{end_d.strftime('%m/%d')} 工数"
                headers.append((curr, end_d, label))
                curr += timedelta(days=7)
        elif unit_type == 'month':
            curr = curr.replace(day=1)
            for _ in range(count):
                last_day = calendar.monthrange(curr.year, curr.month)[1]
                headers.append((curr, curr.replace(day=last_day), f"{curr.strftime('%Y/%m')} 工数"))
                m = curr.month + 1
                y = curr.year
                if m > 12: m = 1; y += 1
                curr = datetime(y, m, 1)
        elif unit_type == 'year':
            curr = curr.replace(month=1, day=1)
            for _ in range(count):
                headers.append((curr, curr.replace(month=12, day=31), f"{curr.strftime('%Y年')} 工数"))
                curr = curr.replace(year=curr.year + 1)
        return headers

    def sync_summary_to_scroll(self, base_date):
        if not hasattr(self, 'table'): return
        headers = self.get_summary_headers(base_date)
        
        self.table.blockSignals(True)
        labels = ["", "", "タスク名", "モード", "人数", "工数補正", "期間/開始日", "色"] + [h[2] for h in headers]
        if self.table.columnCount() != len(labels):
            self.table.setColumnCount(len(labels))
        self.table.setHorizontalHeaderLabels(labels)
        
        for r, info in enumerate(self.visible_tasks_info):
            t = info['task']
            for i, (h_start, h_end, _) in enumerate(headers):
                col_idx = 8 + i
                item_s = self.table.item(r, col_idx)
                if not item_s:
                    item_s = QTableWidgetItem()
                    self.table.setItem(r, col_idx, item_s)
                
                is_auto = not t.get('is_group') and t.get('mode') == 'auto'
                is_memo = not t.get('is_group') and t.get('mode') in ['memo', 'heading']
                
                if is_auto:
                    item_s.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                    item_s.setText(f"{t.get('workload', 1.0):.1f}工数")
                elif is_memo:
                    item_s.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    item_s.setText("-")
                else:
                    item_s.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    day_map = self.get_task_workload_in_range(t, info['index'], h_start, h_end)
                    item_s.setText(self.format_summary_workload(day_map))
                
                item_s.setTextAlignment(Qt.AlignCenter)
                
                if t.get('is_group'):
                    item_s.setBackground(QColor(242, 242, 242))
                else:
                    if is_auto:
                        item_s.setBackground(QColor(245, 250, 255))
                    elif t.get('mode') == 'memo':
                        item_s.setBackground(QColor(252, 252, 252))
                    elif t.get('mode') == 'heading':
                        item_s.setBackground(QColor(245, 245, 245))
                    else:
                        item_s.setBackground(QColor(255, 255, 255))
                
                if self.summary_visible:
                    self.table.setColumnWidth(col_idx, 100)
                self.table.setColumnHidden(col_idx, not self.summary_visible)
                
        self.table.blockSignals(False)

    def get_task_workload_in_range(self, t, start_idx, timeline_start=None, timeline_end=None):
        if timeline_start is None: timeline_start = self.min_date
        if timeline_end is None: timeline_end = self.min_date + timedelta(days=self.display_days - 1)
        
        timeline_start = timeline_start.replace(hour=0, minute=0, second=0, microsecond=0)
        timeline_end = timeline_end.replace(hour=0, minute=0, second=0, microsecond=0)
        
        if t.get('is_group'):
            manual_total = 0.0
            auto_total = 0.0
            for i in range(start_idx + 1, len(self.tasks)):
                if self.tasks[i].get('is_group'): break
                child = self.tasks[i]
                mode = child.get('mode', 'manual')
                if mode == 'manual':
                    hc = float(child.get('headcount', 1.0)) * float(child.get('efficiency', 1.0))
                    t_color = child.get('color', '#808080')
                    t_color_name = QColor(t_color).name()
                    for p in child.get('periods', []):
                        if not p.get('start_date') or not p.get('end_date'): continue
                        try:
                            psd = datetime.strptime(p['start_date'], "%Y-%m-%d")
                            ped = datetime.strptime(p['end_date'], "%Y-%m-%d")
                            calc_start = max(psd, timeline_start)
                            calc_end = min(ped, timeline_end)
                            if calc_start <= calc_end:
                                p_color = p.get('color')
                                if p_color and QColor(p_color).name() != t_color_name:
                                    continue
                                overlap = (calc_end - calc_start).days + 1
                                manual_total += (overlap * hc)
                        except ValueError:
                            pass
                elif mode == 'auto':
                    allocs = child.get('daily_allocations', {})
                    if allocs:
                        curr = timeline_start
                        while curr <= timeline_end:
                            d_str = curr.strftime("%Y-%m-%d")
                            auto_total += allocs.get(d_str, 0.0)
                            curr += timedelta(days=1)
            return {'is_group_summary': True, 'manual': manual_total, 'auto': auto_total}
        
        day_map = {}
        if t.get('mode') in ['auto', 'memo', 'heading']:
            return day_map
            
        t_color = t.get('color', '#808080')
        t_color_name = QColor(t_color).name()
        hc = float(t.get('headcount', 1.0)) * float(t.get('efficiency', 1.0))
            
        for p in t.get('periods', []):
            if not p.get('start_date') or not p.get('end_date'): continue
            try:
                psd = datetime.strptime(p['start_date'], "%Y-%m-%d")
                ped = datetime.strptime(p['end_date'], "%Y-%m-%d")
                calc_start = max(psd, timeline_start)
                calc_end = min(ped, timeline_end)
                if calc_start <= calc_end:
                    p_color = p.get('color')
                    if p_color and QColor(p_color).name() != t_color_name:
                        continue
                    overlap = (calc_end - calc_start).days + 1
                    day_map[t_color] = day_map.get(t_color, 0) + (overlap * hc)
            except ValueError:
                continue
        return day_map

    def sync_table_from_tasks(self):
        self.table.blockSignals(True)
        headers = self.get_summary_headers()
        for r, info in enumerate(self.visible_tasks_info):
            t = info['task']
            if t.get('is_group'):
                for i, (h_start, h_end, _) in enumerate(headers):
                    col_idx = 8 + i
                    item_s = self.table.item(r, col_idx)
                    if item_s:
                        day_map = self.get_task_workload_in_range(t, info['index'], h_start, h_end)
                        item_s.setText(self.format_summary_workload(day_map))
                continue
            
            period_item = self.table.item(r, 6)
            if period_item:
                if t.get('mode') == 'auto':
                    period_item.setText(t.get('auto_start_date', ''))
                else:
                    periods = t.get('periods', [])
                    p_strs = []
                    for p in periods:
                        if not p.get('start_date') or not p.get('end_date'): continue
                        s = p['start_date'].replace('-', '/')
                        e = p['end_date'].replace('-', '/')
                        p_strs.append(f"{s}-{e}")
                    period_item.setText(", ".join(p_strs))
            
            for i, (h_start, h_end, _) in enumerate(headers):
                col_idx = 8 + i
                item_s = self.table.item(r, col_idx)
                if item_s:
                    is_auto = not t.get('is_group') and t.get('mode') == 'auto'
                    is_memo = not t.get('is_group') and t.get('mode') in ['memo', 'heading']
                    
                    if is_auto:
                        item_s.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                        item_s.setText(f"{t.get('workload', 1.0):.1f}工数")
                    elif is_memo:
                        item_s.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                        item_s.setText("-")
                    else:
                        item_s.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                        day_map = self.get_task_workload_in_range(t, info['index'], h_start, h_end)
                        item_s.setText(self.format_summary_workload(day_map))
        self.table.blockSignals(False)

    def update_ui(self, refresh_chart=True):
        self.visible_tasks_info = self.get_visible_tasks_info()
        self.table.blockSignals(True)
        
        scroll_val = self.chart_view.horizontalScrollBar().value()
        days_scrolled = scroll_val / self.day_width if self.day_width > 0 else 0
        visible_start = self.min_date + timedelta(days=days_scrolled)
        threshold_date = self.get_threshold_date(visible_start)
        
        headers = self.get_summary_headers(threshold_date)
        base_col_count = 8
        total_cols = base_col_count + len(headers)
        self.table.setColumnCount(total_cols)
        
        labels = ["", "", "タスク名", "モード", "人数", "工数補正", "期間/開始日", "色"] + [h[2] for h in headers]
        self.table.setHorizontalHeaderLabels(labels)
        
        new_rows = len(self.visible_tasks_info)
        if self.table.rowCount() < new_rows:
            self.table.setRowCount(new_rows)
            
        for r, info in enumerate(self.visible_tasks_info):
            t = info['task']
            indent = "    " * info['indent']
            is_group = t.get('is_group', False)
            is_auto = not is_group and t.get('mode') == 'auto'
            is_memo = not is_group and t.get('mode') == 'memo'
            is_heading = not is_group and t.get('mode') == 'heading'
            
            if is_group:
                bg_row = QColor(235, 235, 235)
            elif is_auto:
                bg_row = QColor(235, 250, 235)
            elif is_memo:
                bg_row = QColor(255, 250, 230)
            elif is_heading:
                bg_row = QColor(225, 240, 255)
            else:
                bg_row = QColor(255, 255, 255)
            
            for c in range(total_cols):
                if self.table.item(r, c) is None:
                    self.table.setItem(r, c, QTableWidgetItem())
            
            mark_item = self.table.item(r, 0)
            mark_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            mark_item.setBackground(bg_row)

            toggle_item = self.table.item(r, 1)
            toggle_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            toggle_item.setForeground(QColor(51, 51, 51))
            f = toggle_item.font(); f.setBold(False); toggle_item.setFont(f)
            toggle_item.setBackground(bg_row)
            
            if is_group:
                toggle_item.setText("▼" if not t.get('collapsed') else "▶")
                toggle_item.setTextAlignment(Qt.AlignCenter)
                f = toggle_item.font(); f.setBold(True); toggle_item.setFont(f)
            else:
                toggle_item.setText("")

            item_name = self.table.item(r, 2)
            item_name.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_name.setForeground(QColor(51, 51, 51))
            f = item_name.font(); f.setBold(False); item_name.setFont(f)
            item_name.setBackground(bg_row)
            item_name.setText(indent + t.get('name', ''))
            
            item_mode = self.table.item(r, 3)
            item_mode.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_mode.setForeground(QColor(51, 51, 51))
            
            item_hc = self.table.item(r, 4)
            item_hc.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_hc.setForeground(QColor(51, 51, 51))
            item_hc.setBackground(bg_row)
            
            item_eff = self.table.item(r, 5)
            item_eff.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_eff.setForeground(QColor(51, 51, 51))
            item_eff.setBackground(bg_row)
            
            item_period = self.table.item(r, 6)
            item_period.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_period.setForeground(QColor(51, 51, 51))
            item_period.setBackground(bg_row)
            
            item_color = self.table.item(r, 7)
            item_color.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_color.setForeground(QColor(51, 51, 51))
            
            if is_group:
                f = item_name.font(); f.setBold(True); item_name.setFont(f)
                
                item_mode.setText("")
                item_mode.setBackground(bg_row)
                
                item_hc.setText(f"{int(t.get('headcount', 1.0))}")
                item_hc.setTextAlignment(Qt.AlignCenter)
                
                item_eff.setText("")
                item_eff.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                
                item_period.setText("")
                item_period.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                item_color.setText("")
                item_color.setBackground(QColor(200, 200, 200))
                item_color.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            else:
                f_m = item_mode.font()
                f_m.setBold(True)
                item_mode.setFont(f_m)
                
                if is_auto:
                    item_mode.setText("⚡ 案件")
                    item_mode.setForeground(QColor(255, 255, 255))
                    item_mode.setBackground(QColor("#323130"))

                    hc = t.get('headcount', 0.0)
                    item_hc.setText(f"{int(hc)}" if hc > 0 else "制限なし")
                    item_hc.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                    item_eff.setText("-")
                    item_eff.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    item_period.setText(t.get('auto_start_date', ''))
                elif is_memo or is_heading:
                    if is_heading:
                        item_mode.setText("📌 見出し")
                        item_mode.setForeground(QColor(255, 255, 255))
                        item_mode.setBackground(QColor("#4169e1"))
                        f_name = item_name.font()
                        f_name.setBold(True)
                        item_name.setFont(f_name)
                    else:
                        item_mode.setText("📝 メモ")
                        item_mode.setForeground(QColor(51, 51, 51))
                        item_mode.setBackground(QColor("#c0c0c0"))

                    item_hc.setText("-")
                    item_hc.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    item_eff.setText("-")
                    item_eff.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    
                    periods = t.get('periods', [])
                    p_strs = []
                    for p in periods:
                        if not p.get('start_date') or not p.get('end_date'): continue
                        s = p['start_date'].replace('-', '/')
                        e = p['end_date'].replace('-', '/')
                        p_strs.append(f"{s}-{e}")
                    item_period.setText(", ".join(p_strs))
                else:
                    item_mode.setText("👤 人員")
                    item_mode.setForeground(QColor(255, 255, 255))
                    item_mode.setBackground(QColor("#808080"))

                    item_hc.setText(f"{int(t.get('headcount', 1.0))}")
                    eff = t.get('efficiency', 1.0)
                    item_eff.setText(f"{int(eff * 100)}%")
                    
                    periods = t.get('periods', [])
                    p_strs = []
                    for p in periods:
                        if not p.get('start_date') or not p.get('end_date'): continue
                        s = p['start_date'].replace('-', '/')
                        e = p['end_date'].replace('-', '/')
                        p_strs.append(f"{s}-{e}")
                    item_period.setText(", ".join(p_strs))
                    
                item_mode.setTextAlignment(Qt.AlignCenter)
                item_hc.setTextAlignment(Qt.AlignCenter)
                item_eff.setTextAlignment(Qt.AlignCenter)
                
                item_color.setText("")
                item_color.setBackground(QColor(t.get('color', '#808080')))
                item_color.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)

            for i, (h_start, h_end, _) in enumerate(headers):
                col_idx = 8 + i
                item_s = self.table.item(r, col_idx)
                
                if not is_group and is_auto:
                    item_s.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                    item_s.setText(f"{t.get('workload', 1.0):.1f}工数")
                elif not is_group and (is_memo or is_heading):
                    item_s.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    item_s.setText("-")
                else:
                    item_s.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    day_map = self.get_task_workload_in_range(t, info['index'], h_start, h_end)
                    item_s.setText(self.format_summary_workload(day_map))
                
                item_s.setForeground(QColor(51, 51, 51))
                item_s.setTextAlignment(Qt.AlignCenter)
                item_s.setBackground(bg_row)
                
        if self.table.rowCount() > new_rows:
            self.table.setRowCount(new_rows)
            
        for i in range(base_col_count, total_cols):
            self.table.setColumnHidden(i, not self.summary_visible)
            if self.summary_visible:
                self.table.setColumnWidth(i, 100)

        self.update_selection_mark()
        self.table.blockSignals(False)
        if refresh_chart:
            self.draw_chart()

    def draw_chart(self):
        self.renderer.draw_chart()

    def on_table_item_changed(self, item):
        row = item.row()
        col = item.column()
        if row >= len(self.visible_tasks_info) or row < 0: return
        info = self.visible_tasks_info[row]
        t = info['task']
        
        if col == 2:
            t['name'] = item.text().strip()
        elif col == 3:
            if t.get('is_group'): return
            new_mode_ja = item.text().strip()
            new_mode_en = 'manual'
            if '案件' in new_mode_ja: new_mode_en = 'auto'
            elif 'メモ' in new_mode_ja: new_mode_en = 'memo'
            elif '見出し' in new_mode_ja: new_mode_en = 'heading'
            
            if t.get('mode') != new_mode_en:
                t['mode'] = new_mode_en
                
                if new_mode_en == 'auto':
                    new_color = '#323130'
                elif new_mode_en == 'memo':
                    new_color = '#c0c0c0'
                elif new_mode_en == 'heading':
                    new_color = '#4169e1'
                else:
                    new_color = '#808080'
                t['color'] = new_color
                
                if new_mode_en == 'auto':
                    if not t.get('auto_start_date') and t.get('periods'):
                        t['auto_start_date'] = t['periods'][0].get('start_date', '')
                    if not t.get('auto_start_date'):
                        scroll_val = self.chart_view.horizontalScrollBar().value()
                        days_scrolled = scroll_val / self.day_width if getattr(self, 'day_width', 0) > 0 else 0
                        visible_start = self.min_date + timedelta(days=days_scrolled)
                        t['auto_start_date'] = visible_start.strftime("%Y-%m-%d")
                    if 'workload' not in t:
                        t['workload'] = 1.0 
                    t['headcount'] = 0.0
                    if t.get('periods'):
                        for p in t['periods']:
                            p['color'] = new_color
                            if p.get('text') and ("⚠️ キャパオーバー" in p.get('text') or "⚠️ 進行不可" in p.get('text')):
                                p['text'] = ""
                elif new_mode_en in ['memo', 'heading']:
                    t['headcount'] = 0.0
                    if t.get('periods'):
                        for p in t['periods']:
                            p['color'] = new_color
                            if p.get('text') and ("⚠️ キャパオーバー" in p.get('text') or "⚠️ 進行不可" in p.get('text')):
                                p['text'] = ""
                else:
                    if t.get('headcount', 0.0) == 0.0:
                        t['headcount'] = 1.0
                    if t.get('periods'):
                        for p in t['periods']:
                            p['color'] = new_color
                            if p.get('text') and ("⚠️ キャパオーバー" in p.get('text') or "⚠️ 進行不可" in p.get('text')):
                                p['text'] = ""
                
                self.recalculate_auto_tasks()
        elif col == 4:
            if t.get('mode') == 'auto':
                try:
                    val = item.text().strip()
                    if val == "" or val == "制限なし":
                        t['headcount'] = 0.0
                    else:
                        t['headcount'] = float(val)
                except ValueError:
                    pass
                
                self.table.blockSignals(True)
                hc = t.get('headcount', 0.0)
                item.setText(f"{int(hc)}" if hc > 0 else "制限なし")
                self.table.blockSignals(False)
                self.recalculate_auto_tasks()
            elif t.get('mode') in ['memo', 'heading']:
                self.table.blockSignals(True)
                item.setText("-")
                self.table.blockSignals(False)
            else:
                try:
                    val = item.text().strip()
                    if val == "" or val == "制限なし":
                        val = "1"
                    t['headcount'] = max(1.0, float(val))
                except ValueError:
                    pass
                
                self.table.blockSignals(True)
                item.setText(f"{int(t.get('headcount', 1.0))}")
                self.table.blockSignals(False)
                self.recalculate_auto_tasks()
                
        elif col == 5:
            if t.get('is_group') or t.get('mode') not in ['manual']: return
            val_str = item.text().replace('%', '').strip()
            try:
                t['efficiency'] = float(val_str) / 100.0
                self.recalculate_auto_tasks()
            except ValueError:
                pass
                
        elif col == 6:
            if t.get('is_group'): return
            period_str = item.text()
            if t.get('mode') == 'auto':
                parsed_date = self.parse_date(period_str)
                if parsed_date:
                    t['auto_start_date'] = parsed_date
                    self.recalculate_auto_tasks()
                else:
                    QMessageBox.warning(self, "エラー", "開始日の形式が正しくありません。\n例: 04/01")
            else:
                parsed = self.get_periods_from_string(period_str)
                if parsed:
                    t['periods'] = parsed
                else:
                    QMessageBox.warning(self, "エラー", "期間の形式が正しくありません。\n例: 04/01-04/05")
        elif col >= 8:
            if not t.get('is_group') and t.get('mode') == 'auto':
                try:
                    val = float(item.text().replace('工数', '').strip())
                    t['workload'] = max(0.1, val)
                except ValueError:
                    pass
                self.recalculate_auto_tasks()
        
        self.update_ui()
        self.save_state_if_changed()

    def on_table_cell_clicked(self, row, col):
        if row >= len(self.visible_tasks_info): return
        self.update_selection_mark()
        if col == 1:
            info = self.visible_tasks_info[row]
            t = info['task']
            if t.get('is_group'):
                t['collapsed'] = not t.get('collapsed', False)
                self.update_ui()
                self.save_state_if_changed()

    def on_table_cell_double_clicked(self, row, col):
        if row >= len(self.visible_tasks_info): return
        info = self.visible_tasks_info[row]
        t = info['task']
        
        if col == 7:
            color_groups = self.get_color_groups()
            dlg = ColorGridDialog(color_groups, self)
            if dlg.exec():
                t['color'] = dlg.selected_color
                self.update_ui()
                self.save_state_if_changed()

    def scroll_to_today(self):
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        target_date = today
        
        if self.display_unit == 0:
            target_date = today - timedelta(days=today.weekday())
        elif self.display_unit == 1:
            target_date = today.replace(day=1)
        elif self.display_unit == 2:
            target_date = today.replace(month=1, day=1)
            
        v = (target_date - self.min_date).days * self.day_width
        self.chart_view.horizontalScrollBar().setValue(int(v))

    def _scroll_to_specific_date(self, target_date):
        if self.day_width <= 0: return
        v = (target_date - self.min_date).days * self.day_width
        self.chart_view.horizontalScrollBar().setValue(int(v))

    def scroll_by_unit(self, unit, direction):
        if self.day_width <= 0: return
        
        v = self.chart_view.horizontalScrollBar().value()
        days_offset = round(v / self.day_width)
        current_date = self.min_date + timedelta(days=days_offset)
        
        if unit == 'day':
            new_date = current_date + timedelta(days=direction)
        elif unit == 'week':
            monday = current_date - timedelta(days=current_date.weekday())
            if direction > 0:
                new_date = monday + timedelta(days=7)
            else:
                new_date = monday - timedelta(days=7) if current_date == monday else monday
        elif unit == 'month':
            first_day = current_date.replace(day=1)
            if direction > 0:
                m = first_day.month % 12 + 1
                y = first_day.year + (1 if first_day.month == 12 else 0)
                new_date = datetime(y, m, 1)
            else:
                if current_date == first_day:
                    m = (first_day.month - 2) % 12 + 1
                    y = first_day.year - (1 if first_day.month == 1 else 0)
                    new_date = datetime(y, m, 1)
                else:
                    new_date = first_day
        elif unit == 'year':
            jan_first = current_date.replace(month=1, day=1)
            if direction > 0:
                new_date = jan_first.replace(year=jan_first.year + 1)
            else:
                if current_date == jan_first:
                    new_date = jan_first.replace(year=jan_first.year - 1)
                else:
                    new_date = jan_first
        
        new_v = (new_date - self.min_date).days * self.day_width
        self.chart_view.horizontalScrollBar().setValue(int(new_v))

    def open_settings(self):
        dlg = SettingsDialog(self)
        if dlg.exec():
            qs = dlg.start_date_edit.date()
            qe = dlg.end_date_edit.date()
            self.min_date = datetime(qs.year(), qs.month(), qs.day())
            self.max_date = datetime(qe.year(), qe.month(), qe.day())
            
            if self.max_date < self.min_date:
                self.max_date = self.min_date
            
            self.update_display_range()

    def open_summary(self):
        dlg = SummaryDialog(self, self.tasks, self)
        dlg.exec()

    def open_help(self):
        dlg = HelpDialog(self)
        dlg.exec()

    def get_icon_path(self):
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, 'icon.ico')
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icon.ico')

    def get_color_groups(self):
        return [
            ("青・水色系", [
                ("青", "#0000ff"), ("水色", "#00bcf2"), ("紺", "#002050"), 
                ("空色", "#87ceeb"), ("ロイヤルブルー", "#4169e1"), ("ネイビー", "#000080")
            ]),
            ("緑・ライム系", [
                ("緑", "#008000"), ("ライム", "#32cd32"), ("深緑", "#004b1c"),
                ("ミント", "#98ffed"), ("フォレストグリーン", "#228b22"), ("シーグリーン", "#2e8b57")
            ]),
            ("赤・桃系", [
                ("赤", "#d13438"), ("ワイン", "#a4262c"), ("ピンク", "#e67a91"),
                ("サーモン", "#fa8072"), ("マゼンタ", "#ff00ff"), ("ホットピンク", "#ff69b4")
            ]),
            ("橙・黄系", [
                ("オレンジ", "#ff8c00"), ("黄色", "#fff100"), ("ゴールド", "#ffd700"),
                ("コーラル", "#ff7f50"), ("アンバー", "#ffbf00"), ("カーキ", "#f0e68c")
            ]),
            ("紫系", [
                ("紫", "#5c2d91"), ("ラベンダー", "#b4a0ff"), ("バイオレット", "#ee82ee"),
                ("プラム", "#8b008b"), ("インディゴ", "#4b0082"), ("オーキッド", "#da70d6")
            ]),
            ("茶・土系", [
                ("茶色", "#8b4513"), ("オリーブ", "#808000"), ("テラコッタ", "#e2725b"),
                ("チョコ", "#d2691e"), ("ベージュ", "#f5f5dc"), ("タン", "#d2b48c")
            ]),
            ("無彩色系", [
                ("黒", "#323130"), ("灰色", "#808080"), ("シルバー", "#c0c0c0"),
                ("白鼠", "#e0e0e0"), ("スレートグレー", "#708090"), ("濃灰", "#404040")
            ])
        ]

    def get_color_name(self, hex_code):
        groups = self.get_color_groups()
        for gn, colors in groups:
            for name, code in colors:
                if code.lower() == hex_code.lower():
                    return name
        return "不明"

    def format_summary_workload(self, data):
        if isinstance(data, dict) and data.get('is_group_summary'):
            m = data['manual']
            a = data['auto']
            diff = m - a
            if m == 0 and a == 0:
                return "-"
            return f"人:{m:g} 案:{a:g}\n差:{diff:g}"

        day_map = data
        if not day_map: return "0工数"
        total = sum(day_map.values())
        total_str = f"{total:g}"
        if len(day_map) <= 1:
            return f"{total_str}工数"
        
        parts = []
        for code in sorted(day_map.keys()):
            days = day_map[code]
            name = self.get_color_name(code)
            parts.append(f"{name}:{days:g}")
        return f"計{total_str}工数 ({', '.join(parts)})"

    def toggle_column_visibility(self, idx, visible):
        if idx < 8:
            self.table.setColumnHidden(idx, not visible)
        else:
            self.summary_visible = visible
            for i in range(8, self.table.columnCount()):
                self.table.setColumnHidden(i, not visible)
        
        if idx in self.col_actions:
            self.col_actions[idx].blockSignals(True)
            self.col_actions[idx].setChecked(visible)
            self.col_actions[idx].blockSignals(False)

    def closeEvent(self, event):
        current_snapshot = self.get_current_data_snapshot()
        is_modified = getattr(self, 'saved_snapshot', "") != current_snapshot
        
        if is_modified:
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle('確認')
            msg_box.setText('未保存の変更があります。保存して終了しますか？\n（「破棄」を選ぶと保存せずに終了します）')
            msg_box.setIcon(QMessageBox.Question)

            btn_save = msg_box.addButton("保存", QMessageBox.AcceptRole)
            btn_discard = msg_box.addButton("破棄", QMessageBox.DestructiveRole)
            btn_cancel = msg_box.addButton("キャンセル", QMessageBox.RejectRole)

            msg_box.exec()

            if msg_box.clickedButton() == btn_save:
                if not self.last_path:
                    save_success = self.save_data_as()
                else:
                    save_success = self.save_data()
                    
                if save_success is True:
                    self._cleanup_before_close()
                    event.accept()
                else:
                    event.ignore()
                    
            elif msg_box.clickedButton() == btn_discard:
                self._cleanup_before_close()
                event.accept()
            else:
                event.ignore()
        else:
            self._cleanup_before_close()
            event.accept()

    def _cleanup_before_close(self):
        try:
            if hasattr(self, 'snap_timer'):
                self.snap_timer.stop()
        except RuntimeError:
            pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = GanttApp()
    window.showMaximized()
    sys.exit(app.exec())