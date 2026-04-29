import sys
import os
import json
import calendar
import copy
import shiboken6
from datetime import datetime, timedelta
import jpholiday
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, 
                               QHeaderView, QSplitter, QGraphicsView, QGraphicsScene, 
                               QDialog, QFormLayout, QLineEdit, QDateEdit, QMessageBox, 
                               QFileDialog, QGraphicsRectItem, QGraphicsTextItem, QSlider, QLabel, QMenu, QSpinBox, QColorDialog, QComboBox, QInputDialog, QAbstractItemView, QScrollArea, QGridLayout, QTabWidget, QTextBrowser)
from PySide6.QtCore import Qt, QDate, QRectF, QPointF, QTimer
from PySide6.QtGui import QBrush, QPen, QColor, QFont, QPainter, QPainterPath, QPixmap, QIcon, QCursor

# TaskDialog was removed in favor of inline editing.

from dialogs import SettingsDialog, ColorGridDialog, SummaryDialog, HelpDialog
from graphics import GanttBarItem, HeaderScene, ChartScene
from components import HideableHeader, TaskTable

from mixins_file_io import FileIOMixin
from mixins_chart import ChartMixin
from mixins_events import EventsMixin
from mixins_sync import SyncMixin
from mixins_util import UtilMixin

class GanttApp(QMainWindow, FileIOMixin, ChartMixin, EventsMixin, SyncMixin, UtilMixin):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MiyaGantt v2.0.0 - Professional Gantt Chart")
        self.resize(1380, 850)
        self.tasks = []
        self.undo_stack = []
        self.redo_stack = []
        self.day_width = 40
        self.row_height = 40
        self.header_height = 70
        self.min_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        self.display_unit = 1  # 0: 週間, 1: 月間, 2: 年間
        self.display_count = 6
        self.zoom_unit = 1     # 0: 週間, 1: 月間, 2: 年間
        self.zoom_count = 1    # デフォルトで1単位分を1画面に収める
        self.update_display_days()
        self.month_label_items = []
        self.visible_tasks_info = [] # [{index, task, indent}]
        self.custom_holidays = {} # カスタム祝日 { 'YYYY-MM-DD': '祝日名' }
        self.summary_visible = True
        self.last_summary_base_key = None
        self.max_date = self.min_date + timedelta(days=180) # 初期範囲
        self.last_path = ""
        self.clipboard_periods = []
        
        self.setWindowIcon(QIcon(self.get_icon_path()))
        
        # スナップ用タイマーの初期化
        self.snap_timer = QTimer()
        self.snap_timer.setSingleShot(True)
        self.snap_timer.timeout.connect(self.snap_horizontal_scroll)
        self.pending_selection = []
        
        self.init_ui()
        self.apply_styles()

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #f0f0f0; }
            QWidget { background-color: #f0f0f0; color: #333333; }
            QTableWidget { background-color: #ffffff; gridline-color: #e0e0e0; border: 1px solid #cccccc; color: #333333; outline: 0; }
            QTableWidget::item:selected { background-color: #e1f0ff; color: #333333; }
            QHeaderView::section { background-color: #e8e8e8; color: #333333; border: 1px solid #cccccc; padding: 4px; font-weight: bold; }
            QPushButton { background-color: #ffffff; border: 1px solid #cccccc; padding: 6px 12px; border-radius: 4px; }
            QPushButton:hover { background-color: #e8e8e8; }
            QPushButton:disabled { background-color: #f5f5f5; color: #bbbbbb; border: 1px solid #eeeeee; }
        """)

    def init_ui(self):
        mw = QWidget()
        self.setCentralWidget(mw)
        ml = QVBoxLayout(mw)
        tl = QHBoxLayout()
        
        self.btn_add = QPushButton("追加")
        self.btn_group = QPushButton("グループ")
        self.btn_up = QPushButton("↑")
        self.btn_down = QPushButton("↓")
        self.btn_del = QPushButton("削除")
        self.btn_undo = QPushButton("戻す")
        self.btn_redo = QPushButton("進む")
        self.btn_add.clicked.connect(self.add_task)
        self.btn_group.clicked.connect(self.add_group)
        self.btn_up.clicked.connect(self.move_row_up)
        self.btn_down.clicked.connect(self.move_row_down)
        self.btn_del.clicked.connect(self.delete_task)
        self.btn_undo.clicked.connect(self.undo)
        self.btn_redo.clicked.connect(self.redo)
        tl.addWidget(self.btn_add)
        tl.addWidget(self.btn_group)
        tl.addWidget(self.btn_up)
        tl.addWidget(self.btn_down)
        tl.addWidget(self.btn_del)
        tl.addWidget(self.btn_undo)
        tl.addWidget(self.btn_redo)
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
        
        self.btn_load = QPushButton("読込")
        self.btn_save = QPushButton("保存")
        self.btn_settings = QPushButton("⚙ 編集期間")
        self.btn_summary = QPushButton("📊 集計")
        self.btn_save_config = QPushButton("💾 設定保存")
        self.btn_help = QPushButton("❓ ヘルプ")
        self.btn_load.clicked.connect(self.load_data)
        self.btn_save.clicked.connect(self.save_data)
        self.btn_settings.clicked.connect(self.open_settings)
        self.btn_summary.clicked.connect(self.open_summary)
        self.btn_save_config.clicked.connect(self.save_app_config)
        self.btn_help.clicked.connect(self.open_help)
        
        tl.addWidget(self.btn_load)
        tl.addWidget(self.btn_save)
        tl.addWidget(self.btn_settings)
        tl.addWidget(self.btn_summary)
        tl.addWidget(self.btn_save_config)
        tl.addWidget(self.btn_help)
        tl.addStretch()
        
        ml.addLayout(tl)
        
        # 移動ナビゲーション用の2段目ツールバー
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
        
        # テーブル側のコンテナ (ボタン + テーブル)
        self.left_container = QWidget()
        self.left_layout = QVBoxLayout(self.left_container)
        self.left_layout.setContentsMargins(0, 0, 0, 0)
        self.left_layout.setSpacing(2)
        
        # 列表示切り替えボタン群 (ヘッダーのすぐ上に配置)
        self.col_toggle_layout = QHBoxLayout()
        self.col_toggle_layout.setContentsMargins(2, 2, 2, 2)
        self.col_toggle_layout.setSpacing(2)
        self.col_actions = {}
        col_info = [
            (0, "マーク"), (1, "開閉"), (3, "種別"), (4, "進捗"), (5, "人数/工数"),
            (6, "期間/開始日"), (7, "色"), (8, "集計列")
        ]
        for idx, name in col_info:
            btn = QPushButton(name) # 👁 を削除
            btn.setCheckable(True)
            btn.setChecked(True)
            btn.setFixedHeight(24)
            btn.setStyleSheet("""
                QPushButton { background-color: #f8f8f8; border: 1px solid #ddd; border-radius: 4px; color: #666; font-size: 10px; padding: 0 5px; }
                QPushButton:checked { background-color: #e1f0ff; border: 1px solid #0078d4; color: #0078d4; font-weight: bold; }
                QPushButton:hover { background-color: #eeeeee; }
            """)
            btn.clicked.connect(lambda checked, i=idx: self.toggle_column_visibility(i, checked))
            self.col_actions[idx] = btn
            self.col_toggle_layout.addWidget(btn)
        self.col_toggle_layout.addStretch()
        self.left_layout.addLayout(self.col_toggle_layout)

        self.table = TaskTable(0, 8) # 集計列は update_ui で動的に追加される
        self.table.setHorizontalHeaderLabels(["", "", "タスク名", "種別", "進捗(%)", "人数/工数", "期間指定/開始日", "色"])
        self.table.setColumnWidth(0, 25)   # マーク
        self.table.setColumnWidth(1, 35)   # 開閉
        self.table.setColumnWidth(2, 200)  # タスク名 (初期幅)
        self.table.setColumnWidth(3, 40)   # 種別
        self.table.setColumnWidth(4, 60)   # 進捗(%)
        self.table.setColumnWidth(5, 70)   # 人数/工数
        self.table.setColumnWidth(6, 165)  # 期間指定/開始日
        self.table.setColumnWidth(7, 40)   # 色
        self.table.horizontalHeader().setFixedHeight(self.header_height)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.table.verticalHeader().setDefaultSectionSize(self.row_height)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        
        self.table.itemChanged.connect(self.on_table_item_changed)
        self.table.cellClicked.connect(self.on_table_cell_clicked)
        self.table.cellDoubleClicked.connect(self.on_table_cell_double_clicked)
        self.table.currentCellChanged.connect(self.update_selection_mark)
        
        self.left_layout.addWidget(self.table)
        self.splitter.addWidget(self.left_container)
        
        rc = QWidget()
        rcl = QVBoxLayout(rc); rcl.setContentsMargins(0,0,0,0); rcl.setSpacing(0)
        
        # 左側のボタンエリアと同じ高さのスペーサーを入れて行のズレを防ぐ
        spacer = QWidget()
        spacer.setFixedHeight(28) # col_toggle_layout の高さ(24) + margins(4)
        rcl.addWidget(spacer)
        
        self.hs = HeaderScene(self)
        self.cs = ChartScene(self)
        self.hv = QGraphicsView(self.hs)
        self.chart_view = QGraphicsView(self.cs)
        self.chart_view.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
        
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
        
        self.update_undo_redo_buttons()

    def on_horizontal_scroll(self, v):
        # スクロール中は同期のみ行い、滑らかに移動させる
        self.hv.horizontalScrollBar().setValue(v)
        self.update_month_labels_pos()
        
        # スクロール停止後にスナップさせるためタイマーを開始
        # スライダー操作中以外（ホイール等）の場合に有効
        if not self.chart_view.horizontalScrollBar().isSliderDown():
            self.snap_timer.start(300) # 300ms後にスナップ実行
        
        # スクロール位置に応じた集計列の同期
        if self.day_width <= 0: return
        days_scrolled = v / self.day_width
        visible_start = self.min_date + timedelta(days=days_scrolled)
        threshold_date = self.get_threshold_date(visible_start)
        
        # 現在の「基準単位」を特定して、変更があった場合のみ更新
        if self.display_unit == 0: # 週間
            base_key = (threshold_date - timedelta(days=threshold_date.weekday())).strftime("%Y-%W")
        elif self.display_unit == 1: # 月間
            base_key = threshold_date.strftime("%Y-%m")
        else: # 年間
            base_key = threshold_date.strftime("%Y")
            
        if base_key != self.last_summary_base_key:
            self.last_summary_base_key = base_key
            self.sync_summary_to_scroll(threshold_date)

    def snap_horizontal_scroll(self):
        if self.day_width > 0:
            v = self.chart_view.horizontalScrollBar().value()
            snapped = round(v / self.day_width) * self.day_width
            self.chart_view.horizontalScrollBar().setValue(int(snapped))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'chart_view'):
            self.calculate_day_width()
            self.update_ui()

    def get_icon_path(self):
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, 'icon.ico')
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icon.ico')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = GanttApp()
    window.showMaximized()
    sys.exit(app.exec())
