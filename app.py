import sys
import json
import calendar
from datetime import datetime, timedelta
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, 
                               QHeaderView, QSplitter, QGraphicsView, QGraphicsScene, 
                               QDialog, QFormLayout, QLineEdit, QDateEdit, QMessageBox, 
                               QFileDialog, QGraphicsRectItem, QGraphicsTextItem, QSlider, QLabel, QMenu, QSpinBox, QColorDialog, QComboBox, QInputDialog, QAbstractItemView)
from PySide6.QtCore import Qt, QDate, QRectF, QPointF, QTimer
from PySide6.QtGui import QBrush, QPen, QColor, QFont, QPainter, QPainterPath

# TaskDialog was removed in favor of inline editing.

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("表示設定")
        self.layout = QFormLayout(self)
        
        self.start_date_edit = QDateEdit(self)
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDate(parent.min_date.date() if parent else QDate.currentDate())
        
        self.unit_combo = QComboBox(self)
        self.unit_combo.addItems(["週間", "月間", "年間"])
        self.unit_combo.setCurrentIndex(parent.display_unit if parent else 1)
        
        self.count_spinbox = QSpinBox(self)
        self.count_spinbox.setRange(1, 500)
        self.count_spinbox.setValue(parent.display_count if parent else 6)
        
        self.layout.addRow("表示開始日:", self.start_date_edit)
        self.layout.addRow("表示単位:", self.unit_combo)
        self.layout.addRow("表示数:", self.count_spinbox)
        
        self.btn_ok = QPushButton("OK", self)
        self.btn_ok.clicked.connect(self.accept)
        self.layout.addRow(self.btn_ok)

class GanttBarItem(QGraphicsRectItem):
    def __init__(self, task, row, period_index, gantt_app, rect=None):
        super().__init__(rect)
        self.task = task
        self.row = row
        self.period_index = period_index
        self.app = gantt_app
        self.setFlags(QGraphicsRectItem.ItemIsMovable | 
                      QGraphicsRectItem.ItemIsSelectable | 
                      QGraphicsRectItem.ItemSendsGeometryChanges)
        self.setAcceptHoverEvents(True)
        
        self.progress_item = QGraphicsRectItem(self)
        # 初期テキストの設定（update_appearanceで上書きされるため空でも可）
        self.text_item = QGraphicsTextItem('', self)
        self.text_item.setDefaultTextColor(Qt.white)
        self.text_item.setZValue(1)
        font = QFont("Segoe UI", 9, QFont.Bold)
        self.text_item.setFont(font)
        self.progress_item.setAcceptHoverEvents(False)
        self.text_item.setAcceptHoverEvents(False)
        
        self.resizing_left = False
        self.resizing_right = False
        self.update_appearance()

    def update_appearance(self):
        bc = QColor(self.task.get('color', '#0078d4'))
        self.setPen(QPen(Qt.black if self.isSelected() else bc.darker(120), 2 if self.isSelected() else 1))
        self.setBrush(QBrush(bc.lighter(150)))
        
        prog = self.task.get('progress', 0)
        
        periods = self.task.get('periods', [self.task])
        valid_periods = []
        for i, p in enumerate(periods):
            try:
                sd = datetime.strptime(p.get('start_date', ''), "%Y-%m-%d")
                ed = datetime.strptime(p.get('end_date', ''), "%Y-%m-%d")
                valid_periods.append({'idx': i, 'start': sd, 'end': ed, 'days': (ed - sd).days + 1})
            except Exception:
                pass
                
        valid_periods.sort(key=lambda x: x['start'])
        total_days = sum(p['days'] for p in valid_periods)
        
        target_days = total_days * (prog / 100.0)
        
        days_allocated_to_this = 0
        for p in valid_periods:
            if target_days <= 0:
                break
            allocate = min(p['days'], target_days)
            if p['idx'] == self.period_index:
                days_allocated_to_this = allocate
                break
            target_days -= allocate
            
        this_period = next((p for p in valid_periods if p['idx'] == self.period_index), None)
        local_prog_ratio = (days_allocated_to_this / this_period['days']) if (this_period and this_period['days'] > 0) else 0
        
        p_rect = QRectF(self.rect().left(), self.rect().top(), self.rect().width() * local_prog_ratio, self.rect().height())
        self.progress_item.setRect(p_rect)
        self.progress_item.setBrush(QBrush(bc))
        self.progress_item.setPen(Qt.NoPen)
        # ツールチップ更新用のデータ取得
        periods = self.task.get('periods')
        if periods is not None and self.period_index < len(periods):
            p_dict = periods[self.period_index]
        else:
            p_dict = self.task

        # バー固有のテキスト（無ければ空）を表示
        bar_text = p_dict.get('text', '')
        self.text_item.setPlainText(bar_text)
        self.text_item.setPos(5, (self.rect().height() - self.text_item.boundingRect().height()) / 2)

        start_d = p_dict.get('start_date', '')
        end_d = p_dict.get('end_date', '')
        self.setToolTip(f"タスク: {self.task.get('name','')}\n期間: {start_d}〜{end_d}")

    def hoverMoveEvent(self, event):
        x = event.pos().x()
        w = self.rect().width()
        # 1日の場合や幅が狭い場合でも確実に反応するように調整
        margin = 12 if w <= self.app.day_width else 10
        margin = min(margin, w / 2 - 2)
        if x < margin or x > w - margin:
            self.setCursor(Qt.SizeHorCursor)
        else:
            self.setCursor(Qt.OpenHandCursor)
        super().hoverMoveEvent(event)

    def hoverLeaveEvent(self, event):
        self.setCursor(Qt.ArrowCursor)
        super().hoverLeaveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            x = event.pos().x()
            w = self.rect().width()
            margin = 12 if w <= self.app.day_width else 10
            margin = min(margin, w / 2 - 2)
            if x < margin:
                self.resizing_left = True
            elif x > w - margin:
                self.resizing_right = True
            else:
                self.setCursor(Qt.ClosedHandCursor)
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        snap = self.app.day_width
        if self.resizing_left:
            diff = event.scenePos().x() - event.lastScenePos().x()
            nr = self.rect()
            if nr.width() - diff >= snap:
                self.setPos(self.pos().x() + diff, self.pos().y())
                self.setRect(0, 0, nr.width() - diff, nr.height())
        elif self.resizing_right:
            diff = event.scenePos().x() - event.lastScenePos().x()
            nr = self.rect()
            if nr.width() + diff >= snap:
                self.setRect(0, 0, nr.width() + diff, nr.height())
        else:
            super().mouseMoveEvent(event)
            # マウスカーソルの位置に基づいて行の中心にスナップ
            row = int(event.scenePos().y() / self.app.row_height)
            # タスクが存在する行の範囲内に制限
            max_row = len(self.app.visible_tasks_info) - 1 if self.app.visible_tasks_info else 0
            row = max(0, min(max_row, row))
            self.setPos(self.pos().x(), row * self.app.row_height + 10)
        self.update_appearance()

    def mouseReleaseEvent(self, event):
        was_resizing = self.resizing_left or self.resizing_right
        self.resizing_left = self.resizing_right = False
        self.setCursor(Qt.OpenHandCursor)
        snap = self.app.day_width
        sx = round(self.pos().x() / snap) * snap
        sw = max(snap, round(self.rect().width() / snap) * snap)
        
        sd = self.app.min_date + timedelta(days=sx / self.app.day_width)
        ed = sd + timedelta(days=sw / self.app.day_width - 0.001)

        # 移動先の行を判定
        new_row = int(event.scenePos().y() / self.app.row_height)
        max_row = len(self.app.visible_tasks_info) - 1 if self.app.visible_tasks_info else 0
        new_row = max(0, min(max_row, new_row))
        
        # サイズ調整中ではなかった場合のみ行移動を許可
        if not was_resizing and new_row != self.row:
            target_info = self.app.visible_tasks_info[new_row]
            target_task = target_info['task']
            
            # グループ行への移動は禁止
            if target_task.get('is_group'):
                QTimer.singleShot(0, self.app.update_ui)
                super().mouseReleaseEvent(event)
                return

            # 移動元・移動先の両方で 'periods' 形式を確定させる
            for t in [self.task, target_task]:
                if 'periods' not in t:
                    t['periods'] = [{'start_date': t.get('start_date', ''), 'end_date': t.get('end_date', '')}]
            
            if 0 <= self.period_index < len(self.task['periods']):
                # 期間データを移動
                p = self.task['periods'].pop(self.period_index)
                p['start_date'] = sd.strftime("%Y-%m-%d")
                p['end_date'] = ed.strftime("%Y-%m-%d")
                
                target_task = self.app.tasks[new_row]
                target_task['periods'].append(p)
                
                # 互換性のためメインの日付フィールドも更新
                for t in [self.task, target_task]:
                    if t['periods']:
                        t['start_date'] = t['periods'][0]['start_date']
                        t['end_date'] = t['periods'][0]['end_date']
                
                QTimer.singleShot(0, self.app.update_ui)
            
            super().mouseReleaseEvent(event)
            return

        self.setPos(sx, self.pos().y())
        self.setRect(0, 0, sw, self.rect().height())
        super().mouseReleaseEvent(event)
        
        if 'periods' not in self.task:
            self.task['periods'] = [{'start_date': self.task.get('start_date', ''), 'end_date': self.task.get('end_date', '')}]
            
        self.task['periods'][self.period_index]['start_date'] = sd.strftime("%Y-%m-%d")
        self.task['periods'][self.period_index]['end_date'] = ed.strftime("%Y-%m-%d")
        
        # update single fields for backwards compatibility
        self.task['start_date'] = self.task['periods'][0]['start_date']
        self.task['end_date'] = self.task['periods'][0]['end_date']
        
        if self.scene():
            for item in self.scene().items():
                if isinstance(item, GanttBarItem) and item.task is self.task:
                    item.update_appearance()
                    
        self.app.sync_table_from_tasks()

    def mouseDoubleClickEvent(self, event):
        # Prevent default double click which was old edit open
        super().mouseDoubleClickEvent(event)
        
        # ダブルクリックでバー固有のテキストを編集
        if 'periods' not in self.task:
            self.task['periods'] = [{'start_date': self.task.get('start_date', ''), 'end_date': self.task.get('end_date', '')}]
            
        if 0 <= self.period_index < len(self.task['periods']):
            p_dict = self.task['periods'][self.period_index]
            current_text = p_dict.get('text', '')
            
            text, ok = QInputDialog.getText(self.app, "テキストの編集", "バーに表示するテキスト:", QLineEdit.Normal, current_text)
            if ok:
                p_dict['text'] = text
                self.update_appearance()
                self.app.update_ui() # 全体を再描画して確実に反映させる

    def contextMenuEvent(self, event):
        menu = QMenu()
        del_action = menu.addAction("削除")
        action = menu.exec(event.screenPos())
        if action == del_action:
            if self.task in self.app.tasks:
                self.app.tasks.remove(self.task)
                QTimer.singleShot(0, self.app.update_ui)

class GanttGroupItem(QGraphicsRectItem):
    def __init__(self, task, row, start_date, end_date, gantt_app):
        super().__init__()
        self.task = task
        self.row = row
        self.app = gantt_app
        self.start_date = start_date
        self.end_date = end_date
        self.setZValue(25)
        self.update_appearance()

    def update_appearance(self):
        try:
            sd = datetime.strptime(self.start_date, "%Y-%m-%d")
            ed = datetime.strptime(self.end_date, "%Y-%m-%d")
            x = (sd - self.app.min_date).days * self.app.day_width
            w = ((ed - sd).days + 1) * self.app.day_width
            y = self.row * self.app.row_height + 12
            h = 10
            
            # ブラケット形状のパスを作成
            path = QPainterPath()
            path.moveTo(x, y + h)
            path.lineTo(x, y)
            path.lineTo(x + w, y)
            path.lineTo(x + w, y + h)
            path.lineTo(x + w - 5, y + h - 5)
            path.lineTo(x + w - 5, y + 5)
            path.lineTo(x + 5, y + 5)
            path.lineTo(x + 5, y + h - 5)
            path.closeSubpath()
            
            # 簡易的に矩形で表現（パスが複雑な場合はこちら）
            self.setRect(x, y, w, h)
            self.setBrush(QBrush(QColor(60, 60, 60)))
            self.setPen(QPen(Qt.black, 1))
            
            self.setToolTip(f"グループ: {self.task.get('name','')}\n期間: {self.start_date}〜{self.end_date}")
        except:
            pass

class ChartScene(QGraphicsScene):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.start_x = 0

    def mousePressEvent(self, e):
        item = self.itemAt(e.scenePos(), self.app.chart_view.transform())
        if not item and e.button() == Qt.LeftButton:
            self.start_x = e.scenePos().x()
        super().mousePressEvent(e)

    def mouseReleaseEvent(self, e):
        if self.start_x > 0:
            item = self.itemAt(e.scenePos(), self.app.chart_view.transform())
            if not item:
                if abs(e.scenePos().x() - self.start_x) > (self.app.day_width * 0.1):
                    self.app.create_task_from_drag(self.start_x, e.scenePos().x(), e.scenePos().y())
            self.start_x = 0
        super().mouseReleaseEvent(e)

    def mouseDoubleClickEvent(self, e):
        items = self.items(e.scenePos(), Qt.IntersectsItemShape, Qt.DescendingOrder, self.app.chart_view.transform())
        gantt_item = next((it for it in items if isinstance(it, (GanttBarItem, GanttGroupItem))), None)
            
        if not gantt_item and e.button() == Qt.LeftButton:
            y = e.scenePos().y()
            row = int(y / self.app.row_height)
            if 0 <= row < len(self.app.visible_tasks_info):
                info = self.app.visible_tasks_info[row]
                task = info['task']
                if task.get('is_group'):
                    # グループの場合は折り畳み
                    task['collapsed'] = not task.get('collapsed', False)
                    self.app.update_ui()
                    return

                x = e.scenePos().x()
                day_idx = int(x / self.app.day_width)
                d_str = (self.app.min_date + timedelta(days=day_idx)).strftime("%Y-%m-%d")
                
                if 'periods' not in task:
                    task['periods'] = [{'start_date': task.get('start_date', ''), 'end_date': task.get('end_date', '')}]
                
                task['periods'].append({"start_date": d_str, "end_date": d_str, "text": ""})
                self.app.update_ui()
                e.accept()
                return
        super().mouseDoubleClickEvent(e)

    def contextMenuEvent(self, e):
        item = self.itemAt(e.scenePos(), self.app.chart_view.transform())
        if item and not isinstance(item, GanttBarItem):
            item = None
            
        if not item:
            y = e.scenePos().y()
            row = int(y / self.app.row_height)
            menu = QMenu()
            
            if 0 <= row < len(self.app.visible_tasks_info):
                info = self.app.visible_tasks_info[row]
                task = info['task']
                if task.get('is_group'):
                    add_task_in_group = menu.addAction("このグループにタスクを追加")
                    add_period_action = None
                else:
                    task_name = task.get('name', '無題')
                    add_period_action = menu.addAction(f"「{task_name}」に期間を追加")
                    add_task_in_group = None
            else:
                add_period_action = None
                add_task_in_group = None
            
            action = menu.exec(e.screenPos())
            x = e.scenePos().x()
            day_idx = int(x / self.app.day_width)
            d_str = (self.app.min_date + timedelta(days=day_idx)).strftime("%Y-%m-%d")
            
            if action == add_period_action and add_period_action:
                if 'periods' not in task:
                    task['periods'] = [{'start_date': task.get('start_date', ''), 'end_date': task.get('end_date', '')}]
                task['periods'].append({"start_date": d_str, "end_date": d_str})
                self.app.update_ui()
            elif action == add_task_in_group and add_task_in_group:
                new_task = {
                    "name": "新規タスク",
                    "periods": [{"start_date": d_str, "end_date": d_str}],
                    "progress": 0,
                    "color": "#0078d4"
                }
                # グループの直後に挿入
                self.app.tasks.insert(info['index'] + 1, new_task)
                self.app.update_ui()
        else:
            super().contextMenuEvent(e)

class TaskTable(QTableWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setSelectionMode(QTableWidget.SingleSelection)
        self.setSelectionBehavior(QTableWidget.SelectRows)

class GanttApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MiyaGantt - Professional Gantt Chart")
        self.resize(1380, 850)
        self.tasks = []
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
        
        self.init_ui()
        self.apply_styles()

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #f0f0f0; }
            QWidget { background-color: #f0f0f0; color: #333333; }
            QTableWidget { background-color: #ffffff; gridline-color: #e0e0e0; border: 1px solid #cccccc; color: #333333; }
            QHeaderView::section { background-color: #e8e8e8; color: #333333; border: 1px solid #cccccc; padding: 4px; font-weight: bold; }
            QPushButton { background-color: #ffffff; border: 1px solid #cccccc; padding: 6px 12px; border-radius: 4px; }
            QPushButton:hover { background-color: #e8e8e8; }
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
        self.btn_add.clicked.connect(self.add_task)
        self.btn_group.clicked.connect(self.add_group)
        self.btn_up.clicked.connect(self.move_row_up)
        self.btn_down.clicked.connect(self.move_row_down)
        self.btn_del.clicked.connect(self.delete_task)
        tl.addWidget(self.btn_add)
        tl.addWidget(self.btn_group)
        tl.addWidget(self.btn_up)
        tl.addWidget(self.btn_down)
        tl.addWidget(self.btn_del)
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
        self.btn_settings = QPushButton("表示設定")
        self.btn_load.clicked.connect(self.load_data)
        self.btn_save.clicked.connect(self.save_data)
        self.btn_settings.clicked.connect(self.open_settings)
        tl.addWidget(self.btn_load)
        tl.addWidget(self.btn_save)
        tl.addWidget(self.btn_settings)
        ml.addLayout(tl)
        
        self.splitter = QSplitter(Qt.Horizontal)
        ml.addWidget(self.splitter)
        
        self.table = TaskTable(0, 5)
        self.table.setHorizontalHeaderLabels(["", "タスク名", "進捗(%)", "期間指定", "色"])
        self.table.setColumnWidth(0, 30)
        self.table.horizontalHeader().setFixedHeight(self.header_height)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.verticalHeader().setDefaultSectionSize(self.row_height)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        
        self.table.itemChanged.connect(self.on_table_item_changed)
        self.table.cellClicked.connect(self.on_table_cell_clicked)
        self.table.cellDoubleClicked.connect(self.on_table_cell_double_clicked)
        
        self.splitter.addWidget(self.table)
        
        rc = QWidget()
        rcl = QVBoxLayout(rc); rcl.setContentsMargins(0,0,0,0); rcl.setSpacing(0)
        
        self.hs = QGraphicsScene()
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
        self.splitter.setSizes([380, 1000])
        
        self.chart_view.horizontalScrollBar().valueChanged.connect(self.on_horizontal_scroll)
        self.table.verticalScrollBar().valueChanged.connect(self.chart_view.verticalScrollBar().setValue)
        self.chart_view.verticalScrollBar().valueChanged.connect(self.table.verticalScrollBar().setValue)

    def on_horizontal_scroll(self, v):
        self.hv.horizontalScrollBar().setValue(v)
        self.update_month_labels_pos()

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

    def change_zoom(self, v):
        self.day_width = v
        self.update_ui()

    def calculate_day_width(self):
        # 1画面に収める日数を計算
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

    def on_zoom_changed(self, *_):
        self.zoom_unit = self.zoom_unit_combo.currentIndex()
        self.zoom_count = self.zoom_count_spin.value()
        self.calculate_day_width()
        self.update_ui()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'chart_view'):
            self.calculate_day_width()
            self.draw_chart()

    def update_display_range(self):
        self.update_display_days()
        self.update_ui()

    def update_display_days(self):
        # 単位と数に基づいて表示日数を計算する
        if self.display_unit == 0: # 週間
            self.display_days = self.display_count * 7
        elif self.display_unit == 1: # 月間
            m = self.min_date.month - 1 + self.display_count
            y = self.min_date.year + m // 12
            m = m % 12 + 1
            last_day = calendar.monthrange(y, m)[1]
            d = min(self.min_date.day, last_day)
            end_date = datetime(y, m, d)
            self.display_days = max(1, (end_date - self.min_date).days)
        elif self.display_unit == 2: # 年間
            count_months = self.display_count * 12
            m = self.min_date.month - 1 + count_months
            y = self.min_date.year + m // 12
            m = m % 12 + 1
            last_day = calendar.monthrange(y, m)[1]
            d = min(self.min_date.day, last_day)
            end_date = datetime(y, m, d)
            self.display_days = max(1, (end_date - self.min_date).days)

    def parse_date(self, s):
        s = s.strip().replace('/', '-')
        parts = s.split('-')
        now = datetime.now()
        if len(parts) == 3: # YYYY-MM-DD
            return datetime.strptime(s, "%Y-%m-%d").strftime("%Y-%m-%d")
        elif len(parts) == 2: # MM-DD
            return f"{now.year}-{int(parts[0]):02d}-{int(parts[1]):02d}"
        return None

    def get_periods_from_string(self, text):
        text = text.replace(' ', '').replace(';', ',').replace('~', '-')
        parts = text.split(',')
        periods = []
        for part in parts:
            if not part: continue
            if '-' in part:
                s_part, e_part = part.split('-', 1)
                start_d = self.parse_date(s_part)
                end_d = self.parse_date(e_part)
                if not start_d or not end_d: return None
                periods.append({"start_date": start_d, "end_date": end_d})
        return periods

    def move_row_up(self):
        row = self.table.currentRow()
        if row <= 0: return
        
        info = self.visible_tasks_info[row]
        task_to_track = info['task']
        
        if task_to_track.get('is_group'):
            # グループを上に移動：前のグループのさらに上へ
            target_v_row = row - 1
            while target_v_row > 0 and not self.visible_tasks_info[target_v_row]['task'].get('is_group'):
                target_v_row -= 1
            self.move_tasks([row], target_v_row)
        else:
            # タスクを上に移動：1行上へ
            self.move_tasks([row], row - 1)
            
        # 移動後の位置を特定して再選択
        for i, n_info in enumerate(self.visible_tasks_info):
            if n_info['task'] == task_to_track:
                self.table.setCurrentCell(i, 1)
                break

    def move_row_down(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self.visible_tasks_info) - 1: return
        
        info = self.visible_tasks_info[row]
        task_to_track = info['task']
        
        if task_to_track.get('is_group'):
            # グループを下に移動：次のグループのさらに下へ
            # まず現在のグループの終わり（可視行）を探す
            current_group_end_v = row + 1
            end_idx = info['index'] + 1
            for i in range(info['index'] + 1, len(self.tasks)):
                if self.tasks[i].get('is_group'): break
                end_idx = i + 1
            for i in range(row + 1, len(self.visible_tasks_info)):
                if self.visible_tasks_info[i]['index'] >= end_idx: break
                current_group_end_v = i + 1
            
            # 次のグループを探す
            next_group_v = -1
            for i in range(current_group_end_v, len(self.visible_tasks_info)):
                if self.visible_tasks_info[i]['task'].get('is_group'):
                    next_group_v = i
                    break
            
            if next_group_v == -1:
                # 次のグループがない場合は末尾へ
                target_v_row = len(self.visible_tasks_info)
            else:
                # 次のグループの末尾を探す
                target_v_row = next_group_v + 1
                ng_end_idx = self.visible_tasks_info[next_group_v]['index'] + 1
                for i in range(self.visible_tasks_info[next_group_v]['index'] + 1, len(self.tasks)):
                    if self.tasks[i].get('is_group'): break
                    ng_end_idx = i + 1
                for i in range(next_group_v + 1, len(self.visible_tasks_info)):
                    if self.visible_tasks_info[i]['index'] >= ng_end_idx: break
                    target_v_row = i + 1
            
            self.move_tasks([row], target_v_row)
        else:
            # タスクを下に移動：1行下へ
            # 下の要素がグループだった場合や、通常のタスクだった場合でも、row+2 を指定すればその要素の後ろに行く
            self.move_tasks([row], row + 2)
            
        # 移動後の位置を特定して再選択
        for i, n_info in enumerate(self.visible_tasks_info):
            if n_info['task'] == task_to_track:
                self.table.setCurrentCell(i, 1)
                break

    def move_tasks(self, source_rows, target_row, refresh_chart=True):
        if not source_rows: return
        
        src_v_row = source_rows[0]
        if src_v_row >= len(self.visible_tasks_info): return
        
        info = self.visible_tasks_info[src_v_row]
        src_idx = info['index']
        is_group = info['task'].get('is_group', False)
        
        # 移動するブロック（実インデックスの範囲）を特定
        start_idx = src_idx
        end_idx = src_idx + 1
        if is_group:
            # グループの場合は、次のグループが現れるまでの全タスクをブロックとする
            for i in range(src_idx + 1, len(self.tasks)):
                if self.tasks[i].get('is_group'):
                    break
                end_idx = i + 1
        
        block = self.tasks[start_idx:end_idx]
        
        # ブロックに含まれる「可視行」の数をカウント（無効なドロップ判定のため）
        visible_count = 1
        if is_group and not info['task'].get('collapsed'):
            for i in range(src_v_row + 1, len(self.visible_tasks_info)):
                if self.visible_tasks_info[i]['index'] >= end_idx:
                    break
                visible_count += 1
        
        # データの並べ替えを実行
        # 1. 移動するブロックを一旦取り出す
        remaining_tasks = self.tasks[:start_idx] + self.tasks[end_idx:]
        
        # 2. 挿入位置にあるタスクを特定して、残ったリスト内での位置を探す
        target_task = None
        if target_row < len(self.visible_tasks_info):
            target_task = self.visible_tasks_info[target_row]['task']
            
        # 自分自身のブロック内へのドロップは無視
        if target_task in block:
            return

        if target_task is None:
            new_target_idx = len(remaining_tasks)
        else:
            try:
                new_target_idx = remaining_tasks.index(target_task)
            except ValueError:
                # 万が一見つからない場合は末尾へ
                new_target_idx = len(remaining_tasks)
                
        # グループ移動時の制約：タスク間には移動できない（次のグループの直前または末尾にスナップさせる）
        if is_group:
            while new_target_idx < len(remaining_tasks):
                if remaining_tasks[new_target_idx].get('is_group'):
                    break
                new_target_idx += 1
            
        # 3. 新しい位置にブロックを挿入
        remaining_tasks[new_target_idx:new_target_idx] = block
        self.tasks = remaining_tasks
        
        self.update_ui(refresh_chart)

    def add_task(self):
        t = {
            "name": f"新規タスク {len(self.tasks)+1}",
            "periods": [],
            "progress": 0,
            "color": "#0078d4"
        }
        self.tasks.append(t)
        self.update_ui()
        self.table.editItem(self.table.item(len(self.visible_tasks_info)-1, 1))

    def add_group(self):
        g = {
            "name": f"新規グループ {len(self.tasks)+1}",
            "is_group": True,
            "collapsed": False,
            "color": "#555555"
        }
        self.tasks.append(g)
        self.update_ui()
        self.table.editItem(self.table.item(len(self.visible_tasks_info)-1, 1))

    def delete_task(self):
        r = self.table.currentRow()
        if r >= 0 and r < len(self.visible_tasks_info):
            if QMessageBox.question(self, "確認", "削除しますか？") == QMessageBox.Yes:
                idx = self.visible_tasks_info[r]['index']
                self.tasks.pop(idx)
                self.update_ui()

    def get_visible_tasks_info(self):
        visible = []
        skip_until_next_group = False
        for i, t in enumerate(self.tasks):
            if t.get('is_group'):
                visible.append({'index': i, 'task': t, 'indent': 0})
                skip_until_next_group = t.get('collapsed', False)
            else:
                if not skip_until_next_group:
                    # 前方にグループがあるか確認
                    has_group = any(self.tasks[j].get('is_group') for j in range(i))
                    visible.append({'index': i, 'task': t, 'indent': 1 if has_group else 0})
        return visible

    def get_group_range(self, group_index):
        start_dates = []
        end_dates = []
        for i in range(group_index + 1, len(self.tasks)):
            t = self.tasks[i]
            if t.get('is_group'):
                break
            s, e = self.get_task_dates(t)
            if s: start_dates.append(s)
            if e: end_dates.append(e)
        if start_dates and end_dates:
            return min(start_dates), max(end_dates)
        return None, None

    def get_task_dates(self, task):
        s_dates = []
        e_dates = []
        periods = task.get('periods', [])
        if not periods and 'start_date' in task:
            periods = [{'start_date': task['start_date'], 'end_date': task['end_date']}]
        for p in periods:
            if p.get('start_date'): s_dates.append(p['start_date'])
            if p.get('end_date'): e_dates.append(p['end_date'])
        if s_dates and e_dates:
            return min(s_dates), max(e_dates)
        return "", ""

    def sync_table_from_tasks(self):
        self.table.blockSignals(True)
        for r, info in enumerate(self.visible_tasks_info):
            t = info['task']
            if t.get('is_group'): continue
            
            periods = t.get('periods', [])
            p_strs = []
            for p in periods:
                if not p.get('start_date') or not p.get('end_date'): continue
                s = p['start_date'].replace('-', '/')
                e = p['end_date'].replace('-', '/')
                p_strs.append(f"{s}-{e}")
            
            period_item = self.table.item(r, 3)
            if period_item:
                period_item.setText(", ".join(p_strs))
        self.table.blockSignals(False)

    def create_task_from_drag(self, x1, x2, y):
        snap = self.day_width
        sx = round(min(x1, x2) / snap) * snap
        ex = round(max(x1, x2) / snap) * snap
        if sx == ex: ex += snap
        sd = self.min_date + timedelta(days=sx/self.day_width)
        ed = self.min_date + timedelta(days=ex/self.day_width - 0.001)
        row = max(0, int(y / self.row_height))
        
        t = {
            "name": f"新規 {len(self.tasks)+1}", 
            "periods": [{"start_date": sd.strftime("%Y-%m-%d"), "end_date": ed.strftime("%Y-%m-%d")}],
            "progress": 0, 
            "color": "#0078d4"
        }
        if row < len(self.visible_tasks_info):
            insert_idx = self.visible_tasks_info[row]['index']
            self.tasks.insert(insert_idx, t)
        else:
            self.tasks.append(t)
        self.update_ui()

    def update_ui(self, refresh_chart=True):
        self.visible_tasks_info = self.get_visible_tasks_info()
        self.table.blockSignals(True)
        
        new_rows = len(self.visible_tasks_info)
        if self.table.rowCount() < new_rows:
            self.table.setRowCount(new_rows)
            
        for r, info in enumerate(self.visible_tasks_info):
            t = info['task']
            indent = "    " * info['indent']
            is_group = t.get('is_group', False)
            
            # セルが存在しない場合のみ生成する
            for c in range(5):
                if self.table.item(r, c) is None:
                    self.table.setItem(r, c, QTableWidgetItem())
            
            # 0: Toggle
            toggle_item = self.table.item(r, 0)
            toggle_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            toggle_item.setForeground(QColor(51, 51, 51))
            f = toggle_item.font(); f.setBold(False); toggle_item.setFont(f)
            toggle_item.setBackground(QColor(255, 255, 255))
            
            if is_group:
                toggle_item.setText("▼" if not t.get('collapsed') else "▶")
                toggle_item.setTextAlignment(Qt.AlignCenter)
                f = toggle_item.font(); f.setBold(True); toggle_item.setFont(f)
                toggle_item.setBackground(QColor(242, 242, 242))
            else:
                toggle_item.setText("")

            # 1: Name
            item_name = self.table.item(r, 1)
            item_name.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_name.setForeground(QColor(51, 51, 51))
            f = item_name.font(); f.setBold(False); item_name.setFont(f)
            item_name.setBackground(QColor(255, 255, 255))
            item_name.setText(indent + t.get('name', ''))
            if is_group:
                f = item_name.font(); f.setBold(True); item_name.setFont(f)
                item_name.setBackground(QColor(242, 242, 242))
            
            # 2: Progress
            item_prog = self.table.item(r, 2)
            item_prog.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_prog.setForeground(QColor(51, 51, 51))
            item_prog.setBackground(QColor(255, 255, 255))
            
            # 3: Period
            item_period = self.table.item(r, 3)
            item_period.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_period.setForeground(QColor(51, 51, 51))
            item_period.setBackground(QColor(255, 255, 255))
            
            # 4: Color
            item_color = self.table.item(r, 4)
            item_color.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_color.setForeground(QColor(51, 51, 51))
            item_color.setBackground(QColor(255, 255, 255))

            if is_group:
                item_prog.setText("")
                item_prog.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                item_period.setText("")
                item_period.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                item_color.setText("")
                item_color.setBackground(QColor(200, 200, 200))
                item_color.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            else:
                item_prog.setText(str(t.get('progress', 0)))
                item_prog.setTextAlignment(Qt.AlignCenter)
                
                periods = t.get('periods', [])
                p_strs = []
                for p in periods:
                    if not p.get('start_date') or not p.get('end_date'): continue
                    s = p['start_date'].replace('-', '/')
                    e = p['end_date'].replace('-', '/')
                    p_strs.append(f"{s}-{e}")
                item_period.setText(", ".join(p_strs))
                
                item_color.setText("")
                item_color.setBackground(QColor(t.get('color', '#0078d4')))
                item_color.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                
        # 余分な行を削除
        if self.table.rowCount() > new_rows:
            self.table.setRowCount(new_rows)
            
        self.table.blockSignals(False)
        if refresh_chart:
            self.draw_chart()

    def on_table_item_changed(self, item):
        row = item.row()
        col = item.column()
        if row >= len(self.visible_tasks_info) or row < 0: return
        info = self.visible_tasks_info[row]
        t = info['task']
        
        if col == 1: # Name
            t['name'] = item.text().strip()
        elif col == 2: # Progress
            if t.get('is_group'): return
            try:
                prog = int(item.text().replace('%', '').strip())
                t['progress'] = max(0, min(100, prog))
            except ValueError:
                pass
            self.table.blockSignals(True)
            item.setText(str(t['progress']))
            self.table.blockSignals(False)
        elif col == 3: # Period
            if t.get('is_group'): return
            period_str = item.text()
            parsed = self.get_periods_from_string(period_str)
            if parsed:
                t['periods'] = parsed
            else:
                QMessageBox.warning(self, "エラー", "期間の形式が正しくありません。\n例: 04/01-04/05")
        
        self.draw_chart()

    def on_table_cell_clicked(self, row, col):
        if row >= len(self.visible_tasks_info): return
        if col == 0:
            info = self.visible_tasks_info[row]
            t = info['task']
            if t.get('is_group'):
                t['collapsed'] = not t.get('collapsed', False)
                self.update_ui()

    def on_table_cell_double_clicked(self, row, col):
        if row >= len(self.visible_tasks_info): return
        info = self.visible_tasks_info[row]
        t = info['task']
        
        if col == 4: # Color column
            color = QColorDialog.getColor(QColor(t.get('color', '#0078d4')), self, "色を選択")
            if color.isValid():
                t['color'] = color.name()
                self.update_ui()


    def open_settings(self):
        dlg = SettingsDialog(self)
        if dlg.exec():
            # QDate を Python の datetime に変換
            qd = dlg.start_date_edit.date()
            self.min_date = datetime(qd.year(), qd.month(), qd.day())
            self.display_unit = dlg.unit_combo.currentIndex()
            self.display_count = dlg.count_spinbox.value()
            
            self.update_display_range()

    def draw_chart(self):
        self.month_label_items = []
        self.hs.clear()
        self.cs.clear()
        tw_total = self.display_days * self.day_width
        ch = max(self.chart_view.height(), len(self.tasks) * self.row_height)
        
        last_m = None
        for i in range(self.display_days):
            d = self.min_date + timedelta(days=i)
            x = i * self.day_width
            
            # 背景
            if d.weekday() >= 5:
                bg = QColor(240, 248, 255) if d.weekday()==5 else QColor(255, 240, 240)
                re = self.cs.addRect(x, 0, self.day_width, ch, QPen(Qt.NoPen), QBrush(bg))
                re.setZValue(-20)
                re.setAcceptedMouseButtons(Qt.NoButton)
            
            # グリッド
            gl = self.cs.addLine(x, 0, x, ch, QPen(QColor(220, 220, 220), 1))
            gl.setZValue(-15)
            gl.setAcceptedMouseButtons(Qt.NoButton)
            
            if self.day_width >= 60:
                for h in [6, 12, 18]:
                    sl = self.cs.addLine(x + (self.day_width * h / 24.0), 0, x + (self.day_width * h / 24.0), ch, QPen(QColor(245, 245, 245), 0.5))
                    sl.setZValue(-15)
                    sl.setAcceptedMouseButtons(Qt.NoButton)
            
            # ヘッダー (日付・曜日)
            self.hs.addRect(x, 35, self.day_width, 35, QPen(QColor(210, 210, 210)), QBrush(QColor(248, 248, 248))).setZValue(5)
            
            if self.day_width >= 35:
                dl = self.hs.addText(d.strftime("%d"))
                dl.setDefaultTextColor(QColor(50, 50, 50))
                dl.setFont(QFont("Segoe UI", 9, QFont.Bold))
                dl.setPos(x + (self.day_width/2) - 10, 35)
                dl.setZValue(10)
                
                w_c = QColor(0, 80, 200) if d.weekday()==5 else QColor(220, 0, 0) if d.weekday()==6 else QColor(60, 60, 60)
                yl = self.hs.addText(["月","火","水","木","金","土","日"][d.weekday()])
                yl.setDefaultTextColor(w_c)
                yl.setFont(QFont("Segoe UI", 7))
                yl.setPos(x + (self.day_width/2) - 8, 52)
                yl.setZValue(10)
            
            # 年月ラベル (Sticky)
            if (cm := d.strftime("%Y/%m")) != last_m:
                if self.month_label_items:
                    self.month_label_items[-1][1] = x
                self.hs.addLine(x, 0, x, 35, QPen(QColor(150, 150, 150), 2)).setZValue(15)
                ml = self.hs.addText(d.strftime("%Y年 %m月"))
                ml.setDefaultTextColor(QColor(0, 90, 180))
                ml.setFont(QFont("Segoe UI", 11, QFont.Bold))
                ml.setPos(x + 5, 5)
                ml.setZValue(25)
                self.month_label_items.append([x, x + self.day_width * 31, ml])
                last_m = cm

        if self.month_label_items:
            self.month_label_items[-1][1] = tw_total

        # 横線
        for r in range(len(self.visible_tasks_info) + 1):
            y = r * self.row_height
            self.cs.addLine(0, y, tw_total, y, QPen(QColor(220, 220, 220), 1)).setZValue(-15)
        
        self.hs.addRect(0, 0, tw_total, 35, QPen(Qt.NoPen), QBrush(QColor(235, 245, 255))).setZValue(0)
        
        # 今日の線
        nx = (datetime.now() - self.min_date).total_seconds() / (24*3600) * self.day_width
        if 0 <= nx < tw_total:
            self.cs.addLine(nx, 0, nx, ch, QPen(QColor(255, 60, 60), 2, Qt.DashLine)).setZValue(25)
            
        for row, info in enumerate(self.visible_tasks_info):
            t = info['task']
            try:
                if t.get('is_group'):
                    gs, ge = self.get_group_range(info['index'])
                    if gs and ge:
                        group_item = GanttGroupItem(t, row, gs, ge, self)
                        self.cs.addItem(group_item)
                    continue

                periods = t.get('periods')
                if periods is None:
                    if t.get('start_date') and t.get('end_date'):
                        periods = [{'start_date': t['start_date'], 'end_date': t['end_date']}]
                    else:
                        continue
                
                if not periods:
                    continue
                    
                for p_idx, p in enumerate(periods):
                    if not p.get('start_date') or not p.get('end_date'): continue
                    sd = datetime.strptime(p['start_date'], "%Y-%m-%d")
                    ed = datetime.strptime(p['end_date'], "%Y-%m-%d")
                    bar_w = ((ed - sd).days + 1) * self.day_width
                    bar = GanttBarItem(t, row, p_idx, self, QRectF(0, 0, bar_w, self.row_height - 20))
                    bar.setPos((sd - self.min_date).days * self.day_width, row * self.row_height + 10)
                    bar.setZValue(30)
                    self.cs.addItem(bar)
            except Exception as e:
                print(f"Error drawing bar for row {row}: {e}")
                
        self.hs.setSceneRect(0, 0, tw_total, self.header_height)
        self.cs.setSceneRect(0, 0, tw_total, ch)
        self.update_month_labels_pos()

    def save_data(self):
        p = QFileDialog.getSaveFileName(self, "保存", "", "JSON (*.json)")[0]
        if p:
            try:
                data_to_save = {
                    "settings": {
                        "min_date": self.min_date.strftime("%Y-%m-%d"),
                        "display_unit": self.display_unit,
                        "display_count": self.display_count,
                        "zoom_unit": self.zoom_unit,
                        "zoom_count": self.zoom_count
                    },
                    "tasks": self.tasks
                }
                with open(p, 'w', encoding='utf-8') as f:
                    json.dump(data_to_save, f, ensure_ascii=False, indent=4)
                QMessageBox.information(self, "成功", "保存しました。")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"保存失敗: {e}")

    def load_data(self):
        p = QFileDialog.getOpenFileName(self, "開く", "", "JSON (*.json)")[0]
        if p:
            try:
                with open(p, 'r', encoding='utf-8') as f:
                    loaded_data = json.load(f)
                
                if isinstance(loaded_data, dict) and "tasks" in loaded_data:
                    self.tasks = loaded_data["tasks"]
                    settings = loaded_data.get("settings", {})
                    min_date_str = settings.get("min_date")
                    if min_date_str:
                        self.min_date = datetime.strptime(min_date_str, "%Y-%m-%d")
                    else:
                        self.min_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
                    
                    self.display_unit = settings.get("display_unit")
                    self.display_count = settings.get("display_count")
                    
                    if self.display_unit is None or self.display_count is None:
                        # 互換性処理
                        if "display_months" in settings:
                            self.display_unit = 1
                            self.display_count = settings["display_months"]
                        else:
                            days = settings.get("display_days", 150)
                            self.display_unit = 1
                            self.display_count = max(1, round(days / 30))
                    
                    self.zoom_unit = settings.get("zoom_unit", 1)
                    self.zoom_count = settings.get("zoom_count", 1)
                    
                    # 読込後の状態をUIに反映
                    self.zoom_unit_combo.blockSignals(True)
                    self.zoom_count_spin.blockSignals(True)
                    self.zoom_unit_combo.setCurrentIndex(self.zoom_unit)
                    self.zoom_count_spin.setValue(self.zoom_count)
                    self.zoom_unit_combo.blockSignals(False)
                    self.zoom_count_spin.blockSignals(False)
                    
                    self.calculate_day_width()
                    self.update_display_days()
                else:
                    self.tasks = loaded_data
                    self.min_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
                    self.display_unit = 1
                    self.display_count = 6
                    self.zoom_unit = 1
                    self.zoom_count = 1
                    self.calculate_day_width()
                    self.update_display_days()
                    
                self.update_ui()
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"読込失敗: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = GanttApp()
    window.showMaximized()
    sys.exit(app.exec())