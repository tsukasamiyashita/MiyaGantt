import sys
import json
from datetime import datetime, timedelta
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, 
                               QHeaderView, QSplitter, QGraphicsView, QGraphicsScene, 
                               QDialog, QFormLayout, QLineEdit, QDateEdit, QMessageBox, 
                               QFileDialog, QGraphicsRectItem, QGraphicsTextItem, QSlider, QLabel, QMenu, QSpinBox, QColorDialog)
from PySide6.QtCore import Qt, QDate, QRectF, QPointF
from PySide6.QtGui import QBrush, QPen, QColor, QFont, QPainter

# TaskDialog was removed in favor of inline editing.

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
        self.text_item = QGraphicsTextItem(task.get('name', ''), self)
        self.text_item.setDefaultTextColor(Qt.white)
        self.text_item.setZValue(1)
        font = QFont("Segoe UI", 9, QFont.Bold)
        self.text_item.setFont(font)
        
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
        self.text_item.setPos(5, (self.rect().height() - self.text_item.boundingRect().height()) / 2)
        p_dict = self.task.get('periods', [self.task])[self.period_index] if 'periods' in self.task else self.task
        start_d = p_dict.get('start_date', '')
        end_d = p_dict.get('end_date', '')
        self.setToolTip(f"タスク: {self.task.get('name','')}\n期間: {start_d}〜{end_d}")

    def mouseMoveEvent(self, event):
        snap = self.app.day_width * 0.25
        if self.resizing_left:
            diff = event.pos().x() - event.lastPos().x()
            nr = self.rect()
            nr.setLeft(nr.left() + diff)
            if nr.width() > snap:
                self.setRect(nr)
                self.setPos(self.pos().x() + diff, self.pos().y())
        elif self.resizing_right:
            diff = event.pos().x() - event.lastPos().x()
            nr = self.rect()
            nr.setRight(nr.right() + diff)
            if nr.width() > snap:
                self.setRect(nr)
        else:
            ly = self.pos().y()
            super().mouseMoveEvent(event)
            self.setPos(self.pos().x(), ly)
        self.update_appearance()

    def mouseReleaseEvent(self, event):
        self.resizing_left = self.resizing_right = False
        snap = self.app.day_width * 0.25
        sx = round((self.pos().x() + self.rect().left()) / snap) * snap
        sw = max(snap, round(self.rect().width() / snap) * snap)
        self.setPos(sx, self.pos().y())
        self.setRect(0, 0, sw, self.rect().height())
        super().mouseReleaseEvent(event)
        
        sd = self.app.min_date + timedelta(days=sx / self.app.day_width)
        ed = sd + timedelta(days=sw / self.app.day_width - 0.001)
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

class GanttApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MiyaGantt - Professional Gantt Chart")
        self.resize(1380, 850)
        self.tasks = []
        self.day_width = 80
        self.row_height = 40
        self.header_height = 70
        self.min_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=14)
        self.month_label_items = []
        
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
        self.btn_del = QPushButton("削除")
        self.btn_add.clicked.connect(self.add_task)
        self.btn_del.clicked.connect(self.delete_task)
        tl.addWidget(self.btn_add)
        tl.addWidget(self.btn_del)
        tl.addStretch()
        
        tl.addWidget(QLabel("ズーム:"))
        self.zs = QSlider(Qt.Horizontal)
        self.zs.setRange(40, 400); self.zs.setValue(self.day_width); self.zs.setFixedWidth(150)
        self.zs.valueChanged.connect(self.change_zoom)
        tl.addWidget(self.zs)
        
        self.btn_load = QPushButton("読込")
        self.btn_save = QPushButton("保存")
        self.btn_load.clicked.connect(self.load_data)
        self.btn_save.clicked.connect(self.save_data)
        tl.addWidget(self.btn_load)
        tl.addWidget(self.btn_save)
        ml.addLayout(tl)
        
        self.splitter = QSplitter(Qt.Horizontal)
        ml.addWidget(self.splitter)
        
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["タスク名", "進捗(%)", "期間指定", "色"])
        self.table.horizontalHeader().setFixedHeight(self.header_height)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.verticalHeader().setDefaultSectionSize(self.row_height)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        
        self.table.itemChanged.connect(self.on_table_item_changed)
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
        view_left = self.hv.mapToScene(0, 0).x()
        for start_x, end_x, item in self.month_label_items:
            tw = item.boundingRect().width()
            new_x = max(start_x + 5, min(view_left + 5, end_x - tw - 5))
            item.setPos(new_x, item.pos().y())

    def change_zoom(self, v):
        self.day_width = v
        self.update_ui()

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

    def add_task(self):
        today = datetime.now()
        next_week = today + timedelta(days=7)
        t = {
            "name": f"新規タスク {len(self.tasks)+1}",
            "periods": [{"start_date": today.strftime("%Y-%m-%d"), "end_date": next_week.strftime("%Y-%m-%d")}],
            "progress": 0,
            "color": "#0078d4"
        }
        self.tasks.append(t)
        self.update_ui()
        self.table.editItem(self.table.item(len(self.tasks)-1, 0))

    def delete_task(self):
        r = self.table.currentRow()
        if r >= 0:
            if QMessageBox.question(self, "確認", "削除しますか？") == QMessageBox.Yes:
                self.tasks.pop(r)
                self.update_ui()

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
        for r, t in enumerate(self.tasks):
            periods = t.get('periods', [])
            p_strs = []
            for p in periods:
                if not p.get('start_date') or not p.get('end_date'): continue
                s = p['start_date'].replace('-', '/')
                e = p['end_date'].replace('-', '/')
                p_strs.append(f"{s}-{e}")
            
            period_item = self.table.item(r, 2)
            if period_item:
                period_item.setText(", ".join(p_strs))
        self.table.blockSignals(False)

    def create_task_from_drag(self, x1, x2, y):
        snap = self.day_width * 0.25
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
        if row < len(self.tasks):
            self.tasks.insert(row, t)
        else:
            self.tasks.append(t)
        self.update_ui()

    def update_ui(self):
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.tasks))
        for r, t in enumerate(self.tasks):
            item_name = QTableWidgetItem(t.get('name', ''))
            self.table.setItem(r, 0, item_name)
            
            item_prog = QTableWidgetItem(str(t.get('progress', 0)))
            item_prog.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(r, 1, item_prog)
            
            periods = t.get('periods', [])
            p_strs = []
            for p in periods:
                if not p.get('start_date') or not p.get('end_date'): continue
                s = p['start_date'].replace('-', '/')
                e = p['end_date'].replace('-', '/')
                p_strs.append(f"{s}-{e}")
                
            item_period = QTableWidgetItem(", ".join(p_strs))
            self.table.setItem(r, 2, item_period)
            
            item_color = QTableWidgetItem("")
            item_color.setBackground(QColor(t.get('color', '#0078d4')))
            item_color.setFlags(item_color.flags() & ~Qt.ItemIsEditable) 
            self.table.setItem(r, 3, item_color)
            
        self.table.blockSignals(False)
        self.draw_chart()

    def on_table_item_changed(self, item):
        row = item.row()
        col = item.column()
        if row >= len(self.tasks) or row < 0: return
        t = self.tasks[row]
        
        if col == 0:
            t['name'] = item.text()
        elif col == 1:
            try:
                prog = int(item.text().replace('%', '').strip())
                t['progress'] = max(0, min(100, prog))
            except ValueError:
                pass
            self.table.blockSignals(True)
            item.setText(str(t['progress']))
            self.table.blockSignals(False)
        elif col == 2:
            period_str = item.text()
            parsed = self.get_periods_from_string(period_str)
            if parsed:
                t['periods'] = parsed
            else:
                QMessageBox.warning(self, "エラー", "期間の形式が正しくありません。\n例: 04/01-04/05")
        
        self.draw_chart()

    def on_table_cell_double_clicked(self, row, col):
        if col == 3: # Color column
            t = self.tasks[row]
            color = QColorDialog.getColor(QColor(t.get('color', '#0078d4')), self, "色を選択")
            if color.isValid():
                t['color'] = color.name()
                self.update_ui()

    def draw_chart(self):
        self.hs.clear()
        self.cs.clear()
        self.month_label_items = []
        tw_total = 150 * self.day_width
        ch = max(self.chart_view.height(), (len(self.tasks) + 10) * self.row_height)
        
        last_m = None
        for i in range(150):
            d = self.min_date + timedelta(days=i)
            x = i * self.day_width
            
            # 背景
            if d.weekday() >= 5:
                bg = QColor(240, 248, 255) if d.weekday()==5 else QColor(255, 240, 240)
                re = self.cs.addRect(x, 0, self.day_width, ch, QPen(Qt.NoPen), QBrush(bg))
                re.setZValue(-20)
            
            # グリッド
            self.cs.addLine(x, 0, x, ch, QPen(QColor(220, 220, 220), 1)).setZValue(-15)
            for h in [6, 12, 18]:
                self.cs.addLine(x + (self.day_width * h / 24.0), 0, x + (self.day_width * h / 24.0), ch, QPen(QColor(245, 245, 245), 0.5)).setZValue(-15)
            
            # ヘッダー (日付・曜日)
            self.hs.addRect(x, 35, self.day_width, 35, QPen(QColor(200, 200, 200)), QBrush(QColor(245, 245, 245))).setZValue(5)
            dl = self.hs.addText(d.strftime("%d"))
            dl.setDefaultTextColor(QColor(50, 50, 50))
            dl.setFont(QFont("Segoe UI", 10, QFont.Bold))
            dl.setPos(x + (self.day_width/2) - 13, 35)
            dl.setZValue(10)
            
            w_c = QColor(0, 80, 200) if d.weekday()==5 else QColor(220, 0, 0) if d.weekday()==6 else QColor(60, 60, 60)
            yl = self.hs.addText(["月","火","水","木","金","土","日"][d.weekday()])
            yl.setDefaultTextColor(w_c)
            yl.setFont(QFont("Segoe UI", 8))
            yl.setPos(x + (self.day_width/2) - 10, 50)
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
        for r in range(len(self.tasks) + 10):
            y = r * self.row_height
            self.cs.addLine(0, y, tw_total, y, QPen(QColor(220, 220, 220), 1)).setZValue(-15)
        
        self.hs.addRect(0, 0, tw_total, 35, QPen(Qt.NoPen), QBrush(QColor(235, 245, 255))).setZValue(0)
        
        # 今日の線
        nx = (datetime.now() - self.min_date).total_seconds() / (24*3600) * self.day_width
        if 0 <= nx < tw_total:
            self.cs.addLine(nx, 0, nx, ch, QPen(QColor(255, 60, 60), 2, Qt.DashLine)).setZValue(25)
            
        for row, t in enumerate(self.tasks):
            try:
                periods = t.get('periods', [])
                if not periods and 'start_date' in t:
                    periods = [{'start_date': t['start_date'], 'end_date': t['end_date']}]
                    
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
                print(f"Error drawing bar for task {row}: {e}")
                
        self.hs.setSceneRect(0, 0, tw_total, self.header_height)
        self.cs.setSceneRect(0, 0, tw_total, ch)
        self.update_month_labels_pos()

    def save_data(self):
        p = QFileDialog.getSaveFileName(self, "保存", "", "JSON (*.json)")[0]
        if p:
            try:
                with open(p, 'w', encoding='utf-8') as f:
                    json.dump(self.tasks, f, ensure_ascii=False, indent=4)
                QMessageBox.information(self, "成功", "保存しました。")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"保存失敗: {e}")

    def load_data(self):
        p = QFileDialog.getOpenFileName(self, "開く", "", "JSON (*.json)")[0]
        if p:
            try:
                with open(p, 'r', encoding='utf-8') as f:
                    self.tasks = json.load(f)
                self.update_ui()
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"読込失敗: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = GanttApp()
    window.show()
    sys.exit(app.exec())