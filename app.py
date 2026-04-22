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

class TaskDialog(QDialog):
    def __init__(self, parent=None, task=None):
        super().__init__(parent)
        self.setWindowTitle("タスクの編集" if task else "タスクの追加")
        self.resize(380, 240)
        self.setStyleSheet("""
            QDialog { background-color: #f7f7f7; color: #333333; }
            QLabel { color: #333333; font-size: 14px; }
            QLineEdit, QDateEdit, QSpinBox { 
                background-color: #ffffff; 
                color: #333333; 
                border: 1px solid #cccccc; 
                padding: 4px; 
                border-radius: 4px; 
            }
            QPushButton { 
                background-color: #0078d4; 
                color: white; 
                border-radius: 4px; 
                padding: 8px 16px; 
                font-weight: bold; 
            }
            QPushButton:hover { background-color: #0086f0; }
        """)
        self.layout = QFormLayout(self)
        self.name_input = QLineEdit()
        self.start_input = QDateEdit()
        self.start_input.setCalendarPopup(True)
        self.end_input = QDateEdit()
        self.end_input.setCalendarPopup(True)
        self.progress_input = QSpinBox()
        self.progress_input.setRange(0, 100)
        self.progress_input.setSuffix(" %")
        
        self.color_btn = QPushButton("色を選択")
        self.selected_color = "#0078d4"
        self.color_btn.clicked.connect(self.choose_color)
        
        if task:
            self.name_input.setText(task.get('name', ''))
            self.start_input.setDate(QDate.fromString(task.get('start_date', ''), Qt.ISODate))
            self.end_input.setDate(QDate.fromString(task.get('end_date', ''), Qt.ISODate))
            self.progress_input.setValue(task.get('progress', 0))
            self.selected_color = task.get('color', '#0078d4')
        else:
            self.start_input.setDate(QDate.currentDate())
            self.end_input.setDate(QDate.currentDate().addDays(7))
            
        self.update_color_btn_style()
            
        self.layout.addRow("タスク名:", self.name_input)
        self.layout.addRow("開始日:", self.start_input)
        self.layout.addRow("終了日:", self.end_input)
        self.layout.addRow("進捗率:", self.progress_input)
        self.layout.addRow("バーの色:", self.color_btn)
        
        btn_layout = QHBoxLayout()
        self.ok_btn = QPushButton("OK")
        self.cancel_btn = QPushButton("キャンセル")
        self.ok_btn.clicked.connect(self.accept)
        self.cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self.ok_btn)
        btn_layout.addWidget(self.cancel_btn)
        self.layout.addRow(btn_layout)

    def choose_color(self):
        color = QColorDialog.getColor(QColor(self.selected_color), self, "色を選択")
        if color.isValid():
            self.selected_color = color.name()
            self.update_color_btn_style()
            
    def update_color_btn_style(self):
        self.color_btn.setStyleSheet(f"background-color: {self.selected_color}; color: white; border-radius: 4px; padding: 6px;")

    def get_data(self):
        return {
            "name": self.name_input.text(),
            "start_date": self.start_input.date().toString(Qt.ISODate),
            "end_date": self.end_input.date().toString(Qt.ISODate),
            "progress": self.progress_input.value(),
            "color": self.selected_color
        }

class GanttBarItem(QGraphicsRectItem):
    def __init__(self, task, row, gantt_app, rect=None):
        super().__init__(rect)
        self.task = task
        self.row = row
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
        p_rect = QRectF(self.rect().left(), self.rect().top(), self.rect().width() * (prog / 100.0), self.rect().height())
        self.progress_item.setRect(p_rect)
        self.progress_item.setBrush(QBrush(bc))
        self.progress_item.setPen(Qt.NoPen)
        self.text_item.setPos(5, (self.rect().height() - self.text_item.boundingRect().height()) / 2)
        self.setToolTip(f"タスク: {self.task.get('name','')}\n期間: {self.task['start_date']}〜{self.task['end_date']}")

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
        self.task['start_date'] = sd.strftime("%Y-%m-%d")
        self.task['end_date'] = ed.strftime("%Y-%m-%d")
        self.app.sync_table_from_tasks()

    def mouseDoubleClickEvent(self, event):
        self.app.edit_task(self.row)
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
        self.btn_edit = QPushButton("編集")
        self.btn_del = QPushButton("削除")
        self.btn_add.clicked.connect(self.add_task)
        self.btn_edit.clicked.connect(self.edit_task)
        self.btn_del.clicked.connect(self.delete_task)
        tl.addWidget(self.btn_add)
        tl.addWidget(self.btn_edit)
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
        self.table.setHorizontalHeaderLabels(["タスク名", "進捗", "開始日", "終了日"])
        self.table.horizontalHeader().setFixedHeight(self.header_height)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.verticalHeader().setDefaultSectionSize(self.row_height)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.cellDoubleClicked.connect(lambda r, c: self.edit_task(r))
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

    def add_task(self):
        dlg = TaskDialog(self)
        if dlg.exec():
            self.tasks.append(dlg.get_data())
            self.update_ui()

    def edit_task(self, r=None):
        if r is None:
            r = self.table.currentRow()
        if 0 <= r < len(self.tasks):
            dlg = TaskDialog(self, self.tasks[r])
            if dlg.exec():
                self.tasks[r] = dlg.get_data()
                self.update_ui()

    def delete_task(self):
        r = self.table.currentRow()
        if r >= 0:
            if QMessageBox.question(self, "確認", "削除しますか？") == QMessageBox.Yes:
                self.tasks.pop(r)
                self.update_ui()

    def sync_table_from_tasks(self):
        for r, t in enumerate(self.tasks):
            s = datetime.strptime(t['start_date'], "%Y-%m-%d")
            e = datetime.strptime(t['end_date'], "%Y-%m-%d")
            self.table.item(r, 2).setText(s.strftime("%m/%d"))
            self.table.item(r, 3).setText(e.strftime("%m/%d"))

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
            "start_date": sd.strftime("%Y-%m-%d"), 
            "end_date": ed.strftime("%Y-%m-%d"), 
            "progress": 0, 
            "color": "#0078d4"
        }
        if row < len(self.tasks):
            self.tasks.insert(row, t)
        else:
            self.tasks.append(t)
        self.update_ui()

    def update_ui(self):
        self.table.setRowCount(len(self.tasks))
        for r, t in enumerate(self.tasks):
            self.table.setItem(r, 0, QTableWidgetItem(t['name']))
            p_item = QTableWidgetItem(f"{t['progress']}%")
            p_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(r, 1, p_item)
            s = datetime.strptime(t['start_date'], "%Y-%m-%d")
            e = datetime.strptime(t['end_date'], "%Y-%m-%d")
            self.table.setItem(r, 2, QTableWidgetItem(s.strftime("%m/%d")))
            self.table.setItem(r, 3, QTableWidgetItem(e.strftime("%m/%d")))
        self.draw_chart()

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
                sd = datetime.strptime(t['start_date'], "%Y-%m-%d")
                ed = datetime.strptime(t['end_date'], "%Y-%m-%d")
                bar_w = ((ed - sd).days + 1) * self.day_width
                bar = GanttBarItem(t, row, self, QRectF(0, 0, bar_w, self.row_height - 20))
                bar.setPos((sd - self.min_date).days * self.day_width, row * self.row_height + 10)
                bar.setZValue(30)
                self.cs.addItem(bar)
            except Exception as e:
                print(f"Error drawing bar: {e}")
                
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