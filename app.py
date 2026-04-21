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
            QDialog { background-color: #2b2b2b; color: #ffffff; }
            QLabel { color: #ffffff; font-size: 14px; }
            QLineEdit, QDateEdit, QSpinBox { 
                background-color: #3b3b3b; 
                color: #ffffff; 
                border: 1px solid #555555; 
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
        self.color_btn.setStyleSheet(f"background-color: {self.selected_color}; color: white; font-weight: bold; border-radius: 4px; padding: 6px;")

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
        
        # プログレス（子アイテム）
        self.progress_item = QGraphicsRectItem(self)
        
        # テキスト
        self.text_item = QGraphicsTextItem(task.get('name', ''), self)
        self.text_item.setDefaultTextColor(Qt.white)
        font = QFont("Segoe UI", 9, QFont.Bold)
        self.text_item.setFont(font)
        
        self.resizing_left = False
        self.resizing_right = False
        self.edge_margin = 8
        self.update_appearance()

    def update_tooltip(self):
        msg = f"タスク名: {self.task.get('name', '')}\n期間: {self.task.get('start_date')} 〜 {self.task.get('end_date')}\n進捗: {self.task.get('progress', 0)}%"
        self.setToolTip(msg)

    def update_appearance(self):
        self.update_tooltip()
        base_color = QColor(self.task.get('color', '#0078d4'))
        
        if self.isSelected():
            self.setPen(QPen(Qt.white, 2))
        else:
            self.setPen(QPen(base_color.darker(150), 1))
            
        # 背景（全体の枠）は少し暗めの色にして未完了部分とする
        self.setBrush(QBrush(base_color.darker(150)))
        
        # プログレス部分の描画
        progress = self.task.get('progress', 0)
        p_width = self.rect().width() * (progress / 100.0)
        p_rect = QRectF(self.rect().left(), self.rect().top(), p_width, self.rect().height())
        self.progress_item.setRect(p_rect)
        self.progress_item.setBrush(QBrush(base_color))
        self.progress_item.setPen(Qt.NoPen)
        
        # テキスト表示位置の調整
        self.text_item.setPos(5, (self.rect().height() - self.text_item.boundingRect().height()) / 2)

    def hoverMoveEvent(self, event):
        pos = event.pos().x()
        if pos < self.edge_margin:
            self.setCursor(Qt.SizeHorCursor)
        elif pos > self.rect().width() - self.edge_margin:
            self.setCursor(Qt.SizeHorCursor)
        else:
            self.setCursor(Qt.ArrowCursor)
        super().hoverMoveEvent(event)

    def mousePressEvent(self, event):
        pos = event.pos().x()
        if pos < self.edge_margin:
            self.resizing_left = True
        elif pos > self.rect().width() - self.edge_margin:
            self.resizing_right = True
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.resizing_left:
            diff = event.pos().x() - event.lastPos().x()
            new_rect = self.rect()
            new_rect.setLeft(new_rect.left() + diff)
            if new_rect.width() > self.app.day_width:
                self.setRect(new_rect)
                self.setPos(self.pos().x() + diff, self.pos().y())
        elif self.resizing_right:
            diff = event.pos().x() - event.lastPos().x()
            new_rect = self.rect()
            new_rect.setRight(new_rect.right() + diff)
            if new_rect.width() > self.app.day_width:
                self.setRect(new_rect)
        else:
            last_pos = self.pos()
            super().mouseMoveEvent(event)
            self.setPos(self.pos().x(), last_pos.y())
        self.update_appearance()

    def mouseReleaseEvent(self, event):
        self.resizing_left = False
        self.resizing_right = False
        super().mouseReleaseEvent(event)
        self.update_task_dates()

    def mouseDoubleClickEvent(self, event):
        self.app.edit_task(self.row)
        super().mouseDoubleClickEvent(event)

    def update_task_dates(self):
        new_x = self.pos().x() + self.rect().left()
        new_width = self.rect().width()
        days_from_min = round(new_x / self.app.day_width)
        duration_days = round(new_width / self.app.day_width) - 1
        start_date = self.app.min_date + timedelta(days=days_from_min)
        end_date = start_date + timedelta(days=duration_days)
        self.task['start_date'] = start_date.strftime("%Y-%m-%d")
        self.task['end_date'] = end_date.strftime("%Y-%m-%d")
        self.app.sync_table_from_tasks()

class GanttScene(QGraphicsScene):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.start_x = 0

    def mousePressEvent(self, event):
        item = self.itemAt(event.scenePos(), self.app.view.transform())
        if not item and event.button() == Qt.LeftButton:
            self.start_x = event.scenePos().x()
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if self.start_x > 0 and not self.itemAt(event.scenePos(), self.app.view.transform()):
            end_x = event.scenePos().x()
            if abs(end_x - self.start_x) > self.app.day_width / 2:
                self.create_task_from_drag(self.start_x, end_x, event.scenePos().y())
            self.start_x = 0
        super().mouseReleaseEvent(event)

    def create_task_from_drag(self, x1, x2, y):
        start_x = min(x1, x2)
        end_x = max(x1, x2)
        days_from_min = round(start_x / self.app.day_width)
        duration_days = max(1, round((end_x - start_x) / self.app.day_width))
        start_date = self.app.min_date + timedelta(days=days_from_min)
        end_date = start_date + timedelta(days=duration_days - 1)
        row = int((y - self.app.header_height) / self.app.row_height)
        if row < 0: row = 0
        task = {
            "name": f"新規タスク {len(self.app.tasks) + 1}",
            "start_date": start_date.strftime("%Y-%m-%d"),
            "end_date": end_date.strftime("%Y-%m-%d"),
            "progress": 0,
            "color": "#0078d4" # デフォルト色
        }
        if row < len(self.app.tasks):
            self.app.tasks.insert(row, task)
        else:
            self.app.tasks.append(task)
        self.app.update_ui()

    def contextMenuEvent(self, event):
        item = self.itemAt(event.scenePos(), self.app.view.transform())
        if isinstance(item, GanttBarItem):
            menu = QMenu()
            edit_action = menu.addAction("このタスクを編集")
            del_action = menu.addAction("このタスクを削除")
            action = menu.exec(event.screenPos())
            if action == edit_action:
                self.app.edit_task(item.row)
            elif action == del_action:
                self.app.tasks.pop(item.row)
                self.app.update_ui()
        super().contextMenuEvent(event)

class GanttApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MiyaGantt - Professional Gantt Chart")
        self.resize(1380, 850)
        self.tasks = []
        self.day_width = 40
        self.row_height = 40
        self.header_height = 60
        self.min_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=7)
        self.init_ui()
        self.apply_styles()

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #1e1e1e; }
            QWidget { background-color: #1e1e1e; color: #d4d4d4; }
            QTableWidget { 
                background-color: #252526; 
                gridline-color: #3f3f3f; 
                border: none;
                color: #cccccc;
            }
            QHeaderView::section { 
                background-color: #333333; 
                color: #ffffff; 
                border: 1px solid #1e1e1e;
                padding: 4px;
            }
            QPushButton { 
                background-color: #333333; 
                border: 1px solid #454545; 
                padding: 6px 12px;
                border-radius: 4px;
                color: #ffffff;
            }
            QPushButton:hover { background-color: #454545; }
            QSlider::handle:horizontal {
                background: #0078d4;
                width: 18px;
                margin: -5px 0;
                border-radius: 9px;
            }
            QSlider::groove:horizontal {
                border: 1px solid #444;
                height: 8px;
                background: #333;
                border-radius: 4px;
            }
        """)

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        toolbar = QHBoxLayout()
        self.btn_add = QPushButton("追加")
        self.btn_edit = QPushButton("編集")
        self.btn_del = QPushButton("削除")
        self.btn_load = QPushButton("読込")
        self.btn_save = QPushButton("保存")
        toolbar.addWidget(self.btn_add)
        toolbar.addWidget(self.btn_edit)
        toolbar.addWidget(self.btn_del)
        toolbar.addStretch()
        toolbar.addWidget(QLabel("ズーム:"))
        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setRange(20, 100)
        self.zoom_slider.setValue(self.day_width)
        self.zoom_slider.setFixedWidth(150)
        self.zoom_slider.valueChanged.connect(self.change_zoom)
        toolbar.addWidget(self.zoom_slider)
        toolbar.addWidget(self.btn_load)
        toolbar.addWidget(self.btn_save)
        
        self.btn_add.clicked.connect(lambda: self.add_task())
        self.btn_edit.clicked.connect(lambda: self.edit_task())
        self.btn_del.clicked.connect(self.delete_task)
        self.btn_save.clicked.connect(self.save_data)
        self.btn_load.clicked.connect(self.load_data)
        
        main_layout.addLayout(toolbar)
        self.splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(self.splitter)
        
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["タスク名", "進捗", "開始日", "終了日"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.verticalHeader().setDefaultSectionSize(self.row_height)
        self.table.verticalHeader().setVisible(False)
        self.table.cellDoubleClicked.connect(lambda row, col: self.edit_task(row))
        self.splitter.addWidget(self.table)
        
        self.scene = GanttScene(self)
        self.view = QGraphicsView(self.scene)
        self.view.setRenderHint(QPainter.Antialiasing)
        self.view.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.splitter.addWidget(self.view)
        
        self.splitter.setSizes([450, 900])
        self.table.verticalScrollBar().valueChanged.connect(self.view.verticalScrollBar().setValue)
        self.view.verticalScrollBar().valueChanged.connect(self.table.verticalScrollBar().setValue)

    def change_zoom(self, value):
        self.day_width = value
        self.update_ui()

    def add_task(self):
        dlg = TaskDialog(self)
        if dlg.exec():
            self.tasks.append(dlg.get_data())
            self.update_ui()
            
    def edit_task(self, row=None):
        if row is None:
            row = self.table.currentRow()
        if row < 0 or row >= len(self.tasks): return
        
        dlg = TaskDialog(self, self.tasks[row])
        if dlg.exec():
            self.tasks[row] = dlg.get_data()
            self.update_ui()

    def delete_task(self):
        row = self.table.currentRow()
        if row < 0: return
        if QMessageBox.question(self, "確認", "選択したタスクを削除しますか？") == QMessageBox.Yes:
            self.tasks.pop(row)
            self.update_ui()

    def sync_table_from_tasks(self):
        for row, task in enumerate(self.tasks):
            self.table.item(row, 2).setText(task.get('start_date', ''))
            self.table.item(row, 3).setText(task.get('end_date', ''))

    def update_ui(self):
        self.table.setRowCount(len(self.tasks))
        for row, task in enumerate(self.tasks):
            self.table.setItem(row, 0, QTableWidgetItem(task.get('name', '')))
            
            prog_item = QTableWidgetItem(f"{task.get('progress', 0)}%")
            prog_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 1, prog_item)
            
            self.table.setItem(row, 2, QTableWidgetItem(task.get('start_date', '')))
            self.table.setItem(row, 3, QTableWidgetItem(task.get('end_date', '')))
        self.draw_chart()

    def draw_chart(self):
        self.scene.clear()
        total_tasks = len(self.tasks)
        chart_height = max(self.view.height() - 20, (total_tasks + 2) * self.row_height + self.header_height)
        display_days = 150 # 広めに描画
        total_width = display_days * self.day_width
        
        for i in range(display_days):
            d = self.min_date + timedelta(days=i)
            x = i * self.day_width
            if d.weekday() >= 5:
                bg_color = QColor(40, 40, 40) if d.weekday() == 5 else QColor(50, 35, 35)
                self.scene.addRect(x, 0, self.day_width, chart_height, QPen(Qt.NoPen), QBrush(bg_color))
            self.scene.addLine(x, 0, x, chart_height, QPen(QColor(60, 60, 60), 1))
            header_rect = QRectF(x, 0, self.day_width, self.header_height)
            self.scene.addRect(header_rect, QPen(QColor(80, 80, 80)), QBrush(QColor(45, 45, 45)))
            
            date_text = self.scene.addText(d.strftime("%d"))
            date_text.setDefaultTextColor(Qt.white)
            date_text.setPos(x + 5, 25)
            if d.day == 1 or i == 0 or d.weekday() == 0:
                month_text = self.scene.addText(d.strftime("%Y/%m"))
                month_text.setDefaultTextColor(QColor(0, 160, 255))
                month_text.setFont(QFont("Segoe UI", 8, QFont.Bold))
                month_text.setPos(x + 2, 5)
                
        # 今日の線（Today Line）
        now_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        days_from_min = (now_date - self.min_date).days
        if 0 <= days_from_min < display_days:
            today_x = days_from_min * self.day_width + (self.day_width / 2)
            pen = QPen(QColor(255, 80, 80), 2, Qt.DashLine)
            today_line = self.scene.addLine(today_x, self.header_height, today_x, chart_height, pen)
            today_line.setZValue(5) # バーの下層
            
            today_label = self.scene.addText("Today")
            today_label.setDefaultTextColor(QColor(255, 80, 80))
            today_label.setFont(QFont("Segoe UI", 8, QFont.Bold))
            today_label.setPos(today_x - 15, self.header_height - 20)

        for row, task in enumerate(self.tasks):
            try:
                sd = datetime.strptime(task.get('start_date', ''), "%Y-%m-%d")
                ed = datetime.strptime(task.get('end_date', ''), "%Y-%m-%d")
            except:
                continue
            start_offset = (sd - self.min_date).days * self.day_width
            duration = (ed - sd).days * self.day_width + self.day_width
            y = self.header_height + row * self.row_height + 10
            bar = GanttBarItem(task, row, self, QRectF(0, 0, duration, self.row_height - 20))
            bar.setPos(start_offset, y)
            bar.setZValue(10)
            self.scene.addItem(bar)
            
        self.scene.setSceneRect(0, 0, total_width, chart_height)

    def save_data(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存", "", "JSON Files (*.json)")
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    json.dump(self.tasks, f, ensure_ascii=False, indent=4)
                QMessageBox.information(self, "成功", "保存しました。")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"保存に失敗しました:\n{e}")
                
    def load_data(self):
        path, _ = QFileDialog.getOpenFileName(self, "開く", "", "JSON Files (*.json)")
        if path:
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    self.tasks = json.load(f)
                self.update_ui()
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"読み込みに失敗しました:\n{e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = GanttApp()
    window.show()
    sys.exit(app.exec())