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
                            c_code = p.get('color', t.get('color', '#0078d4'))
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
        try:
            periods = self.task.get('periods', [])
            if self.period_index >= len(periods) and self.period_index != 0:
                return
            p_dict = periods[self.period_index] if self.period_index < len(periods) else self.task
            color_code = p_dict.get('color', self.task.get('color', '#0078d4'))
            bc = QColor(color_code)
            
            if self.isSelected():
                # 選択時は太いオレンジ色の枠線で強調
                self.setPen(QPen(QColor("#ff8c00"), 3))
                self.setBrush(QBrush(bc.lighter(170)))
                self.setZValue(40) # 選択中のアイテムを最前面に
            else:
                self.setPen(QPen(bc.darker(120), 1))
                self.setBrush(QBrush(bc.lighter(150)))
                self.setZValue(30)
            
            prog = self.task.get('progress', 0)
        
            p_data = self.task.get('periods', [self.task])
            valid_periods = []
            for i, p in enumerate(p_data):
                try:
                    sd = datetime.strptime(p.get('start_date', ''), "%Y-%m-%d")
                    ed = datetime.strptime(ed_str := p.get('end_date', ''), "%Y-%m-%d")
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
        except Exception:
            pass

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

    def itemChange(self, change, value):
        # 選択状態が変わった時に見た目を更新する
        if change == QGraphicsRectItem.ItemSelectedChange:
            QTimer.singleShot(0, self.update_appearance)
        return super().itemChange(change, value)

    def mousePressEvent(self, event):
        if event.button() == Qt.RightButton:
            if self.isSelected():
                event.accept()
                return
            else:
                if self.scene():
                    self.scene().clearSelection()
                self.setSelected(True)
                event.accept()
                return
        
        if event.button() == Qt.LeftButton:
            # リサイズハンドルの判定
            x = event.pos().x()
            w = self.rect().width()
            margin = 12 if w <= self.app.day_width else 10
            margin = min(margin, w / 2 - 2)
            
            if x < margin:
                self.app.save_state()
                self.resizing_left = True
            elif x > w - margin:
                self.app.save_state()
                self.resizing_right = True
            else:
                self.setCursor(Qt.ClosedHandCursor)
            
            # テーブルの行選択を同期
            self.app.table.setCurrentCell(self.row, 2)
        
        super().mousePressEvent(event)
        self.update_appearance()

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
            # 複数選択されている場合は、全ての選択アイテムを一緒に移動させる
            selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttBarItem)]
            if len(selected_items) > 1 and self.isSelected():
                delta = event.scenePos() - event.lastScenePos()
                for it in selected_items:
                    it.setPos(it.pos() + delta)
                    # 垂直方向のスナップ（各アイテムごとの行へ）
                    it_row = int(it.scenePos().y() / self.app.row_height)
                    it.setPos(it.pos().x(), it_row * self.app.row_height + 10)
                    it.update_appearance()
            else:
                super().mouseMoveEvent(event)
                # マウスカーソルの位置に基づいて行の中心にスナップ
                row = int(event.scenePos().y() / self.app.row_height)
                max_row = len(self.app.visible_tasks_info) - 1 if self.app.visible_tasks_info else 0
                row = max(0, min(max_row, row))
                self.setPos(self.pos().x(), row * self.app.row_height + 10)
        self.update_appearance()

    def mouseReleaseEvent(self, event):
        was_resizing = self.resizing_left or self.resizing_right
        selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttBarItem)]
        
        # 複数移動の確定処理
        if not was_resizing and len(selected_items) > 1 and self.isSelected():
            self.app.save_state()
            
            # 各アイテムの状態を確定させる
            # インデックスが変わるのを防ぐため、移動対象を整理
            move_targets = []
            for it in selected_items:
                snap = self.app.day_width
                sx = round(it.pos().x() / snap) * snap
                sw = max(snap, round(it.rect().width() / snap) * snap)
                sd = self.app.min_date + timedelta(days=sx / self.app.day_width)
                ed = sd + timedelta(days=sw / self.app.day_width - 0.001)
                
                new_row = int(it.scenePos().y() / self.app.row_height)
                max_row = len(self.app.visible_tasks_info) - 1 if self.app.visible_tasks_info else 0
                new_row = max(0, min(max_row, new_row))
                
                move_targets.append({
                    'item': it,
                    'new_row': new_row,
                    'start_date': sd.strftime("%Y-%m-%d"),
                    'end_date': ed.strftime("%Y-%m-%d")
                })
            
            # データの反映（一度全ての期間を取り出してから、新しい場所に挿入する）
            # これにより、移動中にインデックスが狂うのを防ぐ
            # タスクごとにインデックスを管理して、逆順で削除を行う
            tasks_to_delete = {} # id(task) -> (task, set(indices_to_delete))
            to_insert = [] # [(target_task, p_data)]

            try:
                self.app.chart_view.setUpdatesEnabled(False)
                for target in move_targets:
                    it = target['item']
                    if 'periods' in it.task and it.period_index < len(it.task['periods']):
                        t_id = id(it.task)
                        if t_id not in tasks_to_delete:
                            tasks_to_delete[t_id] = (it.task, set())
                        
                        # 同一期間の二重処理を防ぐ
                        if it.period_index not in tasks_to_delete[t_id][1]:
                            tasks_to_delete[t_id][1].add(it.period_index)
                            
                            # 移動データの作成
                            p = it.task['periods'][it.period_index].copy()
                            p['start_date'] = target['start_date']
                            p['end_date'] = target['end_date']
                            
                            target_task = self.app.visible_tasks_info[target['new_row']]['task']
                            to_insert.append((target_task, p))
                
                # 1. 削除処理（リスト再構築方式）
                for task, indices in tasks_to_delete.values():
                    if 'periods' in task:
                        task['periods'] = [p for i, p in enumerate(task['periods']) if i not in indices]
                
                # 2. 挿入処理
                self.app.pending_selection = []
                for target_task, p_data in to_insert:
                    if target_task.get('is_group'):
                        continue
                    if 'periods' not in target_task:
                        target_task['periods'] = []
                    target_task['periods'].append(p_data)
                    # 再選択のために、新しく作成されたデータのIDを記憶
                    self.app.pending_selection.append(id(p_data))

                QTimer.singleShot(100, self.app.update_ui)
            except Exception as e:
                print(f"Error in multi-move: {e}")
            finally:
                self.app.chart_view.setUpdatesEnabled(True)
            
            super().mouseReleaseEvent(event)
            return

        # 単一移動の確定処理（既存）
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
                QTimer.singleShot(100, self.app.update_ui)
                super().mouseReleaseEvent(event)
                return

            # 移動元・移動先の両方で 'periods' 形式を確定させる
            for t in [self.task, target_task]:
                if 'periods' not in t:
                    t['periods'] = [{'start_date': t.get('start_date', ''), 'end_date': t.get('end_date', '')}]
            
            if 0 <= self.period_index < len(self.task['periods']):
                # 期間データを移動
                self.app.save_state()
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
                
                if self.scene():
                    self.scene().clearSelection()
                QTimer.singleShot(100, self.app.update_ui)
            
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
        QTimer.singleShot(100, self.app.update_ui)

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
                self.app.save_state()
                p_dict['text'] = text
                self.update_appearance()
                QTimer.singleShot(100, self.app.update_ui) # 全体を再描画して確実に反映させる

    def contextMenuEvent(self, event):
        selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttBarItem)]
        if self not in selected_items:
            selected_items = [self]
            
        menu = QMenu()
        copy_action = menu.addAction(f"コピー ({len(selected_items)})" if len(selected_items) > 1 else "コピー")
        cut_action = menu.addAction(f"切り取り ({len(selected_items)})" if len(selected_items) > 1 else "切り取り")
        menu.addSeparator()
        color_action = menu.addAction("色を変更")
        del_action = menu.addAction("この期間を削除")
        action = menu.exec(event.screenPos())
        
        if action == copy_action:
            if not selected_items: return
            try:
                temp_clipboard = []
                # 生きているアイテムのみを抽出
                valid_items = []
                for it in selected_items:
                    if shiboken6.isValid(it):
                        valid_items.append(it)
                
                if not valid_items: return
                
                # 基準となる座標を特定
                min_x = min(it.pos().x() for it in valid_items)
                min_row = min(it.row for it in valid_items)
                
                for it in valid_items:
                    try:
                        if shiboken6.isValid(it) and 'periods' in it.task and it.period_index < len(it.task['periods']):
                            p_data = it.task['periods'][it.period_index].copy()
                            temp_clipboard.append({
                                'data': p_data,
                                'rel_row': it.row - min_row,
                                'rel_days': (it.pos().x() - min_x) / self.app.day_width
                            })
                    except (RuntimeError, AttributeError, IndexError):
                        continue
                
                if temp_clipboard:
                    self.app.clipboard_periods = temp_clipboard
                    print(f"Copied {len(temp_clipboard)} items.")
            except Exception as e:
                print(f"Error in copy action: {e}")
        elif action == cut_action:
            if not selected_items: return
            try:
                self.app.chart_view.setUpdatesEnabled(False)
                # 生きているアイテムのみを抽出
                valid_items = []
                for it in selected_items:
                    if shiboken6.isValid(it):
                        valid_items.append(it)
                
                if not valid_items: return

                temp_clipboard = []
                # 基準座標の計算
                min_x = min(it.pos().x() for it in valid_items)
                min_row = min(it.row for it in valid_items)
                
                tasks_to_delete = {} # id(task) -> (task, set(indices))
                for it in valid_items:
                    try:
                        if shiboken6.isValid(it) and 'periods' in it.task:
                            if it.period_index < len(it.task['periods']):
                                p_data = it.task['periods'][it.period_index].copy()
                                temp_clipboard.append({
                                    'data': p_data,
                                    'rel_row': it.row - min_row,
                                    'rel_days': (it.pos().x() - min_x) / self.app.day_width
                                })
                                
                                t_id = id(it.task)
                                if t_id not in tasks_to_delete:
                                    tasks_to_delete[t_id] = (it.task, set())
                                tasks_to_delete[t_id][1].add(it.period_index)
                    except (RuntimeError, AttributeError, IndexError):
                        continue
                
                if temp_clipboard:
                    self.app.save_state()
                    self.app.clipboard_periods = temp_clipboard
                    
                    # 削除処理（リスト再構築方式）
                    for task, indices in tasks_to_delete.values():
                        if 'periods' in task:
                            task['periods'] = [p for i, p in enumerate(task['periods']) if i not in indices]
                    
                    if self.scene():
                        self.scene().clearSelection()
                    QTimer.singleShot(100, self.app.update_ui)
                    print(f"Cut {len(temp_clipboard)} items.")
            except Exception as e:
                print(f"Error in cut action: {e}")
            finally:
                self.app.chart_view.setUpdatesEnabled(True)
        elif action == color_action:
            color_groups = self.app.get_color_groups()
            dlg = ColorGridDialog(color_groups, self.app)
            if dlg.exec():
                self.app.save_state()
                if 'periods' not in self.task:
                    self.task['periods'] = [{'start_date': self.task.get('start_date', ''), 'end_date': self.task.get('end_date', '')}]
                self.task['periods'][self.period_index]['color'] = dlg.selected_color
                QTimer.singleShot(100, self.app.update_ui)
        elif action == del_action:
            if 'periods' in self.task:
                try:
                    self.app.save_state()
                    self.task['periods'].pop(self.period_index)
                    if self.scene():
                        self.scene().clearSelection()
                    QTimer.singleShot(100, self.app.update_ui)
                except IndexError:
                    pass

class HeaderScene(QGraphicsScene):
    def __init__(self, app):
        super().__init__()
        self.app = app

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            # ヘッダーの下半分（日付・曜日エリア）をクリックした場合のみ反応
            y = event.scenePos().y()
            if 35 <= y <= 70:
                self.app.save_state()
                x = event.scenePos().x()
                day_idx = int(x / self.app.day_width)
                if 0 <= day_idx < self.app.display_days:
                    d = self.app.min_date + timedelta(days=day_idx)
                    d_str = d.strftime("%Y-%m-%d")
                    
                    if d_str in self.app.custom_holidays:
                        del self.app.custom_holidays[d_str]
                    else:
                        self.app.custom_holidays[d_str] = "休日"
                    
                    self.app.draw_chart()
                    event.accept()
                    return
        super().mousePressEvent(event)

class ChartScene(QGraphicsScene):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.start_x = 0

    def mousePressEvent(self, e):
        # 1. アイテムの特定
        item = self.itemAt(e.scenePos(), self.app.chart_view.transform())
        target_bar = None
        if item:
            temp = item
            while temp:
                if isinstance(temp, GanttBarItem):
                    target_bar = temp
                    break
                temp = temp.parentItem()

        # 2. ドラッグモードの動的制御
        if target_bar:
            self.app.chart_view.setDragMode(QGraphicsView.DragMode.NoDrag)
        else:
            self.app.chart_view.setDragMode(QGraphicsView.DragMode.RubberBandDrag)

        # 3. すでに選択されているバーを左クリックした場合の選択維持（ドラッグ準備）
        if target_bar and e.button() == Qt.LeftButton and not (e.modifiers() & (Qt.ControlModifier | Qt.ShiftModifier)):
            if target_bar.isSelected():
                # 重要：現在選択されている全てのアイテムを記録し、
                # super().mousePressEvent(e) による他アイテムの選択解除を防ぐために後で復元する
                selected_items = self.selectedItems()
                super().mousePressEvent(e)
                for it in selected_items:
                    it.setSelected(True)
                return

        # 4. 右クリック処理
        if target_bar and e.button() == Qt.RightButton:
            target_bar.mousePressEvent(e)
            e.accept()
            return

        # 5. 背景クリックまたは未選択アイテムの通常処理
        if not target_bar and e.button() == Qt.LeftButton:
            if e.modifiers() & Qt.AltModifier:
                self.start_x = e.scenePos().x()
            else:
                self.start_x = 0
        
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
        gantt_item = next((it for it in items if isinstance(it, GanttBarItem)), None)
            
        if not gantt_item and e.button() == Qt.LeftButton:
            # Altキーが押されている場合のみ、ダブルクリックによる期間作成を許可
            if not (e.modifiers() & Qt.AltModifier):
                super().mouseDoubleClickEvent(e)
                return

            y = e.scenePos().y()
            row = int(y / self.app.row_height)
            if 0 <= row < len(self.app.visible_tasks_info):
                info = self.app.visible_tasks_info[row]
                task = info['task']
                if task.get('is_group'):
                    return

                x = e.scenePos().x()
                day_idx = int(x / self.app.day_width)
                d_str = (self.app.min_date + timedelta(days=day_idx)).strftime("%Y-%m-%d")
                
                if 'periods' not in task:
                    task['periods'] = [{'start_date': task.get('start_date', ''), 'end_date': task.get('end_date', '')}]
                
                self.app.save_state()
                task['periods'].append({"start_date": d_str, "end_date": d_str, "text": ""})
                QTimer.singleShot(100, self.app.update_ui)
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
            task = None
            
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
            
            paste_action = None
            if self.app.clipboard_periods:
                menu.addSeparator()
                paste_action = menu.addAction(f"貼り付け ({len(self.app.clipboard_periods)})")
                if not (0 <= row < len(self.app.visible_tasks_info)) or (task and task.get('is_group')):
                    paste_action.setEnabled(False)

            action = menu.exec(e.screenPos())
            x = e.scenePos().x()
            day_idx = int(x / self.app.day_width)
            
            if action == paste_action and paste_action:
                try:
                    self.app.save_state()
                    self.app.pending_selection = []
                    for cp in self.app.clipboard_periods:
                        new_period = cp['data'].copy()
                        target_v_row = row + cp['rel_row']
                        if 0 <= target_v_row < len(self.app.visible_tasks_info):
                            info = self.app.visible_tasks_info[target_v_row]
                            target_task = info['task']
                            if not target_task.get('is_group'):
                                # 日付の計算
                                try:
                                    sd = datetime.strptime(new_period['start_date'], "%Y-%m-%d")
                                    ed = datetime.strptime(new_period['end_date'], "%Y-%m-%d")
                                    duration = (ed - sd).days
                                    
                                    base_start = self.app.min_date + timedelta(days=day_idx)
                                    new_start = base_start + timedelta(days=cp['rel_days'])
                                    new_end = new_start + timedelta(days=duration)
                                    
                                    new_period['start_date'] = new_start.strftime("%Y-%m-%d")
                                    new_period['end_date'] = new_end.strftime("%Y-%m-%d")
                                    
                                    if 'periods' not in target_task:
                                        target_task['periods'] = []
                                    target_task['periods'].append(new_period)
                                    # 貼り付けたアイテムを記憶
                                    self.app.pending_selection.append(id(new_period))
                                except Exception:
                                    continue
                    QTimer.singleShot(100, self.app.update_ui)
                except Exception as e:
                    print(f"Error in paste action: {e}")
            elif action == add_period_action and add_period_action:
                if 'periods' not in task:
                    task['periods'] = [{'start_date': task.get('start_date', ''), 'end_date': task.get('end_date', '')}]
                self.app.save_state()
                task['periods'].append({"start_date": d_str, "end_date": d_str})
                QTimer.singleShot(100, self.app.update_ui)
            elif action == add_task_in_group and add_task_in_group:
                self.app.save_state()
                new_task = {
                    "name": "新規タスク",
                    "periods": [{"start_date": d_str, "end_date": d_str}],
                    "progress": 0,
                    "color": "#0078d4"
                }
                # グループの直後に挿入
                self.app.tasks.insert(info['index'] + 1, new_task)
                QTimer.singleShot(100, self.app.update_ui)
        else:
            super().contextMenuEvent(e)

class HideableHeader(QHeaderView):
    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self.setSectionsClickable(True)
        self.btn_size = 16
        
    def paintSection(self, painter, rect, logicalIndex):
        # 目玉アイコンなどのカスタム描画を削除し、標準のヘッダー表示に戻す
        super().paintSection(painter, rect, logicalIndex)

    def mouseReleaseEvent(self, e):
        # カスタムボタンを削除したため、標準のイベント処理のみ行う
        super().mouseReleaseEvent(e)

class TaskTable(QTableWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setHorizontalHeader(HideableHeader(Qt.Horizontal, self))
        self.setSelectionMode(QTableWidget.NoSelection)
        self.setSelectionBehavior(QTableWidget.SelectRows)
        self.setFocusPolicy(Qt.NoFocus)
        
        # ヘッダーの右クリックメニューを有効化
        self.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.horizontalHeader().customContextMenuRequested.connect(self.show_header_menu)

    def show_header_menu(self, pos):
        menu = QMenu(self)
        column_names = ["選択マーク", "開閉ボタン", "タスク名", "種別", "進捗(%)", "人数/工数", "期間指定/開始日", "色", "集計列"]
        for i, name in enumerate(column_names):
            action = menu.addAction(name)
            action.setCheckable(True)
            if i < 8:
                action.setChecked(not self.isColumnHidden(i))
            else:
                action.setChecked(self.window().summary_visible)
            
            if i == 2:
                action.setEnabled(False)
                
            action.toggled.connect(lambda checked, idx=i: self.window().toggle_column_visibility(idx, checked))
        
        menu.exec(self.horizontalHeader().mapToGlobal(pos))

class GanttApp(QMainWindow):
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
        
        # スクロールステップの更新 (スナップ用)
        if hasattr(self, 'chart_view'):
            self.chart_view.horizontalScrollBar().setSingleStep(max(1, int(self.day_width)))

    def on_zoom_changed(self, *_):
        self.zoom_unit = self.zoom_unit_combo.currentIndex()
        self.zoom_count = self.zoom_count_spin.value()
        # ズーム単位を表示単位（集計単位）にも適用する
        self.display_unit = self.zoom_unit
        self.update_display_days()
        self.calculate_day_width()
        self.update_ui()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'chart_view'):
            self.calculate_day_width()
            self.update_ui()

    def update_display_range(self):
        # ツールバーの状態を更新
        self.zoom_unit_combo.blockSignals(True)
        self.zoom_unit_combo.setCurrentIndex(self.display_unit)
        self.zoom_unit_combo.blockSignals(False)
        self.zoom_unit = self.display_unit
        
        self.update_display_days()
        self.calculate_day_width()
        self.update_ui()

    def update_display_days(self):
        # 明示的な終了日が設定されている場合はそれを使用
        if hasattr(self, 'max_date') and self.max_date:
            self.display_days = max(1, (self.max_date - self.min_date).days + 1)
            return

        # 単位と数に基づいて表示日数を計算する (互換性用)
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
        
        self.max_date = self.min_date + timedelta(days=self.display_days - 1)

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
        self.save_state()
        
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
        # 移動後の行に選択マークを更新
        self.update_selection_mark()

    def update_selection_mark(self, *args):
        self.table.blockSignals(True)
        curr = self.table.currentRow()
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it:
                it.setText("●" if r == curr else "")
                it.setTextAlignment(Qt.AlignCenter)
                it.setForeground(QColor(0, 120, 212)) # 青いドット
        self.table.blockSignals(False)

    def add_task(self):
        self.save_state()
        t = {
            "name": f"新規タスク {len(self.tasks)+1}",
            "periods": [],
            "progress": 0,
            "person_count": 1,
            "color": "#0078d4"
        }
        
        row = self.table.currentRow()
        if 0 <= row < len(self.visible_tasks_info):
            insert_idx = self.visible_tasks_info[row]['index'] + 1
            self.tasks.insert(insert_idx, t)
            self.update_ui()
            # 挿入された行を特定して選択
            for i, info in enumerate(self.visible_tasks_info):
                if info['task'] == t:
                    self.table.setCurrentCell(i, 2)
                    break
        else:
            self.tasks.append(t)
            self.update_ui()
            self.table.setCurrentCell(len(self.visible_tasks_info) - 1, 2)

    def add_group(self):
        self.save_state()
        g = {
            "name": f"新規グループ {len(self.tasks)+1}",
            "is_group": True,
            "collapsed": False,
            "color": "#555555"
        }
        
        row = self.table.currentRow()
        if 0 <= row < len(self.visible_tasks_info):
            insert_idx = self.visible_tasks_info[row]['index'] + 1
            self.tasks.insert(insert_idx, g)
            self.update_ui()
            # 挿入された行を特定して選択
            for i, info in enumerate(self.visible_tasks_info):
                if info['task'] == g:
                    self.table.setCurrentCell(i, 2)
                    break
        else:
            self.tasks.append(g)
            self.update_ui()
            self.table.setCurrentCell(len(self.visible_tasks_info) - 1, 2)

    def delete_task(self):
        r = self.table.currentRow()
        if r >= 0 and r < len(self.visible_tasks_info):
            self.save_state()
            idx = self.visible_tasks_info[r]['index']
            self.tasks.pop(idx)
            self.update_ui()

    def save_state(self):
        try:
            # 高速かつ確実なシリアライズ方式に戻す
            state = {
                "tasks": json.loads(json.dumps(self.tasks)),
                "custom_holidays": self.custom_holidays.copy()
            }
            # 直前の状態と同じなら保存しない
            if self.undo_stack and self.undo_stack[-1] == state:
                return
                
            self.undo_stack.append(state)
            if len(self.undo_stack) > 50:
                self.undo_stack.pop(0)
            self.redo_stack = []
            self.update_undo_redo_buttons()
        except Exception as e:
            print(f"Error saving state: {e}")

    def undo(self):
        if not self.undo_stack: return
        current_state = {
            "tasks": json.loads(json.dumps(self.tasks)),
            "custom_holidays": self.custom_holidays.copy()
        }
        self.redo_stack.append(current_state)
        state = self.undo_stack.pop()
        self.tasks = state["tasks"]
        self.custom_holidays = state["custom_holidays"]
        self.update_ui()
        self.update_undo_redo_buttons()

    def redo(self):
        if not self.redo_stack: return
        current_state = {
            "tasks": json.loads(json.dumps(self.tasks)),
            "custom_holidays": self.custom_holidays.copy()
        }
        self.undo_stack.append(current_state)
        state = self.redo_stack.pop()
        self.tasks = state["tasks"]
        self.custom_holidays = state["custom_holidays"]
        self.update_ui()
        self.update_undo_redo_buttons()

    def update_undo_redo_buttons(self):
        can_undo = len(self.undo_stack) > 0
        can_redo = len(self.redo_stack) > 0
        
        self.btn_undo.setEnabled(can_undo)
        self.btn_redo.setEnabled(can_redo)
        
        self.btn_undo.setToolTip("戻す" if can_undo else "戻す (操作履歴がありません)")
        self.btn_redo.setToolTip("進む" if can_redo else "進む (やり直しできる操作がありません)")

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

    def get_threshold_date(self, visible_start):
        if self.display_unit == 0: # 週間
            return visible_start + timedelta(days=2)
        elif self.display_unit == 1: # 月間
            return visible_start + timedelta(days=7)
        else: # 年間
            # 1ヶ月前 (約30日)
            return visible_start + timedelta(days=30)

    def get_summary_headers(self, base_date=None, count=None):
        if base_date is None: base_date = self.min_date
        if count is None:
            # 1画面の表示枠（ズーム設定）に合わせて集計列の数を決定する
            if self.zoom_unit == 0: # 週間
                visible_days = self.zoom_count * 7
            elif self.zoom_unit == 1: # 月間
                visible_days = self.zoom_count * 30.416
            else: # 年間
                visible_days = self.zoom_count * 365.25
                
            if self.display_unit == 0: # 週間
                count = max(1, round(visible_days / 7))
            elif self.display_unit == 1: # 月間
                count = max(1, round(visible_days / 30.416))
            else: # 年間
                count = max(1, round(visible_days / 365.25))
        
        headers = []
        curr = base_date
        unit_type = ['week', 'month', 'year'][self.display_unit]
        
        if unit_type == 'week':
            # 週の初め（月曜日）に合わせる
            curr = curr - timedelta(days=curr.weekday())
            for _ in range(count):
                end_d = curr + timedelta(days=6)
                label = f"{curr.strftime('%m/%d')}~{end_d.strftime('%m/%d')}"
                headers.append((curr, end_d, label))
                curr += timedelta(days=7)
        elif unit_type == 'month':
            curr = curr.replace(day=1)
            for _ in range(count):
                last_day = calendar.monthrange(curr.year, curr.month)[1]
                headers.append((curr, curr.replace(day=last_day), curr.strftime("%Y/%m")))
                m = curr.month + 1
                y = curr.year
                if m > 12: m = 1; y += 1
                curr = datetime(y, m, 1)
        elif unit_type == 'year':
            curr = curr.replace(month=1, day=1)
            for _ in range(count):
                headers.append((curr, curr.replace(month=12, day=31), curr.strftime("%Y年")))
                curr = curr.replace(year=curr.year + 1)
        return headers

    def sync_summary_to_scroll(self, base_date):
        if not hasattr(self, 'table'): return
        headers = self.get_summary_headers(base_date)
        
        self.table.blockSignals(True)
        # ヘッダーラベルの更新
        labels = ["", "", "タスク名", "種別", "進捗(%)", "人数/工数", "期間指定/開始日", "色"] + [h[2] for h in headers]
        # 現在の列数と合わない場合は調整（通常は update_ui で調整済み）
        if self.table.columnCount() != len(labels):
            self.table.setColumnCount(len(labels))
        self.table.setHorizontalHeaderLabels(labels)
        
        # 各行の集計値を更新
        for r, info in enumerate(self.visible_tasks_info):
            t = info['task']
            for i, (h_start, h_end, _) in enumerate(headers):
                col_idx = 8 + i
                item_s = self.table.item(r, col_idx)
                if not item_s:
                    item_s = QTableWidgetItem()
                    self.table.setItem(r, col_idx, item_s)
                
                item_s.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                item_s.setTextAlignment(Qt.AlignCenter)
                day_map = self.get_task_day_map_in_range(t, info['index'], h_start, h_end)
                item_s.setText(self.format_total_days(day_map))
                
                if t.get('is_group'):
                    item_s.setBackground(QColor(242, 242, 242))
                else:
                    item_s.setBackground(QColor(255, 255, 255))
                
                if self.summary_visible:
                    self.table.setColumnWidth(col_idx, 90)
                self.table.setColumnHidden(col_idx, not self.summary_visible)
                
        self.table.blockSignals(False)

    def get_task_day_map_in_range(self, t, start_idx, timeline_start=None, timeline_end=None):
        day_map = {}
        if timeline_start is None: timeline_start = self.min_date
        if timeline_end is None: timeline_end = self.min_date + timedelta(days=self.display_days - 1)
        
        tasks_to_sum = [t]
        if t.get('is_group'):
            tasks_to_sum = []
            for i in range(start_idx + 1, len(self.tasks)):
                if self.tasks[i].get('is_group'): break
                tasks_to_sum.append(self.tasks[i])
        
        for task in tasks_to_sum:
            for p in task.get('periods', []):
                if not p.get('start_date') or not p.get('end_date'): continue
                try:
                    psd = datetime.strptime(p['start_date'], "%Y-%m-%d")
                    ped = datetime.strptime(p['end_date'], "%Y-%m-%d")
                    overlap = (min(ped, timeline_end) - max(psd, timeline_start)).days + 1
                    if overlap > 0:
                        color = p.get('color', task.get('color', '#0078d4'))
                        # 人数を考慮して加算
                        day_map[color] = day_map.get(color, 0) + overlap * task.get('person_count', 1)
                except ValueError:
                    continue
        return day_map

    def sync_table_from_tasks(self):
        self.table.blockSignals(True)
        headers = self.get_summary_headers()
        for r, info in enumerate(self.visible_tasks_info):
            t = info['task']
            if t.get('is_group'):
                # グループの場合も集計を更新
                person_item = self.table.item(r, 5)
                if person_item:
                    total_p = 0
                    for i in range(info['index'] + 1, len(self.tasks)):
                        if self.tasks[i].get('is_group'): break
                        total_p += self.tasks[i].get('person_count', 1)
                    person_item.setText(str(total_p))

                for i, (h_start, h_end, _) in enumerate(headers):
                    col_idx = 8 + i
                    item_s = self.table.item(r, col_idx)
                    if item_s:
                        day_map = self.get_task_day_map_in_range(t, info['index'], h_start, h_end)
                        item_s.setText(self.format_total_days(day_map))
                continue
            
            periods = t.get('periods', [])
            p_strs = []
            for p in periods:
                if not p.get('start_date') or not p.get('end_date'): continue
                s = p['start_date'].replace('-', '/')
                e = p['end_date'].replace('-', '/')
                p_strs.append(f"{s}-{e}")
            
            person_item = self.table.item(r, 5)
            if person_item:
                if t.get('task_type') == 'generation':
                    person_item.setText(str(t.get('gen_workload', 1)))
                else:
                    person_item.setText(str(t.get('person_count', 1)))
                
            period_item = self.table.item(r, 6)
            if period_item:
                if t.get('task_type') == 'generation':
                    period_item.setText(t.get('gen_start_date', ''))
                else:
                    period_item.setText(", ".join(p_strs))
            
            # 動的な集計列の更新
            for i, (h_start, h_end, _) in enumerate(headers):
                col_idx = 8 + i
                item_s = self.table.item(r, col_idx)
                if item_s:
                    day_map = self.get_task_day_map_in_range(t, info['index'], h_start, h_end)
                    item_s.setText(self.format_total_days(day_map))
        self.table.blockSignals(False)

    def create_task_from_drag(self, x1, x2, y):
        self.save_state()
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
            "person_count": 1,
            "color": "#0078d4"
        }
        if row < len(self.visible_tasks_info):
            insert_idx = self.visible_tasks_info[row]['index']
            self.tasks.insert(insert_idx, t)
        else:
            self.tasks.append(t)
        self.update_ui()

    def recalculate_generation_tasks(self):
        current_group_idx = -1
        for i, t in enumerate(self.tasks):
            if t.get('is_group'):
                current_group_idx = i
            t['_group_idx'] = current_group_idx

        groups = {}
        for t in self.tasks:
            if t.get('is_group'): continue
            g_idx = t.get('_group_idx', -1)
            if g_idx not in groups:
                groups[g_idx] = {'creation': [], 'generation': []}
            if t.get('task_type') == 'generation':
                groups[g_idx]['generation'].append(t)
            else:
                groups[g_idx]['creation'].append(t)

        for g_idx, data in groups.items():
            creation_tasks = data['creation']
            generation_tasks = data['generation']
            if not generation_tasks:
                continue

            capacity = {}
            for ct in creation_tasks:
                for p in ct.get('periods', []):
                    if not p.get('start_date') or not p.get('end_date'): continue
                    try:
                        sd = datetime.strptime(p['start_date'], "%Y-%m-%d")
                        ed = datetime.strptime(p['end_date'], "%Y-%m-%d")
                        c = ct.get('person_count', 1)
                        curr = sd
                        while curr <= ed:
                            ds = curr.strftime("%Y-%m-%d")
                            capacity[ds] = capacity.get(ds, 0) + c
                            curr += timedelta(days=1)
                    except ValueError:
                        pass

            valid_gts = []
            for gt in generation_tasks:
                gt['_rem'] = float(gt.get('gen_workload', 0))
                sd_str = gt.get('gen_start_date', '')
                gt['_actual_start'] = None
                gt['_actual_end'] = None
                gt['person_count'] = 0 # Avoid double counting in summary
                if not sd_str or gt['_rem'] <= 0:
                    continue
                try:
                    gt['_sd'] = datetime.strptime(sd_str, "%Y-%m-%d")
                    valid_gts.append(gt)
                except ValueError:
                    continue

            if not valid_gts:
                for gt in generation_tasks:
                    if gt.get('gen_start_date') and gt.get('gen_workload', 0) <= 0:
                        sd = gt['gen_start_date']
                        gt['periods'] = [{'start_date': sd, 'end_date': sd, 'color': gt.get('color', '#0078d4')}]
                    elif not gt.get('gen_start_date'):
                        gt['periods'] = []
                continue

            current_date = min(gt['_sd'] for gt in valid_gts)
            end_limit = current_date + timedelta(days=3650)
            cap_copy = capacity.copy()

            while valid_gts and current_date <= end_limit:
                ds = current_date.strftime("%Y-%m-%d")
                daily_cap = cap_copy.get(ds, 0)
                
                active = [gt for gt in valid_gts if gt['_sd'] <= current_date]
                
                if active and daily_cap > 0:
                    active.sort(key=lambda x: x['_rem'])
                    rem_cap = daily_cap
                    
                    while rem_cap > 0.001 and active:
                        share = rem_cap / len(active)
                        if active[0]['_rem'] <= share + 0.001:
                            used = active[0]['_rem']
                            active[0]['_rem'] = 0
                            if active[0]['_actual_start'] is None:
                                active[0]['_actual_start'] = ds
                            active[0]['_actual_end'] = ds
                            rem_cap -= used
                            active.pop(0)
                        else:
                            for gt in active:
                                gt['_rem'] -= share
                                if gt['_actual_start'] is None:
                                    gt['_actual_start'] = ds
                                gt['_actual_end'] = ds
                            rem_cap = 0
                
                valid_gts = [gt for gt in valid_gts if gt['_rem'] > 0.001]
                current_date += timedelta(days=1)

            for gt in generation_tasks:
                if gt.get('_actual_start'):
                    sd = gt['_actual_start']
                    ed = gt['_actual_end'] if gt.get('_actual_end') else sd
                    color = gt.get('color', '#0078d4')
                    gt['periods'] = [{'start_date': sd, 'end_date': ed, 'color': color, 'text': f"{gt.get('gen_workload')}工数"}]
                else:
                    if gt.get('gen_start_date'):
                        sd = gt['gen_start_date']
                        gt['periods'] = [{'start_date': sd, 'end_date': sd, 'color': gt.get('color', '#0078d4'), 'text': '未着手'}]
                    else:
                        gt['periods'] = []

        for t in self.tasks:
            temp_keys = [k for k in t.keys() if str(k).startswith('_')]
            for k in temp_keys:
                del t[k]

    def update_ui(self, refresh_chart=True):
        if getattr(self, '_updating_ui', False):
            return
        self._updating_ui = True
        try:
            self.recalculate_generation_tasks()
            
            self.visible_tasks_info = self.get_visible_tasks_info()
            self.table.blockSignals(True)
            
            # 現在のスクロール位置から表示基準日を計算
            scroll_val = self.chart_view.horizontalScrollBar().value()
            days_scrolled = scroll_val / self.day_width if self.day_width > 0 else 0
            visible_start = self.min_date + timedelta(days=days_scrolled)
            threshold_date = self.get_threshold_date(visible_start)
            
            headers = self.get_summary_headers(threshold_date)
            base_col_count = 8
            total_cols = base_col_count + len(headers)
            self.table.setColumnCount(total_cols)
            
            labels = ["", "", "タスク名", "種別", "進捗(%)", "人数/工数", "期間指定/開始日", "色"] + [h[2] for h in headers]
            self.table.setHorizontalHeaderLabels(labels)
            
            new_rows = len(self.visible_tasks_info)
            if self.table.rowCount() < new_rows:
                self.table.setRowCount(new_rows)
                
            for r, info in enumerate(self.visible_tasks_info):
                t = info['task']
                indent = "    " * info['indent']
                is_group = t.get('is_group', False)
                
                # セルが存在しない場合のみ生成する
                for c in range(total_cols):
                    if self.table.item(r, c) is None:
                        self.table.setItem(r, c, QTableWidgetItem())
                
                # 0: Selection Mark
                mark_item = self.table.item(r, 0)
                mark_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                mark_item.setBackground(QColor(255, 255, 255))
                if is_group: mark_item.setBackground(QColor(242, 242, 242))
                # テキストは update_selection_mark で設定

                # 1: Toggle
                toggle_item = self.table.item(r, 1)
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

                # 2: Name
                item_name = self.table.item(r, 2)
                item_name.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                item_name.setForeground(QColor(51, 51, 51))
                f = item_name.font(); f.setBold(False); item_name.setFont(f)
                item_name.setBackground(QColor(255, 255, 255))
                item_name.setText(indent + t.get('name', ''))
                if is_group:
                    f = item_name.font(); f.setBold(True); item_name.setFont(f)
                    item_name.setBackground(QColor(242, 242, 242))
                
                # 3: Type
                item_type = self.table.item(r, 3)
                item_type.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                item_type.setForeground(QColor(51, 51, 51))
                item_type.setBackground(QColor(255, 255, 255))
                
                # 4: Progress
                item_prog = self.table.item(r, 4)
                item_prog.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                item_prog.setForeground(QColor(51, 51, 51))
                item_prog.setBackground(QColor(255, 255, 255))
                
                # 5: Person Count / Workload
                item_person = self.table.item(r, 5)
                item_person.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                item_person.setForeground(QColor(51, 51, 51))
                item_person.setBackground(QColor(255, 255, 255))
                
                # 6: Period / Start Date
                item_period = self.table.item(r, 6)
                item_period.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                item_period.setForeground(QColor(51, 51, 51))
                item_period.setBackground(QColor(255, 255, 255))
                
                # 7: Color
                item_color = self.table.item(r, 7)
                item_color.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                item_color.setForeground(QColor(51, 51, 51))
                item_color.setBackground(QColor(255, 255, 255))
                
                if is_group:
                    item_type.setText("")
                    item_type.setBackground(QColor(242, 242, 242))
                    
                    item_prog.setText("")
                    item_prog.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    
                    # 人数列にグループ内の合計人数を表示
                    total_p = 0
                    for i in range(info['index'] + 1, len(self.tasks)):
                        if self.tasks[i].get('is_group'): break
                        total_p += self.tasks[i].get('person_count', 1)
                    item_person.setText(str(total_p))
                    item_person.setTextAlignment(Qt.AlignCenter)
                    item_person.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    
                    item_period.setText("")
                    item_period.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    item_color.setText("")
                    item_color.setBackground(QColor(200, 200, 200))
                    item_color.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                else:
                    ttype = t.get('task_type', 'creation')
                    item_type.setText("生成" if ttype == 'generation' else "作成")
                    item_type.setTextAlignment(Qt.AlignCenter)
                    item_type.setToolTip("ダブルクリックで変更")
                    
                    if ttype == 'generation':
                        item_prog.setText("")
                        item_prog.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                        
                        item_person.setText(str(t.get('gen_workload', 1)))
                        item_person.setTextAlignment(Qt.AlignCenter)
                        
                        item_period.setText(t.get('gen_start_date', ''))
                    else:
                        item_prog.setText(str(t.get('progress', 0)))
                        item_prog.setTextAlignment(Qt.AlignCenter)
                        
                        item_person.setText(str(t.get('person_count', 1)))
                        item_person.setTextAlignment(Qt.AlignCenter)
                        
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

                # 8 onwards: Dynamic Summary Columns
                for i, (h_start, h_end, _) in enumerate(headers):
                    col_idx = 8 + i
                    item_s = self.table.item(r, col_idx)
                    item_s.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    item_s.setForeground(QColor(51, 51, 51))
                    item_s.setTextAlignment(Qt.AlignCenter)
                    
                    day_map = self.get_task_day_map_in_range(t, info['index'], h_start, h_end)
                    item_s.setText(self.format_total_days(day_map))
                    
                    if is_group:
                        item_s.setBackground(QColor(242, 242, 242))
                    else:
                        item_s.setBackground(QColor(255, 255, 255))
                    
            # 余分な行を削除
            if self.table.rowCount() > new_rows:
                self.table.setRowCount(new_rows)
                
            # 集計列の表示・非表示を一括反映
            for i in range(base_col_count, total_cols):
                self.table.setColumnHidden(i, not self.summary_visible)
                if self.summary_visible:
                    self.table.setColumnWidth(i, 90)

            self.update_selection_mark()
            self.table.blockSignals(False)
        finally:
            self._updating_ui = False
            
        if refresh_chart:
            self.draw_chart()

    def on_table_item_changed(self, item):
        self.save_state()
        row = item.row()
        col = item.column()
        if row >= len(self.visible_tasks_info) or row < 0: return
        info = self.visible_tasks_info[row]
        t = info['task']
        
        if col == 2: # Name
            t['name'] = item.text().strip()
        elif col == 4: # Progress
            if t.get('is_group') or t.get('task_type') == 'generation': return
            try:
                prog = int(item.text().replace('%', '').strip())
                t['progress'] = max(0, min(100, prog))
            except ValueError:
                pass
            self.table.blockSignals(True)
            item.setText(str(t['progress']))
            self.table.blockSignals(False)
        elif col == 5: # Person Count / Workload
            if t.get('is_group'): return
            try:
                val = float(item.text().strip()) if t.get('task_type') == 'generation' else int(item.text().strip())
                val = max(0, val)
                if t.get('task_type') == 'generation':
                    t['gen_workload'] = val
                else:
                    t['person_count'] = max(1, int(val))
            except ValueError:
                if t.get('task_type') != 'generation':
                    t['person_count'] = 1
            self.table.blockSignals(True)
            if t.get('task_type') == 'generation':
                item.setText(str(t.get('gen_workload', 1)))
            else:
                item.setText(str(t.get('person_count', 1)))
            self.table.blockSignals(False)
            self.update_ui(refresh_chart=False) # 集計と生成タスク再計算
        elif col == 6: # Period / Start Date
            if t.get('is_group'): return
            if t.get('task_type') == 'generation':
                parsed = self.parse_date(item.text())
                if parsed:
                    t['gen_start_date'] = parsed
                else:
                    QMessageBox.warning(self, "エラー", "日付の形式が正しくありません。\n例: 04/01")
                self.update_ui()
                return
            else:
                period_str = item.text()
                parsed = self.get_periods_from_string(period_str)
                if parsed:
                    t['periods'] = parsed
                else:
                    QMessageBox.warning(self, "エラー", "期間の形式が正しくありません。\n例: 04/01-04/05")
        
        self.draw_chart()

    def on_table_cell_clicked(self, row, col):
        if row >= len(self.visible_tasks_info): return
        self.update_selection_mark()
        if col == 1: # Toggle column
            info = self.visible_tasks_info[row]
            t = info['task']
            if t.get('is_group'):
                t['collapsed'] = not t.get('collapsed', False)
                self.update_ui()

    def on_table_cell_double_clicked(self, row, col):
        if row >= len(self.visible_tasks_info): return
        info = self.visible_tasks_info[row]
        t = info['task']
        
        if col == 3 and not t.get('is_group'): # Type
            self.save_state()
            if t.get('task_type') == 'generation':
                t['task_type'] = 'creation'
            else:
                t['task_type'] = 'generation'
                if 'gen_workload' not in t:
                    t['gen_workload'] = t.get('person_count', 1)
                if 'gen_start_date' not in t:
                    if t.get('periods') and t['periods'][0].get('start_date'):
                        t['gen_start_date'] = t['periods'][0]['start_date']
                    else:
                        t['gen_start_date'] = datetime.now().strftime("%Y-%m-%d")
            self.update_ui()
            
        elif col == 7: # Color column
            color_groups = self.get_color_groups()
            
            dlg = ColorGridDialog(color_groups, self)
            if dlg.exec():
                self.save_state()
                t['color'] = dlg.selected_color
                self.update_ui()


    def scroll_to_today(self):
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        target_date = today
        
        if self.display_unit == 0: # 週間
            target_date = today - timedelta(days=today.weekday())
        elif self.display_unit == 1: # 月間
            target_date = today.replace(day=1)
        elif self.display_unit == 2: # 年間
            target_date = today.replace(month=1, day=1)
            
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
            # 現在の週の月曜日を基準にする
            monday = current_date - timedelta(days=current_date.weekday())
            if direction > 0:
                # 次の週の月曜日へ
                new_date = monday + timedelta(days=7)
            else:
                # 現在が月曜日なら前週、そうでなければ今週の月曜日へ
                new_date = monday - timedelta(days=7) if current_date == monday else monday
        elif unit == 'month':
            # 現在の月の1日を基準にする
            first_day = current_date.replace(day=1)
            if direction > 0:
                # 翌月の1日へ
                m = first_day.month % 12 + 1
                y = first_day.year + (1 if first_day.month == 12 else 0)
                new_date = datetime(y, m, 1)
            else:
                # 現在が1日なら前月、そうでなければ今月の1日へ
                if current_date == first_day:
                    m = (first_day.month - 2) % 12 + 1
                    y = first_day.year - (1 if first_day.month == 1 else 0)
                    new_date = datetime(y, m, 1)
                else:
                    new_date = first_day
        elif unit == 'year':
            # 現在の年の1/1を基準にする
            jan_first = current_date.replace(month=1, day=1)
            if direction > 0:
                # 翌年の1/1へ
                new_date = jan_first.replace(year=jan_first.year + 1)
            else:
                # 現在が1/1なら前年、そうでなければ今年の1/1へ
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

    def draw_chart(self):
        # 再描画前に現在の選択状態を一時保存（復元用）
        if not getattr(self, 'pending_selection', []):
            sel = []
            for item in self.cs.selectedItems():
                if isinstance(item, GanttBarItem):
                    p_data = None
                    periods = item.task.get('periods', [])
                    if item.period_index < len(periods):
                        p_data = periods[item.period_index]
                    else:
                        # periodsがない場合やインデックス外の場合はタスク自身を対象とする
                        p_data = item.task
                    if p_data is not None:
                        sel.append(id(p_data))
            self.pending_selection = sel

        self.month_label_items = []
        self.hs.clear()
        self.cs.clear()
        tw_total = self.display_days * self.day_width
        ch = max(self.chart_view.height(), len(self.tasks) * self.row_height)
        if len(self.visible_tasks_info) > 0:
            ch = max(ch, len(self.visible_tasks_info) * self.row_height)
        
        last_m = None
        for i in range(self.display_days):
            d = self.min_date + timedelta(days=i)
            x = i * self.day_width
            d_str = d.strftime("%Y-%m-%d")
            
            is_custom = d_str in self.custom_holidays
            is_public = jpholiday.is_holiday(d)
            
            # 背景色の決定
            bg = None
            if d.weekday() == 5: # 土曜日
                if is_custom:
                    bg = QColor(255, 240, 240) # 既存機能：クリックで赤
                else:
                    bg = QColor(240, 248, 255) # デフォルト：青
            elif d.weekday() == 6 or is_public: # 日曜日または公的祝日
                if is_custom:
                    bg = None # クリックで無色（白）に反転
                else:
                    bg = QColor(255, 240, 240) # デフォルト：赤
            elif is_custom: # 平日でクリックされた場合
                bg = QColor(255, 240, 240) # 赤
            
            if bg:
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
            h_bg = QColor(255, 255, 225) if is_custom else QColor(248, 248, 248)
            self.hs.addRect(x, 35, self.day_width, 35, QPen(QColor(210, 210, 210)), QBrush(h_bg)).setZValue(5)
            
            if self.day_width >= 35:
                dl = self.hs.addText(d.strftime("%d"))
                # 日付数字の色（すべて一律で黒系にする）
                day_color = QColor(50, 50, 50)
                
                dl.setDefaultTextColor(day_color)
                dl.setFont(QFont("Segoe UI", 9, QFont.Bold))
                dl.setPos(x + (self.day_width/2) - 10, 35)
                dl.setZValue(10)
                
                # 曜日の色
                if d.weekday() == 5 and not is_public: # 土曜日
                    w_c = QColor(0, 80, 200)
                elif d.weekday() == 6 or is_public: # 日曜または公的祝日
                    w_c = QColor(220, 0, 0)
                else:
                    w_c = QColor(60, 60, 60)
                
                yl = self.hs.addText(["月","火","水","木","金","土","日"][d.weekday()])
                yl.setDefaultTextColor(w_c)
                yl.setFont(QFont("Segoe UI", 7))
                yl.setPos(x + (self.day_width/2) - 8, 52)
                yl.setZValue(10)
                
                if is_public or is_custom:
                    h_name = self.custom_holidays.get(d_str) or jpholiday.is_holiday_name(d)
                    if h_name:
                        dl.setToolTip(h_name)
                        yl.setToolTip(h_name)
                        
                        # ヘッダー背景にもツールチップを設定して、どこをホバーしても祝日名が出るようにする
                        header_bg = self.hs.addRect(x, 35, self.day_width, 35, QPen(Qt.NoPen), QBrush(Qt.transparent))
                        header_bg.setZValue(11)
                        header_bg.setToolTip(h_name)
            
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
                    # グループ内のタスクのバーの数を集計
                    counts = [0] * self.display_days
                    for i in range(info['index'] + 1, len(self.tasks)):
                        sub_t = self.tasks[i]
                        if sub_t.get('is_group'): break
                        sub_periods = sub_t.get('periods', [])
                        for p in sub_periods:
                            if not p.get('start_date') or not p.get('end_date'): continue
                            psd = datetime.strptime(p['start_date'], "%Y-%m-%d")
                            ped = datetime.strptime(p['end_date'], "%Y-%m-%d")
                            # 表示範囲内での重なりを計算
                            s_idx = max(0, (psd - self.min_date).days)
                            e_idx = min(self.display_days - 1, (ped - self.min_date).days)
                            p_count = sub_t.get('person_count', 1)
                            for d_idx in range(s_idx, e_idx + 1):
                                counts[d_idx] += p_count
                    
                    # 集計結果を描画
                    for d_idx, count in enumerate(counts):
                        if count > 0:
                            x = d_idx * self.day_width
                            # 背景に薄い円を表示
                            r = min(self.day_width * 0.8, self.row_height * 0.6)
                            self.cs.addEllipse(x + (self.day_width - r)/2, row * self.row_height + (self.row_height - r)/2, r, r, 
                                               QPen(Qt.NoPen), QBrush(QColor(0, 120, 212, 40))).setZValue(15)
                            
                            # 数字を表示
                            txt = self.cs.addText(str(count))
                            txt.setFont(QFont("Segoe UI", 9, QFont.Bold))
                            txt.setDefaultTextColor(QColor(0, 120, 212))
                            tw = txt.boundingRect().width()
                            th = txt.boundingRect().height()
                            txt.setPos(x + (self.day_width - tw)/2, row * self.row_height + (self.row_height - th)/2)
                            txt.setZValue(20)
                    continue

                periods = t.get('periods')
                if periods is None:
                    if t.get('start_date') and t.get('end_date'):
                        # 互換性のため、データモデル自体を periods 形式に移行して ID を安定させる
                        t['periods'] = [{'start_date': t['start_date'], 'end_date': t['end_date']}]
                        periods = t['periods']
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
                    
                    # 移動後などの再選択
                    if id(p) in getattr(self, 'pending_selection', []):
                        bar.setSelected(True)
                        
                    self.cs.addItem(bar)
            except Exception as e:
                print(f"Error drawing bar for row {row}: {e}")
                
        self.hs.setSceneRect(0, 0, tw_total, self.header_height)
        self.cs.setSceneRect(0, 0, tw_total, ch)
        self.update_month_labels_pos()
        
        # 再選択リストをクリア
        if hasattr(self, 'pending_selection'):
            self.pending_selection = []

    def save_data(self):
        initial_dir = os.path.dirname(self.last_path) if self.last_path and os.path.exists(os.path.dirname(self.last_path)) else ""
        p = QFileDialog.getSaveFileName(self, "保存", initial_dir, "JSON (*.json)")[0]
        if p:
            self.last_path = p
            try:
                data_to_save = {
                    "settings": {
                        "min_date": self.min_date.strftime("%Y-%m-%d"),
                        "max_date": self.max_date.strftime("%Y-%m-%d") if hasattr(self, 'max_date') else None,
                        "display_unit": self.display_unit,
                        "display_count": self.display_count,
                        "zoom_unit": self.zoom_unit,
                        "zoom_count": self.zoom_count,
                        "custom_holidays": self.custom_holidays
                    },
                    "tasks": self.tasks
                }
                with open(p, 'w', encoding='utf-8') as f:
                    json.dump(data_to_save, f, ensure_ascii=False, indent=4, default=str)
                QMessageBox.information(self, "成功", "保存しました。")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"保存失敗: {e}")

    def get_color_groups(self):
        return [
            ("青・水色系", [
                ("青", "#0078d4"), ("水色", "#00bcf2"), ("紺", "#002050"), 
                ("空色", "#87ceeb"), ("ロイヤルブルー", "#4169e1"), ("ネイビー", "#000080")
            ]),
            ("緑・ライム系", [
                ("緑", "#107c10"), ("ライム", "#32cd32"), ("深緑", "#004b1c"),
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
                ("黒", "#323130"), ("灰色", "#7a7574"), ("シルバー", "#c0c0c0"),
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

    def format_total_days(self, day_map):
        if not day_map: return "0工数"
        total = sum(day_map.values())
        if len(day_map) <= 1:
            return f"{total}工数"
        
        parts = []
        # 色コードでソートして順序を固定
        for code in sorted(day_map.keys()):
            days = day_map[code]
            name = self.get_color_name(code)
            parts.append(f"{name}:{days}")
        return f"計{total}工数 ({', '.join(parts)})"

    def toggle_column_visibility(self, idx, visible):
        if idx < 8:
            self.table.setColumnHidden(idx, not visible)
        else:
            self.summary_visible = visible
            # 8列目以降の全列をトグル
            for i in range(8, self.table.columnCount()):
                self.table.setColumnHidden(i, not visible)
        
        if idx in self.col_actions:
            self.col_actions[idx].blockSignals(True)
            self.col_actions[idx].setChecked(visible)
            self.col_actions[idx].blockSignals(False)

    def save_app_config(self):
        config_dir = os.path.join(os.environ.get('USERPROFILE', os.path.expanduser('~')), 'MiyaGantt')
        if not os.path.exists(config_dir):
            try:
                os.makedirs(config_dir)
            except Exception as e:
                QMessageBox.warning(self, "エラー", f"フォルダの作成に失敗しました: {e}")
                return
        
        path = os.path.join(config_dir, 'config.json')
        
        column_visibility = {}
        for i in range(9):
            column_visibility[str(i)] = not self.table.isColumnHidden(i)
        
        config = {
            "zoom_unit": self.zoom_unit,
            "zoom_count": self.zoom_count,
            "display_unit": self.display_unit,
            "display_count": self.display_count,
            "summary_visible": self.summary_visible,
            "column_visibility": column_visibility,
            "custom_holidays": self.custom_holidays,
            "last_path": self.last_path
        }
        
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4, default=str)
            QMessageBox.information(self, "完了", f"基本設定を保存しました:\n{path}")
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"設定の保存に失敗しました: {e}")

    def load_app_config(self):
        config_dir = os.path.join(os.environ.get('USERPROFILE', os.path.expanduser('~')), 'MiyaGantt')
        path = os.path.join(config_dir, 'config.json')
        if not os.path.exists(path): return
        
        try:
            with open(path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            self.zoom_unit = config.get("zoom_unit", self.zoom_unit)
            self.zoom_count = config.get("zoom_count", self.zoom_count)
            self.display_unit = config.get("display_unit", self.display_unit)
            self.display_count = config.get("display_count", self.display_count)
            self.summary_visible = config.get("summary_visible", self.summary_visible)
            self.custom_holidays = config.get("custom_holidays", self.custom_holidays)
            self.last_path = config.get("last_path", "")
            
            # UI部品への反映
            if hasattr(self, 'zoom_unit_combo'):
                self.zoom_unit_combo.blockSignals(True)
                self.zoom_unit_combo.setCurrentIndex(self.zoom_unit)
                self.zoom_unit_combo.blockSignals(False)
            if hasattr(self, 'zoom_count_spin'):
                self.zoom_count_spin.blockSignals(True)
                self.zoom_count_spin.setValue(self.zoom_count)
                self.zoom_count_spin.blockSignals(False)
            
            # 列の表示非表示
            col_vis = config.get("column_visibility", {})
            for idx_str, visible in col_vis.items():
                self.toggle_column_visibility(int(idx_str), visible)
        except Exception as e:
            print(f"Config load error: {e}")

    def load_data(self):
        initial_dir = os.path.dirname(self.last_path) if self.last_path and os.path.exists(os.path.dirname(self.last_path)) else ""
        p = QFileDialog.getOpenFileName(self, "開く", initial_dir, "JSON (*.json)")[0]
        if p:
            self.last_path = p
            try:
                with open(p, 'r', encoding='utf-8') as f:
                    loaded_data = json.load(f)
                
                if isinstance(loaded_data, dict) and "tasks" in loaded_data:
                    self.tasks = loaded_data["tasks"]
                    settings = loaded_data.get("settings", {})
                    min_date_str = settings.get("min_date")
                    max_date_str = settings.get("max_date")
                    if min_date_str:
                        self.min_date = datetime.strptime(min_date_str, "%Y-%m-%d")
                    else:
                        self.min_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
                    
                    if max_date_str:
                        self.max_date = datetime.strptime(max_date_str, "%Y-%m-%d")
                    else:
                        self.max_date = None # 下記の update_display_days で計算される
                    
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
                    self.custom_holidays = settings.get("custom_holidays", {})
                    
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