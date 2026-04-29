# tsukasamiyashita/miyagantt/MiyaGantt-fdab2c007da130510fa926a9e04a8f0ba70d9678/graphics.py
import sys
import os
import calendar
from datetime import datetime, timedelta
import jpholiday
import shiboken6
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *

from dialogs import ColorGridDialog

def log_event(event_name, details=""):
    """デバッグ用ログ出力関数"""
    try:
        ts = datetime.now().strftime('%H:%M:%S.%f')[:-3]
        print(f"[{ts}] {event_name}: {details}")
    except Exception as e:
        print(f"Log Error: {e}")

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
                self.setPen(QPen(QColor("#ff8c00"), 3))
                self.setBrush(QBrush(bc.lighter(170)))
                self.setZValue(40)
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

            periods = self.task.get('periods')
            if periods is not None and self.period_index < len(periods):
                p_dict = periods[self.period_index]
            else:
                p_dict = self.task

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
        if change == QGraphicsRectItem.ItemSelectedChange:
            log_event("GanttBarItem.itemChange", f"ItemSelectedChange triggered, Row: {self.row}")
            QTimer.singleShot(0, self.update_appearance)
        return super().itemChange(change, value)

    def mousePressEvent(self, event):
        log_event("GanttBarItem.mousePressEvent", f"Row: {self.row}, Button: {event.button()}")
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
            selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttBarItem)]
            if len(selected_items) > 1 and self.isSelected():
                delta = event.scenePos() - event.lastScenePos()
                for it in selected_items:
                    it.setPos(it.pos() + delta)
                    it_row = int(it.scenePos().y() / self.app.row_height)
                    it.setPos(it.pos().x(), it_row * self.app.row_height + 10)
                    it.update_appearance()
            else:
                super().mouseMoveEvent(event)
                row = int(event.scenePos().y() / self.app.row_height)
                max_row = len(self.app.visible_tasks_info) - 1 if self.app.visible_tasks_info else 0
                row = max(0, min(max_row, row))
                self.setPos(self.pos().x(), row * self.app.row_height + 10)
        self.update_appearance()

    def mouseReleaseEvent(self, event):
        log_event("GanttBarItem.mouseReleaseEvent", f"Row: {self.row}, Resizing: left={self.resizing_left}, right={self.resizing_right}")
        was_resizing = self.resizing_left or self.resizing_right
        selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttBarItem)]
        
        if not was_resizing and len(selected_items) > 1 and self.isSelected():
            self.app.save_state()
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
            
            tasks_to_delete = {}
            to_insert = []

            try:
                self.app.chart_view.setUpdatesEnabled(False)
                for target in move_targets:
                    it = target['item']
                    if 'periods' in it.task and it.period_index < len(it.task['periods']):
                        t_id = id(it.task)
                        if t_id not in tasks_to_delete:
                            tasks_to_delete[t_id] = (it.task, set())
                        
                        if it.period_index not in tasks_to_delete[t_id][1]:
                            tasks_to_delete[t_id][1].add(it.period_index)
                            p = it.task['periods'][it.period_index].copy()
                            p['start_date'] = target['start_date']
                            p['end_date'] = target['end_date']
                            target_task = self.app.visible_tasks_info[target['new_row']]['task']
                            to_insert.append((target_task, p))
                
                for task, indices in tasks_to_delete.values():
                    if 'periods' in task:
                        task['periods'] = [p for i, p in enumerate(task['periods']) if i not in indices]
                
                self.app.pending_selection = []
                for target_task, p_data in to_insert:
                    if target_task.get('is_group'):
                        continue
                    if 'periods' not in target_task:
                        target_task['periods'] = []
                    target_task['periods'].append(p_data)
                    self.app.pending_selection.append(id(p_data))

                QTimer.singleShot(100, self.app.update_ui)
            except Exception as e:
                log_event("Error in multi-move", str(e))
            finally:
                self.app.chart_view.setUpdatesEnabled(True)
            
            super().mouseReleaseEvent(event)
            return

        self.resizing_left = self.resizing_right = False
        self.setCursor(Qt.OpenHandCursor)
        snap = self.app.day_width
        sx = round(self.pos().x() / snap) * snap
        sw = max(snap, round(self.rect().width() / snap) * snap)
        
        sd = self.app.min_date + timedelta(days=sx / self.app.day_width)
        ed = sd + timedelta(days=sw / self.app.day_width - 0.001)

        new_row = int(event.scenePos().y() / self.app.row_height)
        max_row = len(self.app.visible_tasks_info) - 1 if self.app.visible_tasks_info else 0
        new_row = max(0, min(max_row, new_row))
        
        if not was_resizing and new_row != self.row:
            target_info = self.app.visible_tasks_info[new_row]
            target_task = target_info['task']
            
            if target_task.get('is_group'):
                QTimer.singleShot(100, self.app.update_ui)
                super().mouseReleaseEvent(event)
                return

            for t in [self.task, target_task]:
                if 'periods' not in t:
                    t['periods'] = [{'start_date': t.get('start_date', ''), 'end_date': t.get('end_date', '')}]
            
            if 0 <= self.period_index < len(self.task['periods']):
                self.app.save_state()
                p = self.task['periods'].pop(self.period_index)
                p['start_date'] = sd.strftime("%Y-%m-%d")
                p['end_date'] = ed.strftime("%Y-%m-%d")
                
                target_task = self.app.tasks[new_row]
                target_task['periods'].append(p)
                
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
        
        self.task['start_date'] = self.task['periods'][0]['start_date']
        self.task['end_date'] = self.task['periods'][0]['end_date']
        
        if self.scene():
            for item in self.scene().items():
                if isinstance(item, GanttBarItem) and item.task is self.task:
                    item.update_appearance()
                    
        self.app.sync_table_from_tasks()
        if was_resizing or new_row != self.row or sx != self.start_pos.x():
            QTimer.singleShot(100, self.app.update_ui)
        else:
            self.update_appearance()

    def mouseDoubleClickEvent(self, event):
        log_event("GanttBarItem.mouseDoubleClickEvent", f"Row: {self.row}")
        super().mouseDoubleClickEvent(event)
        
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
                QTimer.singleShot(100, self.app.update_ui)

    def contextMenuEvent(self, event):
        log_event("GanttBarItem.contextMenuEvent", f"Row: {self.row}")
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
                valid_items = []
                for it in selected_items:
                    if shiboken6.isValid(it):
                        valid_items.append(it)
                
                if not valid_items: return
                
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
                    log_event("ContextMenu", f"Copied {len(temp_clipboard)} items.")
            except Exception as e:
                log_event("Error in copy action", str(e))
        elif action == cut_action:
            if not selected_items: return
            try:
                self.app.chart_view.setUpdatesEnabled(False)
                valid_items = []
                for it in selected_items:
                    if shiboken6.isValid(it):
                        valid_items.append(it)
                
                if not valid_items: return

                temp_clipboard = []
                min_x = min(it.pos().x() for it in valid_items)
                min_row = min(it.row for it in valid_items)
                
                tasks_to_delete = {}
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
                    
                    for task, indices in tasks_to_delete.values():
                        if 'periods' in task:
                            task['periods'] = [p for i, p in enumerate(task['periods']) if i not in indices]
                    
                    if self.scene():
                        self.scene().clearSelection()
                    QTimer.singleShot(100, self.app.update_ui)
                    log_event("ContextMenu", f"Cut {len(temp_clipboard)} items.")
            except Exception as e:
                log_event("Error in cut action", str(e))
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
        log_event("HeaderScene.mousePressEvent", f"Pos: {event.scenePos().x()}, {event.scenePos().y()}")
        if hasattr(self.app, 'hv') and event.widget() != self.app.hv.viewport():
            return

        if event.button() == Qt.LeftButton:
            y = event.scenePos().y()
            if 35 <= y <= 70:
                self.app.save_state()
                x = event.scenePos().x()
                day_idx = int(x / self.app.day_width)
                if 0 <= day_idx < self.app.display_days:
                    d = self.app.min_date + timedelta(days=day_idx)
                    d_str = d.strftime("%Y-%m-%d")
                    
                    is_public = jpholiday.is_holiday(d)
                    is_default_holiday = d.weekday() in (5, 6) or is_public
                    
                    if d_str in self.app.custom_holidays:
                        del self.app.custom_holidays[d_str]
                    else:
                        if is_default_holiday:
                            self.app.custom_holidays[d_str] = "営業日"
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
        log_event("ChartScene.mousePressEvent", f"Pos: {e.scenePos().x()}, {e.scenePos().y()}, Button: {e.button()}")
        
        item = self.itemAt(e.scenePos(), self.app.chart_view.transform())
        target_bar = None
        if item:
            temp = item
            while temp:
                if isinstance(temp, GanttBarItem):
                    target_bar = temp
                    break
                temp = temp.parentItem()

        if not target_bar:
            if e.button() == Qt.LeftButton:
                y = e.scenePos().y()
                row = int(y / self.app.row_height)
                if 0 <= row < len(self.app.visible_tasks_info):
                    self.app.table.setCurrentCell(row, 2)

                if e.modifiers() & Qt.AltModifier:
                    self.start_x = e.scenePos().x()
                else:
                    self.start_x = 0
            
            e.accept()
            return

        if target_bar and e.button() == Qt.LeftButton and not (e.modifiers() & (Qt.ControlModifier | Qt.ShiftModifier)):
            if target_bar.isSelected():
                selected_items = self.selectedItems()
                super().mousePressEvent(e)
                for it in selected_items:
                    it.setSelected(True)
                return

        if target_bar and e.button() == Qt.RightButton:
            target_bar.mousePressEvent(e)
            e.accept()
            return
        
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        item = self.itemAt(e.scenePos(), self.app.chart_view.transform())
        target_bar = None
        if item:
            temp = item
            while temp:
                if isinstance(temp, GanttBarItem):
                    target_bar = temp
                    break
                temp = temp.parentItem()
                
        if not target_bar:
            e.accept()
            return
            
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        log_event("ChartScene.mouseReleaseEvent", f"Pos: {e.scenePos().x()}, {e.scenePos().y()}, start_x: {self.start_x}")
        if self.start_x > 0:
            item = self.itemAt(e.scenePos(), self.app.chart_view.transform())
            target_bar = None
            if item:
                temp = item
                while temp:
                    if isinstance(temp, GanttBarItem):
                        target_bar = temp
                        break
                    temp = temp.parentItem()

            if not target_bar:
                if abs(e.scenePos().x() - self.start_x) > (self.app.day_width * 0.1):
                    sx = self.start_x
                    ex = e.scenePos().x()
                    ey = e.scenePos().y()
                    self.app.create_task_from_drag(sx, ex, ey)
            self.start_x = 0
            
        item = self.itemAt(e.scenePos(), self.app.chart_view.transform())
        target_bar = None
        if item:
            temp = item
            while temp:
                if isinstance(temp, GanttBarItem):
                    target_bar = temp
                    break
                temp = temp.parentItem()
                
        if not target_bar:
            e.accept()
            return
            
        super().mouseReleaseEvent(e)

    def mouseDoubleClickEvent(self, e):
        log_event("ChartScene.mouseDoubleClickEvent", f"Pos: {e.scenePos().x()}, {e.scenePos().y()}")
        item = self.itemAt(e.scenePos(), self.app.chart_view.transform())
        target_bar = None
        if item:
            temp = item
            while temp:
                if isinstance(temp, GanttBarItem):
                    target_bar = temp
                    break
                temp = temp.parentItem()
            
        if not target_bar:
            if e.button() == Qt.LeftButton:
                y = e.scenePos().y()
                row = int(y / self.app.row_height)
                if 0 <= row < len(self.app.visible_tasks_info):
                    info = self.app.visible_tasks_info[row]
                    task = info['task']
                    if not task.get('is_group'):
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
        log_event("ChartScene.contextMenuEvent", f"Pos: {e.scenePos().x()}, {e.scenePos().y()}")
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
            d_str = (self.app.min_date + timedelta(days=day_idx)).strftime("%Y-%m-%d")
            
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
                                    self.app.pending_selection.append(id(new_period))
                                except Exception:
                                    continue
                    QTimer.singleShot(100, self.app.update_ui)
                except Exception as e:
                    log_event("Error in paste action", str(e))
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
                self.app.tasks.insert(info['index'] + 1, new_task)
                QTimer.singleShot(100, self.app.update_ui)
        else:
            super().contextMenuEvent(e)