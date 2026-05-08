# tsukasamiyashita/miyagantt/MiyaGantt-46a1664b6d1737cb32f1dd17429ce06cca8dc678/gantt_items.py
from datetime import datetime, timedelta
from PySide6.QtWidgets import (QGraphicsRectItem, QGraphicsTextItem, QGraphicsScene, 
                               QMenu, QLineEdit, QInputDialog)
from PySide6.QtCore import Qt, QRectF, QTimer, QPointF
from PySide6.QtGui import QBrush, QPen, QColor, QFont, QPainter
from dialogs import ColorGridDialog

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
        
        self.text_item = QGraphicsTextItem('', self)
        self.text_item.setDefaultTextColor(Qt.white)
        self.text_item.setZValue(1)
        font = QFont("Segoe UI", 9, QFont.Bold)
        self.text_item.setFont(font)
        self.text_item.setAcceptHoverEvents(False)
        
        self.resizing_left = False
        self.resizing_right = False
        
        periods = self.task.get('periods', [])
        if self.period_index < len(periods):
            self.period_dict = periods[self.period_index]
        else:
            self.period_dict = self.task
            
        self.update_appearance()

    def update_appearance(self):
        periods = self.task.get('periods', [])
        p_dict = periods[self.period_index] if self.period_index < len(periods) else self.task
        color_code = p_dict.get('color', self.task.get('color', '#0078d4'))
        bc = QColor(color_code)
        
        self.setPen(QPen(Qt.black if self.isSelected() else bc.darker(120), 2 if self.isSelected() else 1))
        self.setBrush(QBrush(bc))

        periods = self.task.get('periods')
        if periods is not None and self.period_index < len(periods):
            p_dict = periods[self.period_index]
        else:
            p_dict = self.task

        bar_text = p_dict.get('text', '')
        self.text_item.setPlainText(bar_text)
        
        text_w = self.text_item.boundingRect().width()
        bar_w = self.rect().width()
        
        if text_w > bar_w - 10 or "⚠️" in bar_text:
            self.text_item.setPos(bar_w + 5, (self.rect().height() - self.text_item.boundingRect().height()) / 2)
            if "⚠️" in bar_text:
                self.text_item.setDefaultTextColor(QColor("#d13438"))
            else:
                self.text_item.setDefaultTextColor(QColor(50, 50, 50))
            self.setZValue(35)
        else:
            self.text_item.setPos(5, (self.rect().height() - self.text_item.boundingRect().height()) / 2)
            self.text_item.setDefaultTextColor(Qt.white)
            self.setZValue(30)

        start_d = p_dict.get('start_date', '')
        end_d = p_dict.get('end_date', '')
        mode_str = "⚡ 案件モード" if self.task.get('mode') == 'auto' else "📝 メモモード" if self.task.get('mode') == 'memo' else "👤 人員モード"
        self.setToolTip(f"タスク: {self.task.get('name','')}\nモード: {mode_str}\n期間: {start_d}〜{end_d}")

    def paint(self, painter, option, widget=None):
        super().paint(painter, option, widget)
        
        if self.task.get('mode') == 'auto' and self.rect().width() > 0:
            allocs = self.task.get('daily_allocations', {})
            if not allocs:
                return
            
            p_dict = self.period_dict
            try:
                sd = datetime.strptime(p_dict.get('start_date', ''), "%Y-%m-%d")
                ed = datetime.strptime(p_dict.get('end_date', ''), "%Y-%m-%d")
            except ValueError:
                return
            
            days = (ed - sd).days + 1
            if days <= 0:
                return
            
            dw = self.app.day_width
            if dw < 15: 
                return
                
            painter.save()
            font = QFont("Segoe UI", max(6, min(9, int(dw/3))))
            painter.setFont(font)
            
            painter.setPen(Qt.white)
            
            h = self.rect().height()
            
            sd_str = sd.strftime("%Y-%m-%d")
            cumulative_val = sum(v for k, v in allocs.items() if k < sd_str)
            
            for i in range(days):
                d_str = (sd + timedelta(days=i)).strftime("%Y-%m-%d")
                val = allocs.get(d_str, 0.0)
                if val > 0.001:
                    cumulative_val += val
                    disp_val = round(cumulative_val, 2)
                    text = f"{disp_val:g}工数" if dw >= 40 else f"{disp_val:g}"
                    rx = self.rect().left() + i * dw
                    t_rect = QRectF(rx, self.rect().top(), dw, h)
                    painter.drawText(t_rect, Qt.AlignCenter, text)
            
            painter.restore()

    def hoverMoveEvent(self, event):
        x = event.pos().x()
        w = self.rect().width()
        margin = 12 if w <= self.app.day_width else 10
        margin = min(margin, w / 2 - 2)
        
        if self.task.get('mode') == 'auto':
            self.setCursor(Qt.OpenHandCursor)
        elif x < margin or x > w - margin:
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
            
            if self.task.get('mode') == 'auto':
                self.setCursor(Qt.ClosedHandCursor)
                self.drag_start_scene_pos = event.scenePos()
                self.drag_item_starts = {}
                selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttBarItem)]
                if self not in selected_items:
                    selected_items.append(self)
                for it in selected_items:
                    it.drag_start_pos = it.pos()
                    it.drag_start_row = it.row
            else:
                if x < margin:
                    self.resizing_left = True
                elif x > w - margin:
                    self.resizing_right = True
                else:
                    self.setCursor(Qt.ClosedHandCursor)
                    self.drag_start_scene_pos = event.scenePos()
                    self.drag_item_starts = {}
                    selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttBarItem)]
                    if self not in selected_items:
                        selected_items.append(self)
                    
                    for it in selected_items:
                        it.drag_start_pos = it.pos()
                        it.drag_start_row = it.row
                    
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        snap_x = self.app.day_width
        snap_y = self.app.row_height
        
        if self.resizing_left:
            diff = event.scenePos().x() - event.lastScenePos().x()
            nr = self.rect()
            if nr.width() - diff >= snap_x:
                self.setPos(self.pos().x() + diff, self.pos().y())
                self.setRect(0, 0, nr.width() - diff, nr.height())
        elif self.resizing_right:
            diff = event.scenePos().x() - event.lastScenePos().x()
            nr = self.rect()
            if nr.width() + diff >= snap_x:
                self.setRect(0, 0, nr.width() + diff, nr.height())
        elif hasattr(self, 'drag_start_scene_pos'):
            delta = event.scenePos() - self.drag_start_scene_pos
            dx = round(delta.x() / snap_x) * snap_x
            dy = round(delta.y() / snap_y) * snap_y
            
            selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttBarItem)]
            if self not in selected_items:
                selected_items.append(self)
                
            for it in selected_items:
                if hasattr(it, 'drag_start_pos'):
                    new_row = it.drag_start_row + int(dy / snap_y)
                    max_row = len(self.app.visible_tasks_info) - 1 if self.app.visible_tasks_info else 0
                    new_row = max(0, min(max_row, new_row))
                    
                    target_task = self.app.visible_tasks_info[new_row]['task']
                    source_mode = it.task.get('mode', 'manual')
                    target_mode = target_task.get('mode', 'manual')
                    
                    if target_task.get('is_group') or source_mode != target_mode:
                        new_row = it.drag_start_row
                        
                    it.setPos(it.drag_start_pos.x() + dx, new_row * snap_y + 10)
                    it.update_appearance()
        else:
            super().mouseMoveEvent(event)
            
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

        if was_resizing and self.task.get('mode') != 'auto':
            if 'periods' not in self.task:
                self.task['periods'] = [{'start_date': self.task.get('start_date', ''), 'end_date': self.task.get('end_date', '')}]
            if 0 <= self.period_index < len(self.task['periods']):
                self.task['periods'][self.period_index]['start_date'] = sd.strftime("%Y-%m-%d")
                self.task['periods'][self.period_index]['end_date'] = ed.strftime("%Y-%m-%d")
        elif hasattr(self, 'drag_start_scene_pos'):
            selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttBarItem)]
            if self not in selected_items:
                selected_items.append(self)
                
            moves = []
            for it in selected_items:
                isx = round(it.pos().x() / snap) * snap
                isw = max(snap, round(it.rect().width() / snap) * snap)
                isd = self.app.min_date + timedelta(days=isx / self.app.day_width)
                ied = isd + timedelta(days=isw / self.app.day_width - 0.001)
                
                new_row = int(it.pos().y() / self.app.row_height)
                moves.append({
                    'item': it,
                    'start_date': isd.strftime("%Y-%m-%d"),
                    'end_date': ied.strftime("%Y-%m-%d"),
                    'new_row': new_row,
                    'task': it.task,
                    'period_idx': it.period_index
                })
            
            moves.sort(key=lambda x: (x['task'] is self.task, x['period_idx']), reverse=True)
            
            for m in moves:
                task = m['task']
                if task.get('mode') == 'auto':
                    task['auto_start_date'] = m['start_date']
                    if 'periods' in task and len(task['periods']) > 0 and 0 <= m['period_idx'] < len(task['periods']):
                        p = task['periods'].pop(m['period_idx'])
                        p['start_date'] = m['start_date']
                        m['period_data'] = p
                    else:
                        m['period_data'] = {'start_date': m['start_date']}
                else:
                    if 'periods' in task and 0 <= m['period_idx'] < len(task['periods']):
                        p = task['periods'].pop(m['period_idx'])
                        p['start_date'] = m['start_date']
                        p['end_date'] = m['end_date']
                        m['period_data'] = p
                    else:
                        m['period_data'] = {'start_date': m['start_date'], 'end_date': m['end_date']}
            
            for m in moves:
                target_idx = m['new_row']
                if target_idx < len(self.app.visible_tasks_info):
                    target_task = self.app.visible_tasks_info[target_idx]['task']
                    source_mode = m['task'].get('mode', 'manual')
                    target_mode = target_task.get('mode', 'manual')
                    
                    if target_task.get('is_group') or source_mode != target_mode:
                        m['task'].setdefault('periods', []).append(m['period_data'])
                    else:
                        target_task.setdefault('periods', []).append(m['period_data'])
                else:
                    m['task'].setdefault('periods', []).append(m['period_data'])

        if hasattr(self, 'drag_start_scene_pos'):
            del self.drag_start_scene_pos
            
        super().mouseReleaseEvent(event)
        
        selected_period_dicts = []
        selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttBarItem)]
        if self not in selected_items:
            selected_items.append(self)
        for it in selected_items:
            if hasattr(it, 'period_dict'):
                selected_period_dicts.append(it.period_dict)

        def finalize_ui():
            self.app.recalculate_auto_tasks()
            self.app.sync_table_from_tasks()
            self.app.update_ui()
            if selected_period_dicts:
                for item in self.app.cs.items():
                    if isinstance(item, GanttBarItem) and hasattr(item, 'period_dict'):
                        if item.period_dict in selected_period_dicts:
                            item.setSelected(True)
            self.app.save_state_if_changed()

        QTimer.singleShot(0, finalize_ui)

    def mouseDoubleClickEvent(self, event):
        super().mouseDoubleClickEvent(event)
        
        if self.task.get('mode') == 'auto':
            return
            
        if 'periods' not in self.task:
            self.task['periods'] = [{'start_date': self.task.get('start_date', ''), 'end_date': self.task.get('end_date', '')}]
            
        if 0 <= self.period_index < len(self.task['periods']):
            p_dict = self.task['periods'][self.period_index]
            current_text = p_dict.get('text', '')
            
            text, ok = QInputDialog.getText(self.app, "テキストの編集", "バーに表示するテキスト:", QLineEdit.Normal, current_text)
            if ok:
                p_dict['text'] = text
                self.update_appearance()
                self.app.update_ui()
                self.app.save_state_if_changed()

    def contextMenuEvent(self, event):
        menu = QMenu()
        color_action = menu.addAction("色を変更")
        
        copy_action = None
        cut_action = None
        if self.task.get('mode') in ['manual', 'memo']:
            copy_action = menu.addAction("コピー")
            cut_action = menu.addAction("切り取り")
            
        del_action = menu.addAction("この期間を削除")

        selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttBarItem) and it.task.get('mode') in ['manual', 'memo']]
        if self not in selected_items:
            selected_items.append(self)
            
        pre_copied_periods = []
        if selected_items:
            min_row = min(it.row for it in selected_items)
            for it in selected_items:
                if hasattr(it, 'period_dict'):
                    p_copy = {}
                    for k in ['start_date', 'end_date', 'color', 'text']:
                        if k in it.period_dict:
                            p_copy[k] = it.period_dict[k]
                    p_copy['row_offset'] = it.row - min_row
                    pre_copied_periods.append({
                        'task': it.task,
                        'period_dict': it.period_dict,
                        'copy_data': p_copy
                    })

        action = menu.exec(event.screenPos())
        
        if action == color_action:
            color_groups = self.app.get_color_groups()
            dlg = ColorGridDialog(color_groups, self.app)
            if dlg.exec():
                if 'periods' not in self.task:
                    self.task['periods'] = [{'start_date': self.task.get('start_date', ''), 'end_date': self.task.get('end_date', '')}]
                self.task['periods'][self.period_index]['color'] = dlg.selected_color
                self.app.update_ui()
                self.app.save_state_if_changed()
        elif copy_action and action == copy_action:
            self.app.clipboard_periods = [p['copy_data'] for p in pre_copied_periods]
        elif cut_action and action == cut_action:
            self.app.clipboard_periods = [p['copy_data'] for p in pre_copied_periods]
            for p in pre_copied_periods:
                task = p['task']
                period_dict = p['period_dict']
                if 'periods' in task:
                    try:
                        task['periods'].remove(period_dict)
                    except ValueError:
                        pass
            QTimer.singleShot(0, lambda: (self.app.update_ui(), self.app.save_state_if_changed()))
        elif action == del_action:
            if 'periods' in self.task:
                try:
                    self.task['periods'].pop(self.period_index)
                    QTimer.singleShot(0, lambda: (self.app.update_ui(), self.app.save_state_if_changed()))
                except IndexError:
                    pass


class GanttCommentItem(QGraphicsRectItem):
    def __init__(self, task, row, comment_index, gantt_app, rect=None):
        super().__init__(rect)
        self.task = task
        self.row = row
        self.comment_index = comment_index
        self.app = gantt_app
        self.setFlags(QGraphicsRectItem.ItemIsMovable | 
                      QGraphicsRectItem.ItemIsSelectable | 
                      QGraphicsRectItem.ItemSendsGeometryChanges)
        self.setAcceptHoverEvents(True)
        self.setPen(QPen(Qt.NoPen))
        self.setBrush(QBrush(Qt.transparent))
        
        self.text_item = QGraphicsTextItem('', self)
        font = QFont("Segoe UI", 10)
        self.text_item.setFont(font)
        
        self.update_appearance()

    def update_appearance(self):
        c_dict = self.task.get('comments', [])[self.comment_index]
        self.text_item.setPlainText(c_dict.get('text', '📝 コメント'))
        
        color_code = c_dict.get('color', '#555555')
        self.text_item.setDefaultTextColor(QColor(color_code))
        
        br = self.text_item.boundingRect()
        self.setRect(0, 0, br.width() + 4, self.app.row_height - 10)
        self.text_item.setPos(2, (self.rect().height() - br.height()) / 2)
        
        if self.isSelected():
            self.setPen(QPen(QColor(0, 120, 212), 1, Qt.DashLine))
            self.setBrush(QBrush(QColor(0, 120, 212, 20)))
        else:
            self.setPen(QPen(Qt.NoPen))
            self.setBrush(QBrush(Qt.transparent))

    def hoverMoveEvent(self, event):
        self.setCursor(Qt.OpenHandCursor)
        super().hoverMoveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.setCursor(Qt.ClosedHandCursor)
            self.drag_start_pos = self.pos()
            self.drag_start_scene_pos = event.scenePos()
            self.drag_start_row = self.row
            
            selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttCommentItem)]
            if self not in selected_items:
                selected_items.append(self)
            for it in selected_items:
                it.drag_start_pos = it.pos()
                it.drag_start_row = it.row
                
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if hasattr(self, 'drag_start_scene_pos'):
            snap_x = self.app.day_width
            snap_y = self.app.row_height
            
            delta = event.scenePos() - self.drag_start_scene_pos
            dx = round(delta.x() / snap_x) * snap_x
            dy = round(delta.y() / snap_y) * snap_y
            
            selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttCommentItem)]
            if self not in selected_items:
                selected_items.append(self)
                
            for it in selected_items:
                if hasattr(it, 'drag_start_pos'):
                    new_row = it.drag_start_row + int(dy / snap_y)
                    max_row = len(self.app.visible_tasks_info) - 1 if self.app.visible_tasks_info else 0
                    new_row = max(0, min(max_row, new_row))
                    
                    target_task = self.app.visible_tasks_info[new_row]['task']
                    
                    if target_task.get('is_group'):
                        new_row = it.drag_start_row
                        
                    it.setPos(it.drag_start_pos.x() + dx, new_row * snap_y + 5)
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self.setCursor(Qt.OpenHandCursor)
        
        if hasattr(self, 'drag_start_scene_pos'):
            snap_x = self.app.day_width
            snap_y = self.app.row_height
            
            selected_items = [it for it in self.scene().selectedItems() if isinstance(it, GanttCommentItem)]
            if self not in selected_items:
                selected_items.append(self)
                
            moves = []
            for it in selected_items:
                sx = round(it.pos().x() / snap_x) * snap_x
                sd = self.app.min_date + timedelta(days=sx / snap_x)
                new_row = int((it.pos().y() - 5) / snap_y)
                
                moves.append({
                    'item': it,
                    'date': sd.strftime("%Y-%m-%d"),
                    'new_row': new_row,
                    'task': it.task,
                    'comment_idx': it.comment_index
                })
                
            moves.sort(key=lambda x: (x['task'] is self.task, x['comment_idx']), reverse=True)
            
            for m in moves:
                c_dict = m['task']['comments'].pop(m['comment_idx'])
                c_dict['date'] = m['date']
                m['comment_data'] = c_dict
                
            for m in moves:
                target_idx = m['new_row']
                if target_idx < len(self.app.visible_tasks_info):
                    target_task = self.app.visible_tasks_info[target_idx]['task']
                    if target_task.get('is_group'):
                        m['task'].setdefault('comments', []).append(m['comment_data'])
                    else:
                        target_task.setdefault('comments', []).append(m['comment_data'])
                else:
                    m['task'].setdefault('comments', []).append(m['comment_data'])

            for it in selected_items:
                if hasattr(it, 'drag_start_scene_pos'):
                    del it.drag_start_scene_pos
        
        super().mouseReleaseEvent(event)
        QTimer.singleShot(0, lambda: (self.app.update_ui(), self.app.save_state_if_changed()))

    def mouseDoubleClickEvent(self, event):
        super().mouseDoubleClickEvent(event)
        c_dict = self.task.get('comments', [])[self.comment_index]
        text, ok = QInputDialog.getText(self.app, "コメント編集", "コメント:", QLineEdit.Normal, c_dict.get('text', '📝 コメント'))
        if ok:
            c_dict['text'] = text
            self.app.update_ui()
            self.app.save_state_if_changed()

    def contextMenuEvent(self, event):
        menu = QMenu()
        color_action = menu.addAction("色を変更")
        del_action = menu.addAction("削除")
        action = menu.exec(event.screenPos())
        if action == color_action:
            color_groups = self.app.get_color_groups()
            dlg = ColorGridDialog(color_groups, self.app)
            if dlg.exec():
                self.task['comments'][self.comment_index]['color'] = dlg.selected_color
                self.app.update_ui()
                self.app.save_state_if_changed()
        elif action == del_action:
            try:
                self.task['comments'].pop(self.comment_index)
                QTimer.singleShot(0, lambda: (self.app.update_ui(), self.app.save_state_if_changed()))
            except IndexError:
                pass


class HeaderScene(QGraphicsScene):
    def __init__(self, app):
        super().__init__()
        self.app = app

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            y = event.scenePos().y()
            if 35 <= y <= 70:
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
                    self.app.save_state_if_changed()
                    return
        super().mousePressEvent(event)

class ChartScene(QGraphicsScene):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.start_x = 0
        self.selection_rect = None
        self.selection_start = None

    def mousePressEvent(self, e):
        items = self.items(e.scenePos(), Qt.IntersectsItemShape, Qt.DescendingOrder, self.app.chart_view.transform())
        gantt_item = next((it for it in items if isinstance(it, (GanttBarItem, GanttCommentItem))), None)

        if not gantt_item and e.button() == Qt.LeftButton:
            if e.modifiers() & Qt.ShiftModifier:
                self.start_x = e.scenePos().x()
            else:
                self.selection_start = e.scenePos()
                self.selection_rect = self.addRect(QRectF(self.selection_start, self.selection_start), 
                                                 QPen(QColor(0, 120, 212), 1, Qt.DashLine), 
                                                 QBrush(QColor(0, 120, 212, 40)))
                self.selection_rect.setZValue(100)
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        if self.selection_rect:
            rect = QRectF(self.selection_start, e.scenePos()).normalized()
            self.selection_rect.setRect(rect)
            return
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        if self.selection_rect:
            rect = self.selection_rect.rect()
            self.clearSelection()
            for item in self.items(rect):
                if isinstance(item, GanttBarItem):
                    if item.task.get('mode') != 'auto':
                        item.setSelected(True)
                elif isinstance(item, GanttCommentItem):
                    item.setSelected(True)
            
            self.removeItem(self.selection_rect)
            self.selection_rect = None
            self.selection_start = None
            return

        if self.start_x > 0:
            items = self.items(e.scenePos(), Qt.IntersectsItemShape, Qt.DescendingOrder, self.app.chart_view.transform())
            gantt_item = next((it for it in items if isinstance(it, (GanttBarItem, GanttCommentItem))), None)
            if not gantt_item:
                if abs(e.scenePos().x() - self.start_x) > (self.app.day_width * 0.1):
                    self.app.create_task_from_drag(self.start_x, e.scenePos().x(), e.scenePos().y())
            self.start_x = 0
        super().mouseReleaseEvent(e)

    def mouseDoubleClickEvent(self, e):
        items = self.items(e.scenePos(), Qt.IntersectsItemShape, Qt.DescendingOrder, self.app.chart_view.transform())
        gantt_item = next((it for it in items if isinstance(it, (GanttBarItem, GanttCommentItem))), None)
            
        if not gantt_item and e.button() == Qt.LeftButton:
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
                
                if task.get('mode') == 'auto':
                    task['auto_start_date'] = d_str
                    self.app.recalculate_auto_tasks()
                else:
                    if 'periods' not in task:
                        task['periods'] = [{'start_date': task.get('start_date', ''), 'end_date': task.get('end_date', '')}]
                    task['periods'].append({"start_date": d_str, "end_date": d_str, "text": ""})
                self.app.update_ui()
                e.accept()
                self.app.save_state_if_changed()
                return
        super().mouseDoubleClickEvent(e)

    def contextMenuEvent(self, e):
        item = self.itemAt(e.scenePos(), self.app.chart_view.transform())
        if item and not isinstance(item, (GanttBarItem, GanttCommentItem)):
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
                    add_comment_action = None
                    paste_action = None
                else:
                    task_name = task.get('name', '無題')
                    if task.get('mode') == 'auto':
                        add_period_action = menu.addAction(f"「{task_name}」の開始日をここに設定")
                        paste_action = None
                    else:
                        add_period_action = menu.addAction(f"「{task_name}」に期間を追加")
                        paste_action = None
                        if hasattr(self.app, 'clipboard_periods') and self.app.clipboard_periods:
                            paste_action = menu.addAction("貼り付け")
                    add_comment_action = menu.addAction(f"「{task_name}」にコメントを追加")
                    add_task_in_group = None
            else:
                add_period_action = None
                add_comment_action = None
                add_task_in_group = None
                paste_action = None
            
            action = menu.exec(e.screenPos())
            x = e.scenePos().x()
            day_idx = int(x / self.app.day_width)
            d_str = (self.app.min_date + timedelta(days=day_idx)).strftime("%Y-%m-%d")
            
            if action == add_period_action and add_period_action:
                if task.get('mode') == 'auto':
                    task['auto_start_date'] = d_str
                    self.app.recalculate_auto_tasks()
                else:
                    if 'periods' not in task:
                        task['periods'] = [{'start_date': task.get('start_date', ''), 'end_date': task.get('end_date', '')}]
                    task['periods'].append({"start_date": d_str, "end_date": d_str})
                self.app.update_ui()
                self.app.save_state_if_changed()
            elif paste_action and action == paste_action:
                try:
                    target_d = datetime.strptime(d_str, "%Y-%m-%d")
                    first_item_sd = datetime.strptime(self.app.clipboard_periods[0]['start_date'], "%Y-%m-%d")
                    delta_days = (target_d - first_item_sd).days
                except (ValueError, KeyError, IndexError):
                    delta_days = 0

                for p_data in self.app.clipboard_periods:
                    new_p = p_data.copy()
                    row_offset = new_p.pop('row_offset', 0)
                    target_row = row + row_offset
                    
                    if target_row < len(self.app.visible_tasks_info):
                        target_task = self.app.visible_tasks_info[target_row]['task']
                        if target_task.get('is_group') or target_task.get('mode') == 'auto':
                            continue
                            
                        if 'periods' not in target_task:
                            target_task['periods'] = []
                            
                        try:
                            orig_sd = datetime.strptime(new_p['start_date'], "%Y-%m-%d")
                            orig_ed = datetime.strptime(new_p['end_date'], "%Y-%m-%d")
                            new_sd = orig_sd + timedelta(days=delta_days)
                            new_ed = orig_ed + timedelta(days=delta_days)
                            new_p['start_date'] = new_sd.strftime("%Y-%m-%d")
                            new_p['end_date'] = new_ed.strftime("%Y-%m-%d")
                        except ValueError:
                            pass
                        target_task['periods'].append(new_p)
                self.app.update_ui()
                self.app.save_state_if_changed()
            elif action == add_comment_action and add_comment_action:
                if 'comments' not in task:
                    task['comments'] = []
                task['comments'].append({"date": d_str, "text": "📝 コメント", "color": "#333333"})
                self.app.update_ui()
                self.app.save_state_if_changed()
            elif action == add_task_in_group and add_task_in_group:
                mode_idx = self.app.mode_combo.currentIndex()
                if mode_idx == 0: 
                    mode = "manual"
                    color = "#008000"
                elif mode_idx == 1: 
                    mode = "auto"
                    color = "#0000ff"
                else: 
                    mode = "memo"
                    color = "#808080"

                new_task = {
                    "name": "新規タスク",
                    "mode": mode,
                    "color": color,
                    "efficiency": 1.0
                }
                if mode == "auto":
                    new_task["auto_start_date"] = d_str
                    new_task["workload"] = 1.0 
                    new_task["periods"] = [{"start_date": d_str, "end_date": d_str, "color": color}]
                    new_task["headcount"] = 0.0
                elif mode == "memo":
                    new_task["periods"] = [{"start_date": d_str, "end_date": d_str, "color": color}]
                    new_task["headcount"] = 0.0
                else:
                    new_task["periods"] = [{"start_date": d_str, "end_date": d_str, "color": color}]
                    new_task["headcount"] = 1.0
                    
                self.app.tasks.insert(info['index'] + 1, new_task)
                self.app.recalculate_auto_tasks()
                self.app.update_ui()
                self.app.save_state_if_changed()
        else:
            super().contextMenuEvent(e)