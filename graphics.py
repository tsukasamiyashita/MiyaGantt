# tsukasamiyashita/miyagantt/MiyaGantt-90775b445eeca08d321c122853c84ad8762e2c95/graphics.py
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
        # 自分がセットされているのが header_view であることを確認
        if hasattr(self.app, 'hv') and event.widget() != self.app.hv.viewport():
            return

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
                    
                    is_public = jpholiday.is_holiday(d)
                    is_default_holiday = d.weekday() in (5, 6) or is_public
                    
                    if d_str in self.app.custom_holidays:
                        # 既にカスタム設定がある場合は元に戻す
                        del self.app.custom_holidays[d_str]
                    else:
                        # カスタム設定がない場合、デフォルトが休日なら「営業日」に、平日なら「休日」にする
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

        # 2. ドラッグモードの制御（背景クリックでの矩形選択による色変化を防ぐため常にNoDrag）
        self.app.chart_view.setDragMode(QGraphicsView.DragMode.NoDrag)

        # 3. 背景クリック（バー以外）の場合
        if not target_bar:
            # 背景クリック時も対応するテーブル行を選択状態にする（同期）
            if e.button() == Qt.LeftButton:
                y = e.scenePos().y()
                row = int(y / self.app.row_height)
                if 0 <= row < len(self.app.visible_tasks_info):
                    self.app.table.setCurrentCell(row, 2)

                if e.modifiers() & Qt.AltModifier:
                    self.start_x = e.scenePos().x()
                else:
                    self.start_x = 0
            
            # 重要：背景クリック時は標準のイベント（super）を呼ばず、ここで処理を終了する
            # これにより、背景色の一時的な変化やチラつき、不要な選択解除を完全に防ぐ
            e.accept()
            return

        # 4. すでに選択されているバーを左クリックした場合の選択維持（ドラッグ準備）
        if target_bar and e.button() == Qt.LeftButton and not (e.modifiers() & (Qt.ControlModifier | Qt.ShiftModifier)):
            if target_bar.isSelected():
                # 重要：現在選択されている全てのアイテムを記録し、
                # super().mousePressEvent(e) による他アイテムの選択解除を防ぐために後で復元する
                selected_items = self.selectedItems()
                super().mousePressEvent(e)
                for it in selected_items:
                    it.setSelected(True)
                return

        # 5. 右クリック処理
        if target_bar and e.button() == Qt.RightButton:
            target_bar.mousePressEvent(e)
            e.accept()
            return
        
        # 6. 未選択のバーをクリックした場合などの通常処理
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