import os
import json
import calendar
import copy
from datetime import datetime, timedelta
import jpholiday
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *
from dialogs import SettingsDialog, SummaryDialog, HelpDialog
from graphics import GanttBarItem

class EventsMixin:
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

