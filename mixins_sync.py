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

class SyncMixin:
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
                        color = task.get('color', '#0078d4')
                        # 人数を考慮して加算
                        day_map[color] = day_map.get(color, 0) + overlap * task.get('person_count', 1)
                except ValueError:
                    continue
        return day_map

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

    def update_selection_mark(self, *args):
        if getattr(self, '_updating_selection', False):
            return
        self._updating_selection = True
        try:
            self.table.blockSignals(True)
            curr = self.table.currentRow()
            
            # 前回の選択行と今回の選択行のみを更新して高速化
            last_row = getattr(self, 'last_current_row', -1)
            
            # 全行を回すのではなく、必要な行のみ更新
            rows_to_update = set()
            if last_row >= 0 and last_row < self.table.rowCount():
                rows_to_update.add(last_row)
            if curr >= 0 and curr < self.table.rowCount():
                rows_to_update.add(curr)
                
            if not rows_to_update:
                # 念のため全行チェック（初回など）
                for r in range(self.table.rowCount()):
                    it = self.table.item(r, 0)
                    if it:
                        text = "●" if r == curr else ""
                        if it.text() != text:
                            it.setText(text)
                            it.setTextAlignment(Qt.AlignCenter)
                            it.setForeground(QColor(0, 120, 212))
            else:
                for r in rows_to_update:
                    it = self.table.item(r, 0)
                    if it:
                        it.setText("●" if r == curr else "")
                        it.setTextAlignment(Qt.AlignCenter)
                        it.setForeground(QColor(0, 120, 212))
            
            self.last_current_row = curr
            self.table.blockSignals(False)
        finally:
            self._updating_selection = False

