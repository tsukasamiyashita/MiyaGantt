# task_manager.py
from datetime import datetime, timedelta
from PySide6.QtWidgets import QMessageBox

class TaskManagerMixin:
    def parse_date(self, s):
        s = s.strip().replace('/', '-')
        parts = s.split('-')
        now = datetime.now()
        if len(parts) == 3:
            try:
                return datetime.strptime(s, "%Y-%m-%d").strftime("%Y-%m-%d")
            except ValueError:
                return None
        elif len(parts) == 2:
            try:
                return f"{now.year}-{int(parts[0]):02d}-{int(parts[1]):02d}"
            except ValueError:
                return None
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
        
        info = getattr(self, 'visible_tasks_info', [])[row]
        task_to_track = info['task']
        
        if task_to_track.get('is_group'):
            target_v_row = row - 1
            while target_v_row > 0 and not self.visible_tasks_info[target_v_row]['task'].get('is_group'):
                target_v_row -= 1
            self.move_tasks([row], target_v_row)
        else:
            self.move_tasks([row], row - 1)
            
        for i, n_info in enumerate(self.visible_tasks_info):
            if n_info['task'] == task_to_track:
                self.table.setCurrentCell(i, 2)
                break
        self.save_state_if_changed()

    def move_row_down(self):
        row = self.table.currentRow()
        visible_info = getattr(self, 'visible_tasks_info', [])
        if row < 0 or row >= len(visible_info) - 1: return
        
        info = visible_info[row]
        task_to_track = info['task']
        
        if task_to_track.get('is_group'):
            current_group_end_v = row + 1
            end_idx = info['index'] + 1
            for i in range(info['index'] + 1, len(self.tasks)):
                if self.tasks[i].get('is_group'): break
                end_idx = i + 1
            for i in range(row + 1, len(visible_info)):
                if visible_info[i]['index'] >= end_idx: break
                current_group_end_v = i + 1
            
            next_group_v = -1
            for i in range(current_group_end_v, len(visible_info)):
                if visible_info[i]['task'].get('is_group'):
                    next_group_v = i
                    break
            
            if next_group_v == -1:
                target_v_row = len(visible_info)
            else:
                target_v_row = next_group_v + 1
                ng_end_idx = visible_info[next_group_v]['index'] + 1
                for i in range(visible_info[next_group_v]['index'] + 1, len(self.tasks)):
                    if self.tasks[i].get('is_group'): break
                    ng_end_idx = i + 1
                for i in range(next_group_v + 1, len(visible_info)):
                    if visible_info[i]['index'] >= ng_end_idx: break
                    target_v_row = i + 1
            
            self.move_tasks([row], target_v_row)
        else:
            self.move_tasks([row], row + 2)
            
        for i, n_info in enumerate(self.visible_tasks_info):
            if n_info['task'] == task_to_track:
                self.table.setCurrentCell(i, 2)
                break
        self.save_state_if_changed()

    def move_tasks(self, source_rows, target_row, refresh_chart=True):
        if not source_rows: return
        
        src_v_row = source_rows[0]
        visible_info = getattr(self, 'visible_tasks_info', [])
        if src_v_row >= len(visible_info): return
        
        info = visible_info[src_v_row]
        src_idx = info['index']
        is_group = info['task'].get('is_group', False)
        
        start_idx = src_idx
        end_idx = src_idx + 1
        if is_group:
            for i in range(src_idx + 1, len(self.tasks)):
                if self.tasks[i].get('is_group'):
                    break
                end_idx = i + 1
        
        block = self.tasks[start_idx:end_idx]
        
        visible_count = 1
        if is_group and not info['task'].get('collapsed'):
            for i in range(src_v_row + 1, len(visible_info)):
                if visible_info[i]['index'] >= end_idx:
                    break
                visible_count += 1
        
        remaining_tasks = self.tasks[:start_idx] + self.tasks[end_idx:]
        
        target_task = None
        if target_row < len(visible_info):
            target_task = visible_info[target_row]['task']
            
        if target_task in block:
            return

        if target_task is None:
            new_target_idx = len(remaining_tasks)
        else:
            try:
                new_target_idx = remaining_tasks.index(target_task)
            except ValueError:
                new_target_idx = len(remaining_tasks)
                
        if is_group:
            while new_target_idx < len(remaining_tasks):
                if remaining_tasks[new_target_idx].get('is_group'):
                    break
                new_target_idx += 1
            
        remaining_tasks[new_target_idx:new_target_idx] = block
        self.tasks = remaining_tasks
        
        self.recalculate_auto_tasks()
        self.update_ui(refresh_chart)
        self.update_selection_mark()
        self.save_state_if_changed()

    def recalculate_auto_tasks(self):
        groups_to_process = []
        temp_group = None
        temp_tasks = []
        for t in self.tasks:
            if t.get('is_group'):
                if temp_group is not None or temp_tasks:
                    groups_to_process.append((temp_group, temp_tasks))
                temp_group = t
                temp_tasks = []
            else:
                temp_tasks.append(t)
        if temp_group is not None or temp_tasks:
            groups_to_process.append((temp_group, temp_tasks))
            
        for g, tasks in groups_to_process:
            daily_speed = {}
            for t in tasks:
                if t.get('mode', 'manual') == 'manual':
                    hc = float(t.get('headcount', 1.0)) * float(t.get('efficiency', 1.0))
                    t_color = t.get('color', '#808080')
                    for p in t.get('periods', []):
                        if not p.get('start_date') or not p.get('end_date'): continue
                        
                        p_color = p.get('color')
                        if p_color and p_color.lower() != t_color.lower():
                            continue
                            
                        try:
                            sd = datetime.strptime(p['start_date'], "%Y-%m-%d")
                            ed = datetime.strptime(p['end_date'], "%Y-%m-%d")
                            for i in range((ed - sd).days + 1):
                                d_str = (sd + timedelta(days=i)).strftime("%Y-%m-%d")
                                daily_speed[d_str] = daily_speed.get(d_str, 0.0) + hc
                        except ValueError:
                            pass
            
            auto_tasks = []
            for t in tasks:
                if t.get('mode') == 'auto':
                    sd_str = t.get('auto_start_date')
                    if not sd_str:
                        sd_str = getattr(self, 'min_date', datetime.now()).strftime("%Y-%m-%d")
                        t['auto_start_date'] = sd_str
                    try:
                        start_date = datetime.strptime(sd_str, "%Y-%m-%d")
                    except ValueError:
                        start_date = getattr(self, 'min_date', datetime.now())
                        
                    auto_tasks.append({
                        'task': t,
                        'start': start_date,
                        'rem_work': float(t.get('workload', 1.0)), 
                        'end': None,
                        'last_progress': None,
                        'daily_allocations': {}
                    })
                    
            if not auto_tasks:
                continue
                
            current_date = min(at['start'] for at in auto_tasks)
            max_days = 3650
            days_simulated = 0
            
            has_speed = any(s > 0 for s in daily_speed.values())
            has_custom = any(t.get('custom_allocations') for t in tasks if t.get('mode') == 'auto')
            
            if not has_speed and not has_custom:
                for at in auto_tasks:
                    t = at['task']
                    sd_str = at['start'].strftime("%Y-%m-%d")
                    rem = at['rem_work']
                    t['periods'] = [{"start_date": sd_str, "end_date": sd_str, "color": "#d13438", "text": f"⚠️ 進行不可 (不足: {rem:g}工数)"}]
                continue
                
            max_speed_date_str = max(daily_speed.keys()) if daily_speed else "2000-01-01"
            custom_dates = []
            for at in auto_tasks:
                custom_dates.extend(at['task'].get('custom_allocations', {}).keys())
            max_custom_date_str = max(custom_dates) if custom_dates else "2000-01-01"
            max_date_str = max(max_speed_date_str, max_custom_date_str)
            
            try:
                max_sim_date = datetime.strptime(max_date_str, "%Y-%m-%d")
            except ValueError:
                max_sim_date = current_date + timedelta(days=365)
            
            while any(at['rem_work'] > 0 for at in auto_tasks) and days_simulated < max_days:
                if current_date > max_sim_date:
                    break
                    
                d_str = current_date.strftime("%Y-%m-%d")
                
                avail_res = daily_speed.get(d_str, 0.0)
                active_tasks = [at for at in auto_tasks if at['start'] <= current_date and at['rem_work'] > 0]
                
                for at in active_tasks:
                    at['daily_assigned'] = 0.0
                    at['is_custom_today'] = False
                
                for at in active_tasks:
                    custom_allocs = at['task'].get('custom_allocations', {})
                    if d_str in custom_allocs:
                        c_val = max(0.0, float(custom_allocs[d_str]))
                        
                        at['rem_work'] -= c_val
                        avail_res -= c_val
                        at['daily_assigned'] += c_val
                        if c_val > 0.0001:
                            at['last_progress'] = current_date
                        at['daily_allocations'][d_str] = c_val
                        at['is_custom_today'] = True
                            
                        if at['rem_work'] <= 0.001:
                            at['rem_work'] = 0.0
                            at['end'] = current_date
                
                avail_res = max(0.0, avail_res)

                tasks_to_allocate = [at for at in active_tasks if at['rem_work'] > 0 and not at['is_custom_today']]
                if tasks_to_allocate and avail_res > 0.001:
                    unallocated_res = avail_res
                    
                    while tasks_to_allocate and unallocated_res > 0.001:
                        alloc_per_task = unallocated_res / len(tasks_to_allocate)
                        next_tasks = []
                        allocated_anything = False
                        
                        for at in tasks_to_allocate:
                            cap = at['task'].get('headcount', 0.0)
                            max_receivable = at['rem_work']
                            if cap > 0.001:
                                max_receivable = min(max_receivable, cap - at['daily_assigned'])
                            
                            if max_receivable < 0.001:
                                continue
                                
                            amount_to_give = min(alloc_per_task, max_receivable)
                            
                            if amount_to_give > 0.0001:
                                allocated_anything = True
                                at['rem_work'] -= amount_to_give
                                unallocated_res -= amount_to_give
                                at['daily_assigned'] += amount_to_give
                                at['last_progress'] = current_date
                                at['daily_allocations'][d_str] = at['daily_allocations'].get(d_str, 0.0) + amount_to_give
                            
                            if at['rem_work'] <= 0.001:
                                at['rem_work'] = 0.0
                                at['end'] = current_date
                            else:
                                if cap <= 0.001 or at['daily_assigned'] + 0.001 < cap:
                                    next_tasks.append(at)
                        
                        if not allocated_anything:
                            break
                        tasks_to_allocate = next_tasks
                
                current_date += timedelta(days=1)
                days_simulated += 1
                
            for at in auto_tasks:
                t = at['task']
                sd_str = at['start'].strftime("%Y-%m-%d")
                
                if at['rem_work'] > 0:
                    if at['last_progress']:
                        ed_str = at['last_progress'].strftime("%Y-%m-%d")
                    else:
                        ed_str = sd_str
                    p_color = "#d13438"
                    p_text = f"⚠️ キャパオーバー (不足: {at['rem_work']:g}工数)"
                else:
                    ed_str = at['end'].strftime("%Y-%m-%d") if at['end'] else sd_str
                    p_color = t.get('color')
                    p_text = ""
                    if t.get('periods') and len(t['periods']) > 0:
                        prev_color = t['periods'][0].get('color', p_color)
                        prev_text = t['periods'][0].get('text', "")
                        if prev_text and ("⚠️ キャパオーバー" in prev_text or "⚠️ 進行不可" in prev_text):
                            p_color = t.get('color')
                            p_text = ""
                        else:
                            p_color = prev_color
                            p_text = prev_text
                
                t['periods'] = [{"start_date": sd_str, "end_date": ed_str, "color": p_color, "text": p_text}]
                t['daily_allocations'] = at.get('daily_allocations', {})

    def add_task(self):
        mode_idx = self.mode_combo.currentIndex()
        if mode_idx == 0: 
            mode = "manual"
            color = "#808080"
        elif mode_idx == 1: 
            mode = "auto"
            color = "#323130"
        elif mode_idx == 2: 
            mode = "memo"
            color = "#c0c0c0"
        else:
            mode = "heading"
            color = "#4169e1"
        
        t = {
            "name": f"新規タスク {len(self.tasks)+1}",
            "mode": mode,
            "color": color,
            "efficiency": 1.0
        }
        
        if mode == "auto":
            t["auto_start_date"] = self.min_date.strftime("%Y-%m-%d")
            t["workload"] = 1.0 
            t["periods"] = [{"start_date": t["auto_start_date"], "end_date": t["auto_start_date"], "color": color}]
            t["headcount"] = 0.0
        elif mode in ["memo", "heading"]:
            t["periods"] = []
            t["headcount"] = 0.0
        else:
            t["periods"] = []
            t["headcount"] = 1.0
            
        r = self.table.currentRow()
        visible_info = getattr(self, 'visible_tasks_info', [])
        if r >= 0 and r < len(visible_info):
            idx = visible_info[r]['index']
            self.tasks.insert(idx + 1, t)
        else:
            self.tasks.append(t)
            
        self.recalculate_auto_tasks()
        self.update_ui()
        
        for i, info in enumerate(getattr(self, 'visible_tasks_info', [])):
            if info['task'] is t:
                self.table.setCurrentCell(i, 2)
                break
        self.save_state_if_changed()

    def add_group(self):
        g = {
            "name": f"新規グループ {len(self.tasks)+1}",
            "is_group": True,
            "collapsed": False,
            "headcount": 1.0,
            "efficiency": 1.0,
            "color": "#555555"
        }
        r = self.table.currentRow()
        visible_info = getattr(self, 'visible_tasks_info', [])
        if r >= 0 and r < len(visible_info):
            idx = visible_info[r]['index']
            self.tasks.insert(idx + 1, g)
        else:
            self.tasks.append(g)
            
        self.update_ui()
        
        for i, info in enumerate(getattr(self, 'visible_tasks_info', [])):
            if info['task'] is g:
                self.table.setCurrentCell(i, 2)
                break
        self.save_state_if_changed()

    def delete_task(self):
        r = self.table.currentRow()
        visible_info = getattr(self, 'visible_tasks_info', [])
        if r >= 0 and r < len(visible_info):
            if QMessageBox.question(self, "確認", "削除しますか？") == QMessageBox.Yes:
                idx = visible_info[r]['index']
                self.tasks.pop(idx)
                self.recalculate_auto_tasks()
                self.update_ui()
                self.save_state_if_changed()

    def create_task_from_drag(self, x1, x2, y):
        snap = self.day_width
        sx = round(min(x1, x2) / snap) * snap
        ex = round(max(x1, x2) / snap) * snap
        if sx == ex: ex += snap
        sd = getattr(self, 'min_date', datetime.now()) + timedelta(days=sx/snap)
        ed = getattr(self, 'min_date', datetime.now()) + timedelta(days=ex/snap - 0.001)
        row = max(0, int(y / getattr(self, 'row_height', 40)))
        
        mode_idx = self.mode_combo.currentIndex()
        if mode_idx == 0: 
            mode = "manual"
            color = "#808080"
        elif mode_idx == 1: 
            mode = "auto"
            color = "#323130"
        elif mode_idx == 2: 
            mode = "memo"
            color = "#c0c0c0"
        else:
            mode = "heading"
            color = "#4169e1"
        
        t = {
            "name": f"新規 {len(self.tasks)+1}", 
            "mode": mode,
            "color": color,
            "efficiency": 1.0
        }
        
        if mode == "auto":
            t["auto_start_date"] = sd.strftime("%Y-%m-%d")
            t["workload"] = 1.0 
            t["periods"] = [{"start_date": sd.strftime("%Y-%m-%d"), "end_date": sd.strftime("%Y-%m-%d"), "color": color}]
            t["headcount"] = 0.0
        elif mode in ["memo", "heading"]:
            t["periods"] = [{"start_date": sd.strftime("%Y-%m-%d"), "end_date": ed.strftime("%Y-%m-%d"), "color": color}]
            t["headcount"] = 0.0
        else:
            t["periods"] = [{"start_date": sd.strftime("%Y-%m-%d"), "end_date": ed.strftime("%Y-%m-%d"), "color": color}]
            t["headcount"] = 1.0
            
        visible_info = getattr(self, 'visible_tasks_info', [])
        if row < len(visible_info):
            insert_idx = visible_info[row]['index']
            self.tasks.insert(insert_idx, t)
        else:
            self.tasks.append(t)
            
        self.recalculate_auto_tasks()
        self.update_ui()
        self.save_state_if_changed()

    def get_visible_tasks_info(self):
        visible = []
        skip_until_next_group = False
        for i, t in enumerate(self.tasks):
            if t.get('is_group'):
                visible.append({'index': i, 'task': t, 'indent': 0})
                skip_until_next_group = t.get('collapsed', False)
            else:
                if not skip_until_next_group:
                    has_group = any(self.tasks[j].get('is_group') for j in range(i))
                    visible.append({'index': i, 'task': t, 'indent': 1 if has_group else 0})
        return visible