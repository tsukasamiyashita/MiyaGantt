# tsukasamiyashita/miyagantt/MiyaGantt-46a1664b6d1737cb32f1dd17429ce06cca8dc678/chart_renderer.py
from datetime import datetime, timedelta
import jpholiday
from PySide6.QtCore import Qt, QRectF
from PySide6.QtGui import QBrush, QPen, QColor, QFont
from PySide6.QtWidgets import QTableWidgetItem
from gantt_items import GanttBarItem

class ChartRenderer:
    def __init__(self, app):
        self.app = app

    def draw_chart(self):
        self.app.month_label_items = []
        self.app.hs.clear()
        self.app.cs.clear()
        tw_total = self.app.display_days * self.app.day_width
        ch = max(self.app.chart_view.height(), len(self.app.tasks) * self.app.row_height)
        
        last_m = None
        for i in range(self.app.display_days):
            d = self.app.min_date + timedelta(days=i)
            x = i * self.app.day_width
            d_str = d.strftime("%Y-%m-%d")
            
            is_custom = d_str in self.app.custom_holidays
            is_public = jpholiday.is_holiday(d)
            
            bg = None
            if d.weekday() == 5:
                if is_custom:
                    bg = QColor(255, 240, 240)
                else:
                    bg = QColor(240, 248, 255)
            elif d.weekday() == 6 or is_public:
                if is_custom:
                    bg = None
                else:
                    bg = QColor(255, 240, 240)
            elif is_custom:
                bg = QColor(255, 240, 240)
            
            if bg:
                re = self.app.cs.addRect(x, 0, self.app.day_width, ch, QPen(Qt.NoPen), QBrush(bg))
                re.setZValue(-20)
                re.setAcceptedMouseButtons(Qt.NoButton)
            
            gl = self.app.cs.addLine(x, 0, x, ch, QPen(QColor(220, 220, 220), 1))
            gl.setZValue(-15)
            gl.setAcceptedMouseButtons(Qt.NoButton)
            
            if self.app.day_width >= 60:
                for h in [6, 12, 18]:
                    sl = self.app.cs.addLine(x + (self.app.day_width * h / 24.0), 0, x + (self.app.day_width * h / 24.0), ch, QPen(QColor(245, 245, 245), 0.5))
                    sl.setZValue(-15)
                    sl.setAcceptedMouseButtons(Qt.NoButton)
            
            h_bg = QColor(255, 255, 225) if is_custom else QColor(248, 248, 248)
            self.app.hs.addRect(x, 35, self.app.day_width, 35, QPen(QColor(210, 210, 210)), QBrush(h_bg)).setZValue(5)
            
            if self.app.day_width >= 35:
                dl = self.app.hs.addText(d.strftime("%d"))
                day_color = QColor(50, 50, 50)
                
                dl.setDefaultTextColor(day_color)
                dl.setFont(QFont("Segoe UI", 9, QFont.Bold))
                dl.setPos(x + (self.app.day_width/2) - 10, 35)
                dl.setZValue(10)
                
                if d.weekday() == 5 and not is_public:
                    w_c = QColor(0, 80, 200)
                elif d.weekday() == 6 or is_public:
                    w_c = QColor(220, 0, 0)
                else:
                    w_c = QColor(60, 60, 60)
                
                yl = self.app.hs.addText(["月","火","水","木","金","土","日"][d.weekday()])
                yl.setDefaultTextColor(w_c)
                yl.setFont(QFont("Segoe UI", 7))
                yl.setPos(x + (self.app.day_width/2) - 8, 52)
                yl.setZValue(10)
                
                if is_public or is_custom:
                    h_name = self.app.custom_holidays.get(d_str) or jpholiday.is_holiday_name(d)
                    if h_name:
                        dl.setToolTip(h_name)
                        yl.setToolTip(h_name)
                        header_bg = self.app.hs.addRect(x, 35, self.app.day_width, 35, QPen(Qt.NoPen), QBrush(Qt.transparent))
                        header_bg.setZValue(11)
                        header_bg.setToolTip(h_name)
            
            if (cm := d.strftime("%Y/%m")) != last_m:
                if self.app.month_label_items:
                    self.app.month_label_items[-1][1] = x
                self.app.hs.addLine(x, 0, x, 35, QPen(QColor(150, 150, 150), 2)).setZValue(15)
                ml = self.app.hs.addText(d.strftime("%Y年 %m月"))
                ml.setDefaultTextColor(QColor(0, 90, 180))
                ml.setFont(QFont("Segoe UI", 11, QFont.Bold))
                ml.setPos(x + 5, 5)
                ml.setZValue(25)
                self.app.month_label_items.append([x, x + self.app.day_width * 31, ml])
                last_m = cm

        if self.app.month_label_items:
            self.app.month_label_items[-1][1] = tw_total

        for r in range(len(self.app.visible_tasks_info) + 1):
            y = r * self.app.row_height
            self.app.cs.addLine(0, y, tw_total, y, QPen(QColor(220, 220, 220), 1)).setZValue(-15)
        
        self.app.hs.addRect(0, 0, tw_total, 35, QPen(Qt.NoPen), QBrush(QColor(235, 245, 255))).setZValue(0)
        
        nx = (datetime.now() - self.app.min_date).total_seconds() / (24*3600) * self.app.day_width
        if 0 <= nx < tw_total:
            self.app.cs.addLine(nx, 0, nx, ch, QPen(QColor(255, 60, 60), 2, Qt.DashLine)).setZValue(25)
            
        for row, info in enumerate(self.app.visible_tasks_info):
            t = info['task']
            try:
                if t.get('is_group'):
                    counts = [0] * self.app.display_days
                    for i in range(info['index'] + 1, len(self.app.tasks)):
                        sub_t = self.app.tasks[i]
                        if sub_t.get('is_group'): break
                        sub_periods = sub_t.get('periods', [])
                        t_color = sub_t.get('color', '#0078d4')
                        for p in sub_periods:
                            if not p.get('start_date') or not p.get('end_date'): continue
                            
                            p_color = p.get('color')
                            if p_color and p_color.lower() != t_color.lower():
                                continue

                            psd = datetime.strptime(p['start_date'], "%Y-%m-%d")
                            ped = datetime.strptime(p['end_date'], "%Y-%m-%d")
                            s_idx = max(0, (psd - self.app.min_date).days)
                            e_idx = min(self.app.display_days - 1, (ped - self.app.min_date).days)
                            
                            if sub_t.get('mode') == 'auto':
                                sub_hc = sub_t.get('workload', 10.0) / max(1, (ped - psd).days + 1)
                            else:
                                sub_hc = sub_t.get('headcount', 1.0)
                                
                            for d_idx in range(s_idx, e_idx + 1):
                                counts[d_idx] += sub_hc
                    
                    for d_idx, count in enumerate(counts):
                        if count > 0.001:
                            x = d_idx * self.app.day_width
                            r = min(self.app.day_width * 0.8, self.app.row_height * 0.6)
                            self.app.cs.addEllipse(x + (self.app.day_width - r)/2, row * self.app.row_height + (self.app.row_height - r)/2, r, r, 
                                               QPen(Qt.NoPen), QBrush(QColor(0, 120, 212, 40))).setZValue(15)
                            
                            txt = self.app.cs.addText(f"{count:g}")
                            txt.setFont(QFont("Segoe UI", 9, QFont.Bold))
                            txt.setDefaultTextColor(QColor(0, 120, 212))
                            tw = txt.boundingRect().width()
                            th = txt.boundingRect().height()
                            txt.setPos(x + (self.app.day_width - tw)/2, row * self.app.row_height + (self.app.row_height - th)/2)
                            txt.setZValue(20)
                    continue

                periods = t.get('periods')
                if periods is None:
                    if t.get('start_date') and t.get('end_date'):
                        periods = [{'start_date': t['start_date'], 'end_date': t['end_date']}]
                    else:
                        continue
                
                if not periods:
                    continue
                    
                for p_idx, p in enumerate(periods):
                    if not p.get('start_date') or not p.get('end_date'): continue
                    sd = datetime.strptime(p['start_date'], "%Y-%m-%d")
                    ed = datetime.strptime(p['end_date'], "%Y-%m-%d")
                    bar_w = ((ed - sd).days + 1) * self.app.day_width
                    bar = GanttBarItem(t, row, p_idx, self.app, QRectF(0, 0, bar_w, self.app.row_height - 20))
                    bar.setPos((sd - self.app.min_date).days * self.app.day_width, row * self.app.row_height + 10)
                    bar.setZValue(30)
                    self.app.cs.addItem(bar)
            except Exception as e:
                print(f"Error drawing bar for row {row}: {e}")
                
        self.app.hs.setSceneRect(0, 0, tw_total, self.app.header_height)
        self.app.cs.setSceneRect(0, 0, tw_total, ch)
        self.app.update_month_labels_pos()

    def update_ui(self, refresh_chart=True):
        self.app.visible_tasks_info = self.app.get_visible_tasks_info()
        self.app.table.blockSignals(True)
        
        scroll_val = self.app.chart_view.horizontalScrollBar().value()
        days_scrolled = scroll_val / self.app.day_width if self.app.day_width > 0 else 0
        visible_start = self.app.min_date + timedelta(days=days_scrolled)
        threshold_date = self.app.get_threshold_date(visible_start)
        
        headers = self.app.get_summary_headers(threshold_date)
        base_col_count = 8
        total_cols = base_col_count + len(headers)
        self.app.table.setColumnCount(total_cols)
        
        labels = ["", "", "タスク名", "モード", "人数/工数", "進捗(%)", "期間/開始日", "色"] + [h[2] for h in headers]
        self.app.table.setHorizontalHeaderLabels(labels)
        
        new_rows = len(self.app.visible_tasks_info)
        if self.app.table.rowCount() < new_rows:
            self.app.table.setRowCount(new_rows)
            
        for r, info in enumerate(self.app.visible_tasks_info):
            t = info['task']
            indent = "    " * info['indent']
            is_group = t.get('is_group', False)
            
            for c in range(total_cols):
                if self.app.table.item(r, c) is None:
                    self.app.table.setItem(r, c, QTableWidgetItem())
            
            mark_item = self.app.table.item(r, 0)
            mark_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            mark_item.setBackground(QColor(255, 255, 255))
            if is_group: mark_item.setBackground(QColor(242, 242, 242))

            toggle_item = self.app.table.item(r, 1)
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

            item_name = self.app.table.item(r, 2)
            item_name.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_name.setForeground(QColor(51, 51, 51))
            f = item_name.font(); f.setBold(False); item_name.setFont(f)
            item_name.setBackground(QColor(255, 255, 255))
            item_name.setText(indent + t.get('name', ''))
            
            item_mode = self.app.table.item(r, 3)
            item_mode.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            item_mode.setForeground(QColor(51, 51, 51))
            item_mode.setBackground(QColor(255, 255, 255))
            
            item_hc = self.app.table.item(r, 4)
            item_hc.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_hc.setForeground(QColor(51, 51, 51))
            item_hc.setBackground(QColor(255, 255, 255))
            
            item_prog = self.app.table.item(r, 5)
            item_prog.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_prog.setForeground(QColor(51, 51, 51))
            item_prog.setBackground(QColor(255, 255, 255))
            
            item_period = self.app.table.item(r, 6)
            item_period.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_period.setForeground(QColor(51, 51, 51))
            item_period.setBackground(QColor(255, 255, 255))
            
            item_color = self.app.table.item(r, 7)
            item_color.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            item_color.setForeground(QColor(51, 51, 51))
            item_color.setBackground(QColor(255, 255, 255))
            
            if is_group:
                f = item_name.font(); f.setBold(True); item_name.setFont(f)
                item_name.setBackground(QColor(242, 242, 242))
                
                item_mode.setText("")
                item_mode.setBackground(QColor(242, 242, 242))
                
                item_hc.setText(f"{t.get('headcount', 1.0):.1f}")
                item_hc.setTextAlignment(Qt.AlignCenter)
                item_hc.setBackground(QColor(242, 242, 242))
                
                item_prog.setText("")
                item_prog.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                item_prog.setBackground(QColor(242, 242, 242))
                item_period.setText("")
                item_period.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                item_period.setBackground(QColor(242, 242, 242))
                item_color.setText("")
                item_color.setBackground(QColor(200, 200, 200))
                item_color.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            else:
                is_auto = t.get('mode') == 'auto'
                item_mode.setText("生成" if is_auto else "作成")
                item_mode.setTextAlignment(Qt.AlignCenter)
                
                if is_auto:
                    item_hc.setText(f"{t.get('workload', 10.0):.1f}")
                    item_period.setText(t.get('auto_start_date', ''))
                else:
                    item_hc.setText(f"{t.get('headcount', 1.0):.1f}")
                    periods = t.get('periods', [])
                    p_strs = []
                    for p in periods:
                        if not p.get('start_date') or not p.get('end_date'): continue
                        s = p['start_date'].replace('-', '/')
                        e = p['end_date'].replace('-', '/')
                        p_strs.append(f"{s}-{e}")
                    item_period.setText(", ".join(p_strs))
                    
                item_hc.setTextAlignment(Qt.AlignCenter)
                item_prog.setText(str(t.get('progress', 0)))
                item_prog.setTextAlignment(Qt.AlignCenter)
                
                item_color.setText("")
                item_color.setBackground(QColor(t.get('color', '#0078d4')))
                item_color.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)

            for i, (h_start, h_end, _) in enumerate(headers):
                col_idx = 8 + i
                item_s = self.app.table.item(r, col_idx)
                item_s.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                item_s.setForeground(QColor(51, 51, 51))
                item_s.setTextAlignment(Qt.AlignCenter)
                
                day_map = self.app.get_task_workload_in_range(t, info['index'], h_start, h_end)
                item_s.setText(self.app.format_summary_workload(day_map))
                
                if is_group:
                    item_s.setBackground(QColor(242, 242, 242))
                else:
                    item_s.setBackground(QColor(255, 255, 255))
                
        if self.app.table.rowCount() > new_rows:
            self.app.table.setRowCount(new_rows)
            
        for i in range(base_col_count, total_cols):
            self.app.table.setColumnHidden(i, not self.app.summary_visible)
            if self.app.summary_visible:
                self.app.table.setColumnWidth(i, 90)

        self.app.update_selection_mark()
        self.app.table.blockSignals(False)
        if refresh_chart:
            self.draw_chart()