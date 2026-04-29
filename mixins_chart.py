# tsukasamiyashita/miyagantt/MiyaGantt-90775b445eeca08d321c122853c84ad8762e2c95/mixins_chart.py
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

class ChartMixin:
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
            
            custom_status = self.custom_holidays.get(d_str)
            is_public = jpholiday.is_holiday(d)
            
            # 背景色の決定
            bg = None
            if custom_status == "営業日":
                pass # 営業日として指定された場合は色なし
            else:
                is_custom_holiday = (custom_status is not None and custom_status != "営業日")
                if d.weekday() == 5: # 土曜日
                    if is_custom_holiday or is_public:
                        bg = QColor(255, 240, 240) # 休日指定、または祝日と重なっている場合は赤
                    else:
                        bg = QColor(240, 248, 255) # 通常の土曜日（青）
                elif d.weekday() == 6 or is_public or is_custom_holiday: # 日曜日、祝日、カスタム祝日
                    bg = QColor(255, 240, 240) # 赤
            
            if bg:
                re = self.cs.addRect(x, 0, self.day_width, ch, QPen(Qt.NoPen), QBrush(bg))
                re.setZValue(-20)
                re.setAcceptedMouseButtons(Qt.NoButton)
                re.setFlag(QGraphicsItem.ItemIsSelectable, False)
            
            # グリッド
            gl = self.cs.addLine(x, 0, x, ch, QPen(QColor(220, 220, 220), 1))
            gl.setZValue(-15)
            gl.setAcceptedMouseButtons(Qt.NoButton)
            gl.setFlag(QGraphicsItem.ItemIsSelectable, False)
            
            if self.day_width >= 60:
                for h in [6, 12, 18]:
                    sl = self.cs.addLine(x + (self.day_width * h / 24.0), 0, x + (self.day_width * h / 24.0), ch, QPen(QColor(245, 245, 245), 0.5))
                    sl.setZValue(-15)
                    sl.setAcceptedMouseButtons(Qt.NoButton)
                    sl.setFlag(QGraphicsItem.ItemIsSelectable, False)
            
            # ヘッダー (日付・曜日)
            h_bg = QColor(255, 255, 225) if custom_status else QColor(248, 248, 248)
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
                if custom_status == "営業日":
                    w_c = QColor(60, 60, 60) # 営業日扱い
                elif custom_status is not None and custom_status != "営業日":
                    w_c = QColor(220, 0, 0) # 休日扱い
                elif d.weekday() == 5 and not is_public: # 土曜日
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
                
                if is_public or custom_status:
                    h_name = custom_status if custom_status else jpholiday.is_holiday_name(d)
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
            line = self.cs.addLine(0, y, tw_total, y, QPen(QColor(220, 220, 220), 1))
            line.setZValue(-15)
            line.setAcceptedMouseButtons(Qt.NoButton)
            line.setFlag(QGraphicsItem.ItemIsSelectable, False)
        
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

    def on_zoom_changed(self, *_):
        self.zoom_unit = self.zoom_unit_combo.currentIndex()
        self.zoom_count = self.zoom_count_spin.value()
        # ズーム単位を表示単位（集計単位）にも適用する
        self.display_unit = self.zoom_unit
        self.update_display_days()
        self.calculate_day_width()
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