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

class UtilMixin:
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

