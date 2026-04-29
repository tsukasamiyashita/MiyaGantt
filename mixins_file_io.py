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

class FileIOMixin:
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

