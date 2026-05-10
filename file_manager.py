# file_manager.py
import os
import json
from datetime import datetime, timedelta
from PySide6.QtWidgets import QFileDialog, QMessageBox
from PySide6.QtCore import QTimer

class FileManagerMixin:
    def save_data(self):
        if hasattr(self, 'last_path') and self.last_path:
            self._perform_save(self.last_path)
        else:
            self.save_data_as()

    def save_data_as(self):
        initial_dir = os.path.dirname(self.last_path) if hasattr(self, 'last_path') and self.last_path and os.path.exists(os.path.dirname(self.last_path)) else ""
        p = QFileDialog.getSaveFileName(self, "名前を付けて保存", initial_dir, "JSON (*.json)")[0]
        if p:
            self.last_path = p
            self._perform_save(p)

    def _perform_save(self, path):
        try:
            scroll_val = self.chart_view.horizontalScrollBar().value()
            days_scrolled = round(scroll_val / self.day_width) if getattr(self, 'day_width', 0) > 0 else 0
            visible_start = self.min_date + timedelta(days=days_scrolled)

            data_to_save = {
                "project_title": getattr(self, 'project_title', ""),
                "settings": {
                    "min_date": self.min_date.strftime("%Y-%m-%d"),
                    "max_date": self.max_date.strftime("%Y-%m-%d") if getattr(self, 'max_date', None) else None,
                    "display_unit": self.display_unit,
                    "display_count": self.display_count,
                    "zoom_unit": self.zoom_unit,
                    "zoom_count": self.zoom_count,
                    "custom_holidays": getattr(self, 'custom_holidays', {}),
                    "last_visible_start": visible_start.strftime("%Y-%m-%d")
                },
                "tasks": self.tasks
            }
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, ensure_ascii=False, indent=4)
            self.statusBar().showMessage(f"保存しました: {os.path.basename(path)}", 5000)
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"保存失敗: {e}")

    def load_data(self):
        initial_dir = os.path.dirname(self.last_path) if hasattr(self, 'last_path') and self.last_path and os.path.exists(os.path.dirname(self.last_path)) else ""
        p = QFileDialog.getOpenFileName(self, "開く", initial_dir, "JSON (*.json)")[0]
        if p:
            self.last_path = p
            try:
                with open(p, 'r', encoding='utf-8') as f:
                    loaded_data = json.load(f)
                
                last_visible_start_str = None

                if isinstance(loaded_data, dict) and "tasks" in loaded_data:
                    self.project_title = loaded_data.get("project_title", "")
                    self.title_edit.blockSignals(True)
                    self.title_edit.setText(self.project_title)
                    self.title_edit.blockSignals(False)
                    self.on_title_changed(self.project_title)

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
                        self.max_date = None
                    
                    self.display_unit = settings.get("display_unit")
                    self.display_count = settings.get("display_count")
                    
                    if self.display_unit is None or self.display_count is None:
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
                    
                    self.zoom_unit_combo.blockSignals(True)
                    self.zoom_count_spin.blockSignals(True)
                    self.zoom_unit_combo.setCurrentIndex(self.zoom_unit)
                    self.zoom_count_spin.setValue(self.zoom_count)
                    self.zoom_unit_combo.blockSignals(False)
                    self.zoom_count_spin.blockSignals(False)
                    
                    self.calculate_day_width()
                    self.update_display_days()
                    
                    last_visible_start_str = settings.get("last_visible_start")
                else:
                    self.tasks = loaded_data
                    self.project_title = ""
                    self.title_edit.blockSignals(True)
                    self.title_edit.setText("")
                    self.title_edit.blockSignals(False)
                    self.on_title_changed("")

                    self.min_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
                    self.display_unit = 1
                    self.display_count = 6
                    self.zoom_unit = 1
                    self.zoom_count = 1
                    self.calculate_day_width()
                    self.update_display_days()
                    last_visible_start_str = None
                    
                self.recalculate_auto_tasks()
                self.update_ui()
                self.init_history()
                
                if last_visible_start_str:
                    try:
                        target_date = datetime.strptime(last_visible_start_str, "%Y-%m-%d")
                        QTimer.singleShot(50, lambda: self._scroll_to_specific_date(target_date))
                    except ValueError:
                        pass
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
        column_widths = {}
        for i in range(8):
            column_visibility[str(i)] = not self.table.isColumnHidden(i)
            column_widths[str(i)] = self.table.columnWidth(i)
        
        config = {
            "zoom_unit": getattr(self, 'zoom_unit', 1),
            "zoom_count": getattr(self, 'zoom_count', 1),
            "display_unit": getattr(self, 'display_unit', 1),
            "display_count": getattr(self, 'display_count', 6),
            "summary_visible": getattr(self, 'summary_visible', True),
            "column_visibility": column_visibility,
            "column_widths": column_widths,
            "splitter_sizes": self.splitter.sizes(),
            "custom_holidays": getattr(self, 'custom_holidays', {}),
            "last_path": getattr(self, 'last_path', "")
        }
        
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
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
            
            self.zoom_unit = config.get("zoom_unit", getattr(self, 'zoom_unit', 1))
            self.zoom_count = config.get("zoom_count", getattr(self, 'zoom_count', 1))
            self.display_unit = config.get("display_unit", getattr(self, 'display_unit', 1))
            self.display_count = config.get("display_count", getattr(self, 'display_count', 6))
            self.summary_visible = config.get("summary_visible", getattr(self, 'summary_visible', True))
            self.custom_holidays = config.get("custom_holidays", getattr(self, 'custom_holidays', {}))
            self.last_path = config.get("last_path", "")
            
            column_widths = config.get("column_widths", {})
            for idx_str, width in column_widths.items():
                if int(idx_str) < 8:
                    self.table.setColumnWidth(int(idx_str), width)
            
            splitter_sizes = config.get("splitter_sizes")
            if splitter_sizes:
                self.splitter.setSizes(splitter_sizes)
            
            if hasattr(self, 'zoom_unit_combo'):
                self.zoom_unit_combo.blockSignals(True)
                self.zoom_unit_combo.setCurrentIndex(self.zoom_unit)
                self.zoom_unit_combo.blockSignals(False)
            if hasattr(self, 'zoom_count_spin'):
                self.zoom_count_spin.blockSignals(True)
                self.zoom_count_spin.setValue(self.zoom_count)
                self.zoom_count_spin.blockSignals(False)
            
            col_vis = config.get("column_visibility", {})
            for idx_str, visible in col_vis.items():
                if int(idx_str) < 8:
                    self.toggle_column_visibility(int(idx_str), visible)
        except Exception as e:
            print(f"Config load error: {e}")