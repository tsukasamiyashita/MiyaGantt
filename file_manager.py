# tsukasamiyashita/miyagantt/MiyaGantt-46a1664b6d1737cb32f1dd17429ce06cca8dc678/file_manager.py
import json
import os
from datetime import datetime, timedelta
from PySide6.QtWidgets import QFileDialog, QMessageBox
from PySide6.QtCore import QTimer

class FileManagerMixin:
    def get_config_path(self):
        return os.path.join(os.path.expanduser("~"), ".miyagantt_config.json")

    def save_app_config(self):
        column_widths = {}
        if hasattr(self, 'table'):
            for i in range(self.table.columnCount()):
                column_widths[str(i)] = self.table.columnWidth(i)
                
        column_visibility = {}
        if hasattr(self, 'col_actions'):
            for idx, btn in self.col_actions.items():
                column_visibility[str(idx)] = btn.isChecked()

        config = {
            "zoom_unit": getattr(self, 'zoom_unit', 1),
            "zoom_count": getattr(self, 'zoom_count', 1),
            "display_unit": getattr(self, 'display_unit', 1),
            "display_count": getattr(self, 'display_count', 6),
            "summary_visible": getattr(self, 'summary_visible', True),
            "column_visibility": column_visibility,
            "column_widths": column_widths,
            "splitter_sizes": self.splitter.sizes() if hasattr(self, 'splitter') else [],
            "custom_holidays": getattr(self, 'custom_holidays', {}),
            "last_path": getattr(self, 'last_path', ""),
            "auto_disp_mode": getattr(self, 'auto_disp_mode', 0)
        }
        
        try:
            with open(self.get_config_path(), 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "設定保存", "現在の表示設定を保存しました。")
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"設定の保存に失敗しました。\n{e}")

    def load_app_config(self):
        config_path = self.get_config_path()
        if not os.path.exists(config_path): return
        
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                
            self.zoom_unit = config.get("zoom_unit", getattr(self, 'zoom_unit', 1))
            self.zoom_count = config.get("zoom_count", getattr(self, 'zoom_count', 1))
            self.display_unit = config.get("display_unit", getattr(self, 'display_unit', 1))
            self.display_count = config.get("display_count", getattr(self, 'display_count', 6))
            self.summary_visible = config.get("summary_visible", getattr(self, 'summary_visible', True))
            self.custom_holidays = config.get("custom_holidays", getattr(self, 'custom_holidays', {}))
            self.last_path = config.get("last_path", "")
            self.auto_disp_mode = config.get("auto_disp_mode", getattr(self, 'auto_disp_mode', 0))
            
            if hasattr(self, 'zoom_unit_combo'):
                self.zoom_unit_combo.blockSignals(True)
                self.zoom_unit_combo.setCurrentIndex(self.zoom_unit)
                self.zoom_unit_combo.blockSignals(False)
            if hasattr(self, 'zoom_count_spin'):
                self.zoom_count_spin.blockSignals(True)
                self.zoom_count_spin.setValue(self.zoom_count)
                self.zoom_count_spin.blockSignals(False)
            if hasattr(self, 'auto_disp_combo'):
                self.auto_disp_combo.blockSignals(True)
                self.auto_disp_combo.setCurrentIndex(self.auto_disp_mode)
                self.auto_disp_combo.blockSignals(False)
                
            column_widths = config.get("column_widths", {})
            if hasattr(self, 'table'):
                for i_str, w in column_widths.items():
                    try:
                        self.table.setColumnWidth(int(i_str), w)
                    except ValueError:
                        pass
                        
            col_vis = config.get("column_visibility", {})
            if hasattr(self, 'col_actions'):
                for idx_str, is_vis in col_vis.items():
                    try:
                        idx = int(idx_str)
                        if idx in self.col_actions:
                            self.col_actions[idx].blockSignals(True)
                            self.col_actions[idx].setChecked(is_vis)
                            self.col_actions[idx].blockSignals(False)
                            self.toggle_column_visibility(idx, is_vis)
                    except ValueError:
                        pass
                        
            splitter_sizes = config.get("splitter_sizes", [])
            if splitter_sizes and hasattr(self, 'splitter'):
                self.splitter.setSizes(splitter_sizes)
                
            self.update_display_range()
            
        except Exception as e:
            print(f"設定の読み込みに失敗しました: {e}")

    def new_project(self):
        reply = QMessageBox.question(self, '確認', '現在のデータを破棄して新規プロジェクトを作成しますか？',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.project_title = ""
            if hasattr(self, 'title_edit'):
                self.title_edit.blockSignals(True)
                self.title_edit.setText(self.project_title)
                self.title_edit.blockSignals(False)
                self.on_title_changed(self.project_title)
            
            self.tasks = []
            self.last_path = ""
            self.custom_holidays = {}
            
            self.min_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            self.update_display_days()
            self.calculate_day_width()
            
            self.recalculate_auto_tasks()
            self.update_ui()
            self.init_history()
            
            # 新規作成直後の状態を保存済みとして記録
            if hasattr(self, 'get_current_data_snapshot'):
                self.saved_snapshot = self.get_current_data_snapshot()

    def load_data(self):
        start_dir = os.path.dirname(self.last_path) if self.last_path else ""
        file_path, _ = QFileDialog.getOpenFileName(self, "ファイルを開く", start_dir, "JSON Files (*.json);;All Files (*)")
        if not file_path: return
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            self.project_title = data.get('project_title', '')
            self.title_edit.blockSignals(True)
            self.title_edit.setText(self.project_title)
            self.on_title_changed(self.project_title)
            self.title_edit.blockSignals(False)
            
            min_date_str = data.get('min_date', datetime.now().strftime("%Y-%m-%d"))
            self.min_date = datetime.strptime(min_date_str, "%Y-%m-%d")
            
            max_date_str = data.get('max_date', (self.min_date + timedelta(days=180)).strftime("%Y-%m-%d"))
            self.max_date = datetime.strptime(max_date_str, "%Y-%m-%d")
            
            self.tasks = data.get('tasks', [])
            self.last_path = file_path
            
            self.update_display_days()
            self.calculate_day_width()
            self.recalculate_auto_tasks()
            self.update_ui()
            self.init_history()
            
            # 読み込み成功時に保存状態としてマーク
            if hasattr(self, 'get_current_data_snapshot'):
                self.saved_snapshot = self.get_current_data_snapshot()
            
            scroll_date_str = data.get('scroll_date')
            if scroll_date_str:
                try:
                    target_date = datetime.strptime(scroll_date_str, "%Y-%m-%d")
                    QTimer.singleShot(100, lambda: self._scroll_to_specific_date(target_date))
                except ValueError:
                    pass
            
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"ファイルの読み込みに失敗しました。\n{e}")

    def save_data(self):
        if not self.last_path:
            return self.save_data_as()
            
        return self._perform_save(self.last_path)

    def save_data_as(self):
        start_dir = os.path.dirname(self.last_path) if self.last_path else ""
        file_path, _ = QFileDialog.getSaveFileName(self, "名前を付けて保存", start_dir, "JSON Files (*.json);;All Files (*)")
        if not file_path:
            return False
        
        return self._perform_save(file_path)

    def _perform_save(self, file_path):
        scroll_val = self.chart_view.horizontalScrollBar().value()
        if getattr(self, 'day_width', 0) > 0:
            days_scrolled = round(scroll_val / self.day_width)
        else:
            days_scrolled = 0
            
        visible_start = getattr(self, 'min_date', datetime.now()) + timedelta(days=days_scrolled)
        
        data = {
            'project_title': getattr(self, 'project_title', ''),
            'min_date': getattr(self, 'min_date', datetime.now()).strftime("%Y-%m-%d"),
            'max_date': getattr(self, 'max_date', datetime.now() + timedelta(days=180)).strftime("%Y-%m-%d"),
            'scroll_date': visible_start.strftime("%Y-%m-%d"),
            'tasks': getattr(self, 'tasks', [])
        }
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.last_path = file_path
            
            # 保存成功時に現在のスナップショットを記録
            if hasattr(self, 'get_current_data_snapshot'):
                self.saved_snapshot = self.get_current_data_snapshot()
                
            QMessageBox.information(self, "保存", "ファイルを保存しました。")
            return True
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"ファイルの保存に失敗しました。\n{e}")
            return False