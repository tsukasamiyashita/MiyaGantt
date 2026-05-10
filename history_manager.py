# history_manager.py
import json

class HistoryManagerMixin:
    def init_history(self):
        self.undo_stack = []
        self.redo_stack = []
        self.current_state_json = self.get_state_json()
        self.update_history_buttons()

    def get_state_json(self):
        return json.dumps({
            "project_title": getattr(self, 'project_title', ""),
            "tasks": self.tasks,
            "custom_holidays": getattr(self, 'custom_holidays', {})
        }, ensure_ascii=False)

    def save_state_if_changed(self):
        new_state = self.get_state_json()
        if new_state != self.current_state_json:
            self.undo_stack.append(self.current_state_json)
            self.redo_stack.clear()
            self.current_state_json = new_state
            if len(self.undo_stack) > 100:
                self.undo_stack.pop(0)
            self.update_history_buttons()

    def undo(self):
        if not self.undo_stack: return
        self.table.clearSelection()
        self.redo_stack.append(self.current_state_json)
        self.current_state_json = self.undo_stack.pop()
        self.restore_state_json(self.current_state_json)
        self.update_history_buttons()

    def redo(self):
        if not self.redo_stack: return
        self.table.clearSelection()
        self.undo_stack.append(self.current_state_json)
        self.current_state_json = self.redo_stack.pop()
        self.restore_state_json(self.current_state_json)
        self.update_history_buttons()

    def update_history_buttons(self):
        if hasattr(self, 'btn_undo'):
            self.btn_undo.setEnabled(len(self.undo_stack) > 0)
        if hasattr(self, 'btn_redo'):
            self.btn_redo.setEnabled(len(self.redo_stack) > 0)

    def restore_state_json(self, state_json):
        state = json.loads(state_json)
        self.project_title = state.get("project_title", "")
        if hasattr(self, 'title_edit'):
            self.title_edit.blockSignals(True)
            self.title_edit.setText(self.project_title)
            self.title_edit.blockSignals(False)
            self.on_title_changed(self.project_title)
        self.tasks = state.get("tasks", [])
        self.custom_holidays = state.get("custom_holidays", {})
        self.recalculate_auto_tasks()
        self.update_ui()