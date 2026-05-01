# tsukasamiyashita/miyagantt/MiyaGantt-46a1664b6d1737cb32f1dd17429ce06cca8dc678/task_table.py
from PySide6.QtWidgets import QHeaderView, QTableWidget, QMenu, QStyledItemDelegate, QComboBox
from PySide6.QtCore import Qt

class HeadcountDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QComboBox(parent)
        editor.addItem("制限なし")
        for i in range(1, 21):
            editor.addItem(str(i))
        return editor

    def setEditorData(self, editor, index):
        text = index.model().data(index, Qt.EditRole)
        if not text or text == "制限なし":
            editor.setCurrentIndex(0)
        else:
            val_str = str(text).replace('.0', '')
            idx = editor.findText(val_str)
            if idx >= 0:
                editor.setCurrentIndex(idx)
            else:
                editor.setCurrentIndex(0)

    def setModelData(self, editor, model, index):
        val = editor.currentText()
        if val == "制限なし":
            model.setData(index, "", Qt.EditRole)
        else:
            model.setData(index, val, Qt.EditRole)

class HideableHeader(QHeaderView):
    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self.setSectionsClickable(True)
        self.btn_size = 16
        
    def paintSection(self, painter, rect, logicalIndex):
        super().paintSection(painter, rect, logicalIndex)

    def mouseReleaseEvent(self, e):
        super().mouseReleaseEvent(e)

class TaskTable(QTableWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setHorizontalHeader(HideableHeader(Qt.Horizontal, self))
        self.setSelectionMode(QTableWidget.NoSelection)
        self.setSelectionBehavior(QTableWidget.SelectRows)
        self.setFocusPolicy(Qt.NoFocus)
        
        self.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.horizontalHeader().customContextMenuRequested.connect(self.show_header_menu)

    def show_header_menu(self, pos):
        menu = QMenu(self)
        column_names = ["選択マーク", "開閉ボタン", "タスク名", "モード", "人数", "進捗(%)", "期間/開始日", "色", "集計"]
        for i, name in enumerate(column_names):
            action = menu.addAction(name)
            action.setCheckable(True)
            if i < 8:
                action.setChecked(not self.isColumnHidden(i))
            else:
                action.setChecked(self.window().summary_visible)
            
            if i == 2:
                action.setEnabled(False)
                
            action.toggled.connect(lambda checked, idx=i: self.window().toggle_column_visibility(idx, checked))
        
        menu.exec(self.horizontalHeader().mapToGlobal(pos))