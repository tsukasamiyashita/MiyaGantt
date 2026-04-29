# tsukasamiyashita/miyagantt/MiyaGantt-46a1664b6d1737cb32f1dd17429ce06cca8dc678/task_table.py
from PySide6.QtWidgets import QHeaderView, QTableWidget, QMenu
from PySide6.QtCore import Qt

class HideableHeader(QHeaderView):
    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self.setSectionsClickable(True)
        self.btn_size = 16
        
    def paintSection(self, painter, rect, logicalIndex):
        # 目玉アイコンなどのカスタム描画を削除し、標準のヘッダー表示に戻す
        super().paintSection(painter, rect, logicalIndex)

    def mouseReleaseEvent(self, e):
        # カスタムボタンを削除したため、標準のイベント処理のみ行う
        super().mouseReleaseEvent(e)

class TaskTable(QTableWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setHorizontalHeader(HideableHeader(Qt.Horizontal, self))
        self.setSelectionMode(QTableWidget.NoSelection)
        self.setSelectionBehavior(QTableWidget.SelectRows)
        self.setFocusPolicy(Qt.NoFocus)
        
        # ヘッダーの右クリックメニューを有効化
        self.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.horizontalHeader().customContextMenuRequested.connect(self.show_header_menu)

    def show_header_menu(self, pos):
        menu = QMenu(self)
        column_names = ["選択マーク", "開閉ボタン", "タスク名", "進捗(%)", "期間指定", "色", "集計"]
        for i, name in enumerate(column_names):
            action = menu.addAction(name)
            action.setCheckable(True)
            if i < 6:
                action.setChecked(not self.isColumnHidden(i))
            else:
                action.setChecked(self.window().summary_visible)
            
            if i == 2:
                action.setEnabled(False)
                
            action.toggled.connect(lambda checked, idx=i: self.window().toggle_column_visibility(idx, checked))
        
        menu.exec(self.horizontalHeader().mapToGlobal(pos))