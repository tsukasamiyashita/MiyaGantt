import sys
import os
import calendar
from datetime import datetime, timedelta
import jpholiday
import shiboken6
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *

class NoHighlightDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        # 選択状態とフォーカス状態のフラグを落としてから標準の描画を行う
        # これにより、クリック時も背景色（グループ行のグレーなど）がハイライト色やCSS設定で上書きされない
        opt = QStyleOptionViewItem(option)
        opt.state &= ~QStyle.State_Selected
        opt.state &= ~QStyle.State_HasFocus
        super().paint(painter, opt, index)

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
        self.setItemDelegate(NoHighlightDelegate(self))
        
        # ヘッダーの右クリックメニューを有効化
        self.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.horizontalHeader().customContextMenuRequested.connect(self.show_header_menu)

    def show_header_menu(self, pos):
        menu = QMenu(self)
        column_names = ["選択マーク", "開閉ボタン", "タスク名", "種別", "進捗(%)", "人数/工数", "期間指定/開始日", "色", "集計列"]
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

