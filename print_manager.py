# print_manager.py
from datetime import timedelta
import jpholiday
from PySide6.QtWidgets import QMessageBox, QStyleOptionGraphicsItem, QGraphicsTextItem
from PySide6.QtGui import QPainter, QPageLayout, QPen, QColor, QFont, QAction
from PySide6.QtPrintSupport import QPrinter, QPrintPreviewDialog
from PySide6.QtCore import Qt, QRectF, QPointF
from dialogs import PrintSettingsDialog
from gantt_items import GanttBarItem, GanttCommentItem

class PrintManagerMixin:
    def print_gantt(self):
        if not self.visible_tasks_info:
            QMessageBox.information(self, "情報", "印刷するタスクがありません。")
            return
            
        scroll_val = self.chart_view.horizontalScrollBar().value()
        days_scrolled = scroll_val / self.day_width if getattr(self, 'day_width', 0) > 0 else 0
        visible_start = self.min_date + timedelta(days=days_scrolled)
        
        viewport_width = self.chart_view.viewport().width()
        visible_days = viewport_width / self.day_width if getattr(self, 'day_width', 0) > 0 else 30
        visible_end = visible_start + timedelta(days=visible_days - 1)
        
        dlg = PrintSettingsDialog(self, self.visible_tasks_info, visible_start, visible_end)
        if not dlg.exec():
            return
            
        sd, ed, selected_row_indices, selected_col_indices = dlg.get_settings()
        if sd > ed:
            QMessageBox.warning(self, "エラー", "開始日が終了日より後になっています。")
            return
        if not selected_row_indices:
            QMessageBox.warning(self, "エラー", "印刷対象の行が選択されていません。")
            return
        if not selected_col_indices:
            QMessageBox.warning(self, "エラー", "印刷対象の列が選択されていません。")
            return

        printer = QPrinter(QPrinter.ScreenResolution)
        printer.setPageOrientation(QPageLayout.Landscape)
        
        preview = QPrintPreviewDialog(printer, self)
        preview.setWindowTitle("印刷プレビュー")
        preview.setWindowFlags(preview.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint)
        
        translations = {
            "print...": "印刷...",
            "print": "印刷",
            "fit width": "幅に合わせる",
            "fit to width": "幅に合わせる",
            "fit page": "ページに合わせる",
            "fit to page": "ページに合わせる",
            "zoom in": "拡大",
            "zoom out": "縮小",
            "portrait": "縦向き",
            "landscape": "横向き",
            "first page": "最初のページ",
            "previous page": "前のページ",
            "next page": "次のページ",
            "last page": "最後のページ",
            "show overview of all pages": "全ページ表示",
            "show single page": "単一ページ表示",
            "show facing pages": "見開き表示",
            "page setup...": "ページ設定...",
            "page setup": "ページ設定",
            "export to pdf...": "PDFへエクスポート...",
            "export to pdf": "PDFへエクスポート"
        }
        
        for action in preview.findChildren(QAction):
            text = action.text().replace("&", "").strip().lower()
            tooltip = action.toolTip().strip()
            tooltip_lower = tooltip.lower()
            
            tooltip_clean = tooltip_lower.split(" (")[0].strip()
            
            if tooltip_clean in translations:
                shortcut_part = tooltip[len(tooltip_clean):] 
                action.setText(translations[tooltip_clean])
                action.setToolTip(translations[tooltip_clean] + shortcut_part)
            elif text in translations:
                action.setText(translations[text])
                action.setToolTip(translations[text])

        preview.paintRequested.connect(lambda p: self.render_to_printer(p, sd, ed, selected_row_indices, selected_col_indices))
        preview.showMaximized()
        preview.exec()

    def render_to_printer(self, printer, sd, ed, selected_row_indices, selected_col_indices):
        original_row_hidden = []
        for r in range(len(self.visible_tasks_info)):
            original_row_hidden.append(self.table.isRowHidden(r))
            self.table.setRowHidden(r, r not in selected_row_indices)

        original_col_hidden = []
        for c in range(self.table.columnCount()):
            original_col_hidden.append(self.table.isColumnHidden(c))
            self.table.setColumnHidden(c, c not in selected_col_indices)

        days = (ed - sd).days + 1
        start_offset_days = (sd - self.min_date).days
        chart_start_x = start_offset_days * getattr(self, 'day_width', 40)
        chart_width = days * getattr(self, 'day_width', 40)
        
        table_width = 0
        for c in range(self.table.columnCount()):
            if not self.table.isColumnHidden(c):
                table_width += self.table.columnWidth(c)
                
        title_height = 50 if getattr(self, 'project_title', '') else 0
        total_height = title_height + self.header_height + len(selected_row_indices) * self.row_height
        
        rect = printer.pageRect(QPrinter.DevicePixel)
        painter = QPainter(printer)
        
        logical_width = table_width + chart_width
        logical_height = total_height
        
        scale_x = rect.width() / logical_width if logical_width > 0 else 1.0
        scale_y = rect.height() / logical_height if logical_height > 0 else 1.0
        scale = min(scale_x, scale_y)
        
        painter.scale(scale, scale)
        painter.fillRect(QRectF(0, 0, logical_width, logical_height), Qt.white)
        
        if getattr(self, 'project_title', ''):
            painter.save()
            painter.setFont(QFont("Segoe UI", 16, QFont.Bold))
            painter.setPen(Qt.black)
            painter.drawText(QRectF(10, 0, logical_width - 10, title_height), Qt.AlignLeft | Qt.AlignVCenter, self.project_title)
            painter.restore()
        
        painter.save()
        painter.translate(0, title_height)
        curr_x = 0
        painter.setFont(QFont("Segoe UI", 9, QFont.Bold))
        for c in range(self.table.columnCount()):
            if not self.table.isColumnHidden(c):
                cw = self.table.columnWidth(c)
                header_rect = QRectF(curr_x, 0, cw, self.header_height)
                painter.setPen(QPen(QColor(200, 200, 200), 1))
                painter.drawRect(header_rect)
                painter.setPen(Qt.black)
                label = self.table.horizontalHeaderItem(c).text() if self.table.horizontalHeaderItem(c) else ""
                painter.drawText(header_rect, Qt.AlignCenter | Qt.TextWordWrap, label)
                curr_x += cw
                
        painter.setFont(QFont("Segoe UI", 9))
        curr_y = self.header_height
        for r in range(len(self.visible_tasks_info)):
            if r in selected_row_indices:
                curr_x = 0
                for c in range(self.table.columnCount()):
                    if not self.table.isColumnHidden(c):
                        cw = self.table.columnWidth(c)
                        cell_rect = QRectF(curr_x, curr_y, cw, self.row_height)
                        painter.setPen(QPen(QColor(200, 200, 200), 1))
                        painter.drawRect(cell_rect)
                        painter.setPen(Qt.black)
                        item = self.table.item(r, c)
                        text = item.text() if item else ""
                        if c == 2 and item:
                            info = self.visible_tasks_info[r]
                            indent = "  " * info.get('indent', 0)
                            text = indent + text
                        if item and item.background().color() != Qt.transparent:
                            painter.fillRect(cell_rect.adjusted(1,1,-1,-1), item.background())
                        painter.drawText(cell_rect.adjusted(4,0,-4,0), Qt.AlignLeft | Qt.AlignVCenter, text)
                        curr_x += cw
                curr_y += self.row_height
        painter.restore()
        
        painter.save()
        painter.translate(table_width, title_height)
        painter.setClipRect(QRectF(0, 0, chart_width, total_height - title_height))
        
        header_source_rect = QRectF(chart_start_x, 0, chart_width, self.header_height)
        header_target_rect = QRectF(0, 0, chart_width, self.header_height)
        self.hs.render(painter, header_target_rect, header_source_rect)
        
        curr_y = self.header_height
        dw = getattr(self, 'day_width', 40)
        for r in range(len(self.visible_tasks_info)):
            if r in selected_row_indices:
                line_rect = QRectF(0, curr_y, chart_width, self.row_height)
                painter.setPen(QColor(230, 230, 230))
                painter.drawLine(line_rect.bottomLeft(), line_rect.bottomRight())
                
                for i in range(days):
                    d = sd + timedelta(days=i)
                    d_str = d.strftime("%Y-%m-%d")
                    if d.weekday() >= 5 or jpholiday.is_holiday(d.date()) or d_str in getattr(self, 'custom_holidays', {}):
                        holiday_rect = QRectF(i * dw, curr_y, dw, self.row_height)
                        painter.fillRect(holiday_rect, QColor(245, 245, 245))
                        
                painter.setPen(QPen(QColor(230, 230, 230), 1))
                for i in range(days + 1):
                    px = i * dw
                    painter.drawLine(QPointF(px, curr_y), QPointF(px, curr_y + self.row_height))
                
                scene_row_y = r * self.row_height
                for item in self.cs.items():
                    if isinstance(item, (GanttBarItem, GanttCommentItem)) and getattr(item, 'row', -1) == r:
                        item_x = item.scenePos().x() - chart_start_x
                        item_y = curr_y + (item.scenePos().y() - scene_row_y)
                        if item_x + item.boundingRect().width() < 0 or item_x > chart_width:
                            continue
                        
                        painter.save()
                        painter.translate(item_x, item_y)
                        opt = QStyleOptionGraphicsItem()
                        item.paint(painter, opt)
                        
                        for child in item.childItems():
                            if isinstance(child, QGraphicsTextItem):
                                painter.save()
                                painter.translate(child.pos())
                                painter.setFont(child.font())
                                painter.setPen(child.defaultTextColor())
                                painter.drawText(child.boundingRect(), Qt.AlignLeft | Qt.AlignTop, child.toPlainText())
                                painter.restore()
                        painter.restore()
                curr_y += self.row_height
                
        painter.restore()
        painter.end()

        for r, hidden in enumerate(original_row_hidden):
            self.table.setRowHidden(r, hidden)
        for c, hidden in enumerate(original_col_hidden):
            self.table.setColumnHidden(c, hidden)