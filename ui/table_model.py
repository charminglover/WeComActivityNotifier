"""Virtualized Qt table model for source records."""
from __future__ import annotations
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, Signal, QEvent
from PySide6.QtGui import QColor, QKeyEvent, QMouseEvent
from PySide6.QtWidgets import QStyledItemDelegate
from engine import text_of

class RecordTableModel(QAbstractTableModel):
    """Virtual table model: Qt paints only visible cells instead of creating row widgets."""
    checkedChanged = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.fields: list[str] = []
        self.rows: list[dict] = []
        self.checked: set[str] = set()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.fields) + 1

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < len(self.rows)):
            return None
        row = self.rows[index.row()]
        if index.column() == 0:
            if role == Qt.ItemDataRole.CheckStateRole:
                return (Qt.CheckState.Checked if row.get('_id') in self.checked
                        else Qt.CheckState.Unchecked)
            if role == Qt.ItemDataRole.TextAlignmentRole:
                return int(Qt.AlignmentFlag.AlignCenter)
            return None
        field = self.fields[index.column() - 1]
        if role == Qt.ItemDataRole.DisplayRole:
            return text_of(row.get(field), field)
        if role == Qt.ItemDataRole.ToolTipRole:
            return text_of(row.get(field), field)
        if role == Qt.ItemDataRole.BackgroundRole and row.get('_id') in self.checked:
            from PySide6.QtGui import QColor
            return QColor('#EAF2FC')
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return '勾选' if section == 0 else self.fields[section - 1]
        return str(section + 1)

    def flags(self, index):
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        base = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        return base | Qt.ItemFlag.ItemIsUserCheckable if index.column() == 0 else base

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if (not index.isValid() or index.column() != 0 or
                role != Qt.ItemDataRole.CheckStateRole):
            return False
        key = self.rows[index.row()].get('_id')
        checked_value = value == Qt.CheckState.Checked or value == Qt.CheckState.Checked.value
        if checked_value:
            self.checked.add(key)
        else:
            self.checked.discard(key)
        self.dataChanged.emit(index, index, [Qt.ItemDataRole.CheckStateRole,
                                              Qt.ItemDataRole.BackgroundRole])
        self.checkedChanged.emit()
        return True

    def set_content(self, fields, rows, checked=None):
        self.beginResetModel()
        self.fields = list(fields)
        self.rows = rows
        if checked is not None:
            self.checked = checked
        self.endResetModel()
        self.checkedChanged.emit()

    def refresh_checks(self, first=0, last=None):
        if self.rows:
            last = len(self.rows) - 1 if last is None else min(last, len(self.rows) - 1)
            first = max(0, first)
            if first <= last:
                self.dataChanged.emit(self.index(first, 0), self.index(last, 0),
                                      [Qt.ItemDataRole.CheckStateRole,
                                       Qt.ItemDataRole.BackgroundRole])
        self.checkedChanged.emit()


class CheckboxDelegate(QStyledItemDelegate):
    """Toggle the full checkbox cell on one click, with normal keyboard support."""
    def editorEvent(self, event, model, option, index):
        if not index.isValid() or index.column() != 0:
            return super().editorEvent(event, model, option, index)
        if isinstance(event, QMouseEvent) and event.type() == QEvent.Type.MouseButtonRelease:
            if event.button() == Qt.MouseButton.LeftButton:
                current = model.data(index, Qt.ItemDataRole.CheckStateRole)
                next_value = (Qt.CheckState.Unchecked if current == Qt.CheckState.Checked
                              else Qt.CheckState.Checked)
                return model.setData(index, next_value, Qt.ItemDataRole.CheckStateRole)
        if isinstance(event, QKeyEvent) and event.type() == QEvent.Type.KeyPress:
            if event.key() in (Qt.Key.Key_Space, Qt.Key.Key_Select):
                current = model.data(index, Qt.ItemDataRole.CheckStateRole)
                next_value = (Qt.CheckState.Unchecked if current == Qt.CheckState.Checked
                              else Qt.CheckState.Checked)
                return model.setData(index, next_value, Qt.ItemDataRole.CheckStateRole)
        return super().editorEvent(event, model, option, index)
