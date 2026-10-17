"""Dynamic filter condition editor built from imported columns."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QComboBox, QHBoxLayout, QLineEdit, QPushButton, QWidget

class FilterCondition(QWidget):
    removed = Signal(object)
    changed = Signal()

    def __init__(self, fields, numeric_fields, parent=None):
        super().__init__(parent)
        self.numeric_fields = numeric_fields
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 2, 0, 2)
        self.field = QComboBox()
        self.field.addItem('选择列…', None)
        self.field.addItems(fields)
        self.field.setCurrentIndex(0)
        self.field.setMinimumWidth(180)
        self.operator = QComboBox()
        self.operator.addItem('比较方式…', None)
        self.operator.setEnabled(False)
        self.operator.setMinimumWidth(118)
        self.value = QLineEdit()
        self.value.setPlaceholderText('筛选值')
        self.value.setEnabled(False)
        self.value.setMinimumWidth(180)
        self.remove_button = QPushButton('删除')
        layout.addWidget(self.field)
        layout.addWidget(self.operator)
        layout.addWidget(self.value, 1)
        layout.addWidget(self.remove_button)
        self.field.currentIndexChanged.connect(self._field_changed)
        self.operator.currentIndexChanged.connect(self._operator_changed)
        self.value.textChanged.connect(self.changed)
        self.remove_button.clicked.connect(lambda: self.removed.emit(self))

    def _field_changed(self):
        self.operator.blockSignals(True)
        self.operator.clear()
        self.operator.addItem('比较方式…', None)
        field = self.field.currentText()
        if self.field.currentIndex() > 0:
            options = (['<', '<=', '=', '!=', '>=', '>', '为空', '不为空']
                       if field in self.numeric_fields else
                       ['包含', '不包含', '等于', '不等于', '为空', '不为空'])
            self.operator.addItems(options)
        self.operator.setCurrentIndex(0)
        self.operator.blockSignals(False)
        self.operator.setEnabled(self.field.currentIndex() > 0)
        self.value.clear()
        self.value.setEnabled(False)
        self.changed.emit()

    def _operator_changed(self):
        op = self.operator.currentText()
        self.value.setEnabled(bool(op and op != '比较方式…' and op not in ('为空', '不为空')))
        if op in ('为空', '不为空'):
            self.value.clear()
        self.changed.emit()

    def condition(self):
        field = self.field.currentText() if self.field.currentIndex() > 0 else ''
        operator = self.operator.currentText() if self.operator.currentIndex() > 0 else ''
        value = self.value.text().strip()
        if not field and not operator and not value:
            return None
        if not field or not operator:
            raise ValueError('请先为每条筛选条件选择列和比较方式。')
        if operator not in ('为空', '不为空') and not value:
            raise ValueError(f'“{field}”条件还没有填写筛选值。')
        return {'field': field, 'operator': operator, 'value': value}
