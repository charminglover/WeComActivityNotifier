"""Generic routing editor with visible destinations and asynchronous persistence."""
from __future__ import annotations

from PySide6.QtCore import QThreadPool
from PySide6.QtWidgets import (QAbstractItemView, QComboBox, QDialog, QFileDialog,
    QFormLayout, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton,
    QTableWidget, QTableWidgetItem, QVBoxLayout)

from services.task_worker import Worker


class MappingEntryDialog(QDialog):
    def __init__(self, kind: str, current=None, parent=None, field=''):
        super().__init__(parent)
        self.current = current or {}
        self.setWindowTitle('编辑映射' if current else '添加映射')
        self.resize(660, 160)
        self.match_value = QLineEdit(self.current.get('value', ''))
        self.match_value.setPlaceholderText(f'填写“{field}”列中的值' if field else '填写 Excel 列中需要匹配的值')
        self.target = QLineEdit(self.current.get('target', ''))
        is_webhook = kind == '群机器人Webhook'
        self.target.setPlaceholderText('https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=…'
                                       if is_webhook else '多个 UserID 可用 | 分隔')
        form = QFormLayout(self)
        form.addRow(f'匹配值（{field}）' if field else '匹配值', self.match_value)
        form.addRow('Webhook URL' if is_webhook else 'UserID', self.target)
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        self.ok_button = QPushButton('确定')
        cancel = QPushButton('取消')
        self.ok_button.clicked.connect(self._accept_if_valid)
        cancel.clicked.connect(self.reject)
        buttons.addWidget(self.ok_button)
        buttons.addWidget(cancel)
        form.addRow(buttons)

    def _accept_if_valid(self):
        try:
            self.values()
        except ValueError as exc:
            QMessageBox.warning(self, '信息不完整', str(exc))
            return
        self.accept()

    def values(self):
        value = self.match_value.text().strip()
        target = self.target.text().strip()
        if not value or not target:
            raise ValueError('匹配值和接收目标均不能为空。')
        return {'id': self.current.get('id'), 'value': value, 'target': target}


class MappingDialog(QDialog):
    def __init__(self, store, kind: str, parent=None, fields=None, rows=None, field=None):
        super().__init__(parent)
        self.store = store
        self.kind = kind
        self._worker = None
        self.setWindowTitle('接收目标映射')
        self.resize(920, 490)
        self.rows = [dict(row) for row in (store.all() if rows is None else rows)]
        saved_field = store.settings()['field'] if field is None else field
        layout = QVBoxLayout(self)
        buttons = QHBoxLayout()
        self.action_buttons = []
        for label, callback in (('添加', self.add_row), ('编辑', self.edit_row),
                                ('删除', self.delete_row), ('从 CSV 导入…', self.import_csv)):
            button = QPushButton(label)
            button.clicked.connect(callback)
            buttons.addWidget(button)
            self.action_buttons.append(button)
        buttons.addStretch(1)
        self.save_button = QPushButton('保存')
        self.save_button.clicked.connect(self.save)
        self.cancel_button = QPushButton('取消')
        self.cancel_button.clicked.connect(self.reject)
        buttons.addWidget(self.save_button)
        buttons.addWidget(self.cancel_button)
        layout.addLayout(buttons)

        self.field_combo = QComboBox()
        self.field_combo.setEditable(True)
        choices = list(dict.fromkeys(fields or []))
        if saved_field and saved_field not in choices:
            choices.insert(0, saved_field)
        self.field_combo.addItems(choices)
        self.field_combo.setCurrentIndex(-1)
        self.field_combo.lineEdit().setPlaceholderText('选择导入列，或输入完整表头名称')
        if saved_field:
            self.field_combo.setCurrentText(saved_field)
        form = QFormLayout()
        form.addRow('匹配字段', self.field_combo)
        layout.addLayout(form)
        self.status = QLabel('按所选列的值匹配下方目标；更换字段后，请同步检查匹配值。所有映射删除并保存后停用映射。')
        self.status.setWordWrap(True)
        layout.addWidget(self.status)
        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(['匹配值', 'Webhook URL' if kind == '群机器人Webhook' else 'UserID'])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setColumnWidth(0, 210)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.doubleClicked.connect(lambda index: self.edit_row())
        layout.addWidget(self.table, 1)
        self._refresh()

    def _refresh(self):
        self.table.setRowCount(len(self.rows))
        for index, mapping in enumerate(self.rows):
            self.table.setItem(index, 0, QTableWidgetItem(mapping['value']))
            item = QTableWidgetItem(mapping.get('target', ''))
            item.setToolTip(mapping.get('target', ''))
            self.table.setItem(index, 1, item)

    def _target_prompt(self, current=None):
        dialog = MappingEntryDialog(self.kind, current, self, self.field_combo.currentText().strip())
        if dialog.exec() == QDialog.DialogCode.Accepted:
            return dialog.values()
        return None

    def add_row(self):
        result = self._target_prompt()
        if result:
            self.rows.append(result)
            self._refresh()
            self.table.selectRow(len(self.rows) - 1)

    def edit_row(self):
        index = self.table.currentRow()
        if index < 0:
            self.status.setText('请先选择要编辑的映射。')
            return
        result = self._target_prompt(self.rows[index])
        if result:
            self.rows[index] = result
            self._refresh()
            self.table.selectRow(index)

    def delete_row(self):
        index = self.table.currentRow()
        if index >= 0:
            del self.rows[index]
            self._refresh()
        else:
            self.status.setText('请先选择要删除的映射。')

    def import_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, '导入映射 CSV', '', 'CSV (*.csv)')
        if not path:
            return
        try:
            imported = self.store.import_csv(path, self.kind)
            by_value = {row['value']: row for row in self.rows}
            by_value.update({row['value']: row for row in imported})
            self.rows = list(by_value.values())
            self._refresh()
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self, '导入失败', str(exc))

    def _set_busy(self, busy):
        for widget in [*self.action_buttons, self.save_button, self.cancel_button,
                       self.field_combo, self.table]:
            widget.setEnabled(not busy)

    def save(self):
        if self._worker is not None:
            return
        field = self.field_combo.currentText().strip()
        rows = [dict(row) for row in self.rows]
        if rows and not field:
            QMessageBox.warning(self, '缺少匹配字段', '请先选择或输入 Excel 匹配列的完整表头名称。')
            return
        self._set_busy(True)
        self.status.setText('正在保存映射…')
        worker = Worker(0, lambda progress, cancel: self.store.save(rows, field=field))
        self._worker = worker
        worker.signals.succeeded.connect(self._saved)
        worker.signals.failed.connect(self._save_failed)
        QThreadPool.globalInstance().start(worker)

    def _saved(self, task_id, result):
        self._worker = None
        self.accept()

    def _save_failed(self, task_id, error):
        self._worker = None
        self._set_busy(False)
        self.status.setText('保存失败，原有映射仍保留；请检查匹配值、目标地址和系统凭据库。')
        QMessageBox.warning(self, '保存失败', error)

    def reject(self):
        if self._worker is None:
            super().reject()

    def closeEvent(self, event):
        if self._worker is not None:
            event.ignore()
        else:
            super().closeEvent(event)
