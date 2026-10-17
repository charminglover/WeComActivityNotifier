"""In-app CRUD for shop routing targets; target values are never displayed in clear text."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QAbstractItemView, QDialog, QFileDialog, QFormLayout,
    QHBoxLayout, QLineEdit, QMessageBox, QPushButton, QTableWidget, QTableWidgetItem,
    QVBoxLayout)


class MappingEntryDialog(QDialog):
    def __init__(self, kind: str, current=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle('编辑映射' if current else '添加映射')
        self.resize(480, 150)
        self.code = QLineEdit((current or {}).get('code', ''))
        self.target = QLineEdit()
        is_webhook = kind == '群机器人Webhook'
        if is_webhook:
            self.target.setEchoMode(QLineEdit.EchoMode.Password)
            self.target.setPlaceholderText('输入 Webhook；编辑时留空表示保留原值')
            target_label = 'Webhook URL'
        else:
            self.target.setPlaceholderText('多个 UserID 可用 | 分隔；留空表示保留原值')
            target_label = 'UserID'
        form = QFormLayout(self)
        form.addRow('服务店编码', self.code)
        form.addRow(target_label, self.target)
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        ok = QPushButton('确定')
        cancel = QPushButton('取消')
        ok.clicked.connect(self._accept_if_valid)
        cancel.clicked.connect(self.reject)
        buttons.addWidget(ok)
        buttons.addWidget(cancel)
        form.addRow(buttons)
        self.current = current or {}

    def _accept_if_valid(self):
        try:
            self.values()
        except ValueError as exc:
            QMessageBox.warning(self, '信息不完整', str(exc))
            return
        self.accept()

    def values(self):
        code = self.code.text().strip()
        target = self.target.text().strip() or self.current.get('target', '')
        if not code or not target:
            raise ValueError('服务店编码和接收目标均不能为空。')
        return {'id': self.current.get('id'), 'code': code, 'target': target}


class MappingDialog(QDialog):
    def __init__(self, store, kind: str, parent=None):
        super().__init__(parent)
        self.store = store
        self.kind = kind
        self.setWindowTitle('服务店接收映射')
        self.resize(680, 420)
        self.rows = store.all()
        layout = QVBoxLayout(self)
        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(['服务店编码', '接收目标（凭据库）'])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table, 1)
        buttons = QHBoxLayout()
        for label, callback in (('添加', self.add_row), ('编辑', self.edit_row),
                                ('删除', self.delete_row), ('从 CSV 导入…', self.import_csv)):
            button = QPushButton(label)
            button.clicked.connect(callback)
            buttons.addWidget(button)
        buttons.addStretch(1)
        save = QPushButton('保存')
        save.clicked.connect(self.save)
        cancel = QPushButton('取消')
        cancel.clicked.connect(self.reject)
        buttons.addWidget(save)
        buttons.addWidget(cancel)
        layout.addLayout(buttons)
        self._refresh()

    def _refresh(self):
        self.table.setRowCount(len(self.rows))
        for index, mapping in enumerate(self.rows):
            self.table.setItem(index, 0, QTableWidgetItem(mapping['code']))
            item = QTableWidgetItem('••••••••  （已保存）' if mapping.get('target') else '（未设置）')
            item.setData(Qt.ItemDataRole.UserRole, mapping.get('id'))
            self.table.setItem(index, 1, item)

    def _target_prompt(self, title, current=None):
        dialog = MappingEntryDialog(self.kind, current, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None
        try:
            return dialog.values()
        except ValueError as exc:
            QMessageBox.warning(self, '信息不完整', str(exc))
            return None

    def add_row(self):
        result = self._target_prompt('添加映射')
        if result:
            self.rows.append(result)
            self._refresh()

    def edit_row(self):
        index = self.table.currentRow()
        if index < 0:
            return
        result = self._target_prompt('编辑映射', self.rows[index])
        if result:
            self.rows[index] = result
            self._refresh()
            self.table.selectRow(index)

    def delete_row(self):
        index = self.table.currentRow()
        if index >= 0:
            del self.rows[index]
            self._refresh()

    def import_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, '导入旧版映射 CSV', '', 'CSV (*.csv)')
        if not path:
            return
        try:
            imported = self.store.import_csv(path, self.kind)
            by_code = {row['code']: row for row in self.rows}
            by_code.update({row['code']: row for row in imported})
            self.rows = list(by_code.values())
            self._refresh()
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self, '导入失败', str(exc))

    def save(self):
        try:
            self.store.save(self.rows)
        except Exception as exc:
            QMessageBox.warning(self, '保存失败', f'无法保存映射或访问系统凭据库：{exc}')
            return
        self.accept()
