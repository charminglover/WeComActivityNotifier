"""In-app CRUD for shop routing targets; target values are never displayed in clear text."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QAbstractItemView, QDialog, QFileDialog, QHBoxLayout,
    QInputDialog, QMessageBox, QPushButton, QTableWidget, QTableWidgetItem,
    QVBoxLayout)


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
        code, ok = QInputDialog.getText(self, title, '服务店编码：', text=current.get('code', '') if current else '')
        if not ok:
            return None
        is_webhook = self.kind == '群机器人Webhook'
        target, ok = QInputDialog.getText(
            self, title,
            ('Webhook URL（留空表示保持现有凭据）' if is_webhook else 'UserID（多个用 | 分隔，留空表示保持现有凭据）'),
            QInputDialog.EchoMode.Password if is_webhook else QInputDialog.EchoMode.Normal,
            '')
        if not ok:
            return None
        target = target.strip() or (current or {}).get('target', '')
        if not code.strip() or not target:
            QMessageBox.warning(self, '信息不完整', '服务店编码和接收目标均不能为空。')
            return None
        return {'id': (current or {}).get('id'), 'code': code.strip(), 'target': target}

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
