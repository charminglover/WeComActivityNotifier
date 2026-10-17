import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from PySide6.QtCore import QCoreApplication, QEvent, QThreadPool, Qt, QTimer
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QDialog, QHeaderView, QLineEdit, QTableView
from ui.table_model import CheckboxDelegate
from ui.mapping_dialog import MappingDialog, MappingEntryDialog
from services.mapping_store import MappingStore
from services.config_store import ConfigStore
from services.notification_workflow import prepare_selection

from main import App, FilterCondition, RecordTableModel


class TestQtTable(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def tearDown(self):
        # Drain asynchronous persistence before deleting test windows. Relying
        # on Python interpreter shutdown leaves Qt objects and worker callbacks
        # alive in an undefined destruction order on Windows.
        self.assertTrue(QThreadPool.globalInstance().waitForDone(5000))
        self.app.processEvents()
        for widget in self.app.topLevelWidgets():
            widget.deleteLater()
        QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
        self.app.processEvents()

    def test_virtual_model_and_check_state(self):
        model = RecordTableModel()
        rows = [
            {'Region': 'North', 'Revenue': 1200, '_id': 'one'},
            {'Region': 'South', 'Revenue': 900, '_id': 'two'},
        ]
        model.set_content(['Region', 'Revenue'], rows, set())
        self.assertEqual(model.rowCount(), 2)
        self.assertEqual(model.columnCount(), 3)  # checkbox + source columns
        index = model.index(0, 0)
        self.assertTrue(model.setData(index, Qt.CheckState.Checked,
                                      Qt.ItemDataRole.CheckStateRole))
        self.assertEqual(model.data(index, Qt.ItemDataRole.CheckStateRole),
                         Qt.CheckState.Checked)


    def test_checkbox_cell_click_toggles_single_record(self):
        table = QTableView()
        model = RecordTableModel()
        rows = [{'Region': 'North', '_id': 'one'}, {'Region': 'South', '_id': 'two'}]
        model.set_content(['Region'], rows, set())
        table.setModel(model)
        table.setItemDelegateForColumn(0, CheckboxDelegate(table))
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.resize(420, 180)
        table.show()
        self.app.processEvents()
        index = model.index(0, 0)
        QTest.mouseClick(table.viewport(), Qt.MouseButton.LeftButton,
                         pos=table.visualRect(index).center())
        self.assertEqual(model.checked, {'one'})
        self.assertEqual(model.data(index, Qt.ItemDataRole.CheckStateRole), Qt.CheckState.Checked)
        self.assertEqual(model.data(model.index(1, 0), Qt.ItemDataRole.CheckStateRole),
                         Qt.CheckState.Unchecked)
        QTest.mouseClick(table.viewport(), Qt.MouseButton.LeftButton,
                         pos=table.visualRect(index).center())
        self.assertEqual(model.checked, set(), '同一条记录再次点击应取消勾选')
        table.close()

    def test_generic_filter_row_and_interactive_column_width(self):
        row = FilterCondition(['Region', 'Revenue'], {'Revenue'})
        row.field.setCurrentIndex(1)
        row.operator.setCurrentText('等于')
        row.value.setText('North')
        self.assertEqual(row.condition(), {
            'field': 'Region', 'operator': '等于', 'value': 'North'})

        table = QTableView()
        model = RecordTableModel()
        model.set_content(['Region'], [{'Region': 'North', '_id': 'one'}], set())
        table.setModel(model)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.setColumnWidth(1, 240)
        self.assertEqual(table.horizontalHeader().sectionResizeMode(1),
                         QHeaderView.ResizeMode.Interactive)
        self.assertEqual(table.columnWidth(1), 240)

    def test_mapping_entry_has_visible_editable_webhook(self):
        dialog = MappingEntryDialog('群机器人Webhook')
        dialog.match_value.setText('REGION-NORTH')
        dialog.target.setText('https://example.test/hook')
        self.assertEqual(dialog.values(), {
            'id': None, 'value': 'REGION-NORTH', 'target': 'https://example.test/hook'})
        self.assertEqual(dialog.target.echoMode(), QLineEdit.EchoMode.Normal)

        edited = MappingEntryDialog('群机器人Webhook', {
            'id': 'saved-id', 'value': 'OLD', 'target': 'https://example.test/old'})
        edited.match_value.setText('NEW')
        self.assertEqual(edited.target.text(), 'https://example.test/old')
        self.assertEqual(edited.values(), {
            'id': 'saved-id', 'value': 'NEW', 'target': 'https://example.test/old'})

    def test_mapping_buttons_add_edit_save_reopen_and_delete(self):
        with tempfile.TemporaryDirectory() as directory:
            secrets = {}
            backend = Mock()
            backend.set_password.side_effect = lambda service, key, target: secrets.update({key: target})
            backend.get_password.side_effect = lambda service, key: secrets.get(key)
            backend.delete_password.side_effect = lambda service, key: secrets.pop(key, None)
            store = MappingStore(Path(directory) / 'mapping.json', backend, kind='群机器人Webhook')
            dialog = MappingDialog(store, '群机器人Webhook', fields=['Region', 'Department'])
            dialog.field_combo.setCurrentText('Region')
            dialog.show()
            errors = []

            def fill_entry(value, target):
                entry = QApplication.activeModalWidget()
                try:
                    entry.match_value.setText(value)
                    entry.target.setText(target)
                    QTest.mouseClick(entry.ok_button, Qt.MouseButton.LeftButton)
                except Exception as exc:
                    errors.append(exc)
                    if entry:
                        entry.reject()

            url = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=ui-test'
            QTimer.singleShot(0, lambda: fill_entry('North', url))
            QTest.mouseClick(dialog.action_buttons[0], Qt.MouseButton.LeftButton)
            self.assertFalse(errors)
            self.assertEqual(dialog.table.item(0, 1).text(), url)
            QTimer.singleShot(0, lambda: fill_entry('East', url + '-edited'))
            QTest.mouseClick(dialog.action_buttons[1], Qt.MouseButton.LeftButton)
            self.assertFalse(errors)
            self.assertEqual(dialog.table.item(0, 0).text(), 'East')
            QTest.mouseClick(dialog.save_button, Qt.MouseButton.LeftButton)
            deadline = time.monotonic() + 5
            while dialog._worker is not None and time.monotonic() < deadline:
                QTest.qWait(10)
            self.assertIsNone(dialog._worker)
            self.assertEqual(dialog.result(), QDialog.DialogCode.Accepted)
            self.assertEqual(store.as_dict(), {'East': url + '-edited'})

            reopened = MappingDialog(store, '群机器人Webhook', fields=['Region'])
            self.assertEqual(reopened.field_combo.currentText(), 'Region')
            self.assertEqual(reopened.table.item(0, 1).text(), url + '-edited')
            reopened.table.selectRow(0)
            QTest.mouseClick(reopened.action_buttons[2], Qt.MouseButton.LeftButton)
            reopened.save()
            deadline = time.monotonic() + 5
            while reopened._worker is not None and time.monotonic() < deadline:
                QTest.qWait(10)
            self.assertIsNone(reopened._worker)
            self.assertFalse(store.has_mappings())
            self.assertFalse(secrets)

    def test_app_snapshot_uses_mapping_column_separate_from_grouping(self):
        with tempfile.TemporaryDirectory() as directory:
            backend = Mock()
            stores = {kind: MappingStore(Path(directory) / f'{index}.json', backend, kind=kind)
                      for index, kind in enumerate(('群机器人Webhook', '自建应用消息'))}
            store = stores['群机器人Webhook']
            store.save([{'value': 'North', 'target':
                         'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=snapshot'}], field='Region')
            config = ConfigStore(Path(directory) / 'prefs.json')
            with patch('main.ConfigStore', return_value=config), patch(
                    'main.MappingStore', side_effect=lambda kind: stores[kind]):
                window = App()
                window.fields = ['Region', 'Department', 'Description']
                window.records = [{'_id': 'r1', 'Region': 'North',
                                   'Department': 'Sales', 'Description': 'Data'}]
                window.checked = {'r1'}
                window._fill_field_list()
                window._fill_group_fields()
                window.set_fields(True)
                window.group_field_combo.setCurrentText('Department')
                snapshot = window._snapshot()
                self.assertEqual(snapshot['route_field'], 'Region')
                self.assertEqual(snapshot['group_field'], 'Department')
                window.include_excel_check.setChecked(True)
                self.assertTrue(window._snapshot()['include_excel'])
                window.mode_combo.setCurrentText('Excel附件')
                self.assertFalse(window.include_excel_check.isEnabled())
                snapshot = window._snapshot()
                self.assertFalse(snapshot['include_excel'])
                prepared = prepare_selection(snapshot)
                preview_state = []

                def inspect_preview():
                    dialog = QApplication.activeModalWidget()
                    try:
                        table = dialog.findChild(QTableView)
                        preview_state.append((table.model().rowCount(), table.isColumnHidden(0)))
                    finally:
                        dialog.reject()

                QTimer.singleShot(0, inspect_preview)
                window._show_preview({'first': prepared[0], 'attachment': prepared[0],
                                      'files': 1, 'total': 1})
                self.assertEqual(preview_state, [(1, True)])
                window.fields.remove('Region')
                with self.assertRaisesRegex(ValueError, '当前文件缺少映射匹配列'):
                    window._snapshot()
                window.close()


if __name__ == '__main__':
    unittest.main()
