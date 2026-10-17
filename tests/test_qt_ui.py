import os
import unittest

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QHeaderView, QTableView
from ui.table_model import CheckboxDelegate
from ui.mapping_dialog import MappingEntryDialog

from main import FilterCondition, RecordTableModel


class TestQtTable(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

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

    def test_mapping_entry_dialog_add_and_preserve_existing_target(self):
        dialog = MappingEntryDialog('群机器人Webhook')
        dialog.code.setText('REGION-NORTH')
        dialog.target.setText('https://example.test/hook')
        self.assertEqual(dialog.values(), {
            'id': None, 'code': 'REGION-NORTH', 'target': 'https://example.test/hook'})

        edited = MappingEntryDialog('群机器人Webhook', {
            'id': 'saved-id', 'code': 'OLD', 'target': 'https://example.test/old'})
        edited.code.setText('NEW')
        self.assertEqual(edited.values(), {
            'id': 'saved-id', 'code': 'NEW', 'target': 'https://example.test/old'})


if __name__ == '__main__':
    unittest.main()
