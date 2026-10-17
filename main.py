"""PySide6 desktop application for importing, filtering, and sending Excel records."""
from __future__ import annotations

import threading
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QThread, QThreadPool, Qt, QTimer
from PySide6.QtGui import QAction, QFont, QFontMetrics, QPixmap
from PySide6.QtWidgets import (QApplication, QComboBox, QDialog,
    QFileDialog, QFormLayout, QGroupBox, QHeaderView, QHBoxLayout, QLabel,
    QLineEdit, QListWidget, QListWidgetItem, QMainWindow, QMenu, QMessageBox,
    QProgressBar, QPushButton, QScrollArea, QTabWidget, QTableView,
    QTextEdit, QVBoxLayout, QWidget)

from engine import filter_records, number, read_excel, text_of, WeComSender
from services.notification_workflow import prepare_selection, validate_routes
from ui.table_model import CheckboxDelegate, RecordTableModel
from ui.filter_condition import FilterCondition
from services.task_worker import Worker
from services.config_store import ConfigStore
from services.mapping_store import MappingStore
from ui.mapping_dialog import MappingDialog

APP_NAME = '企业微信活动通知工具'










class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f'{APP_NAME}  v3.0')
        self.resize(1320, 860)
        self.setMinimumSize(1000, 680)
        self.setFont(QFont('Microsoft YaHei UI', 9))
        self.records: list[dict] = []
        self.fields: list[str] = []
        self.filtered: list[dict] = []
        self.checked: set[str] = set()
        self.numeric_filter_fields: set[str] = set()
        self.filter_rows: list[FilterCondition] = []
        self.highlight_rows: list[FilterCondition] = []
        self.filter_generation = 0
        self.selection_generation = 0
        self._fit_after_filter = False
        self.stop_event = threading.Event()
        self.sending = False
        self._next_task_id = 1
        self._workers = {}
        self._closing = False
        self._callbacks = {}
        self.pool = QThreadPool(self)
        self.pool.setMaxThreadCount(max(2, min(3, QThread.idealThreadCount() or 2)))
        self.pool.setExpiryTimeout(1000)
        self.config_store = ConfigStore()
        self.mapping_stores = {
            '群机器人Webhook': MappingStore(kind='群机器人Webhook'),
            '自建应用消息': MappingStore(kind='自建应用消息'),
        }
        self.prefs = self.config_store.load()
        self._build_ui()
        self._update_count()

    def _build_ui(self):
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.setContentsMargins(12, 10, 12, 8)
        file_bar = QHBoxLayout()
        file_bar.addWidget(QLabel('Excel 文件：'))
        self.file_path = QLineEdit()
        self.file_path.setReadOnly(True)
        self.file_path.setPlaceholderText('尚未导入 Excel 或 CSV')
        file_bar.addWidget(self.file_path, 1)
        self.load_button = QPushButton('选择文件…')
        self.load_button.clicked.connect(self.choose_excel)
        file_bar.addWidget(self.load_button)
        layout.addLayout(file_bar)
        self.count_label = QLabel('尚未导入文件')
        layout.addWidget(self.count_label)

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, 1)
        self.data_page = QWidget()
        self.fields_page = QWidget()
        self.send_page = QWidget()
        self.tabs.addTab(self.data_page, '1  数据筛选与记录勾选')
        self.tabs.addTab(self.fields_page, '2  推送内容与格式')
        self.tabs.addTab(self.send_page, '3  企业微信配置与发送')
        self._build_data_page()
        self._build_fields_page()
        self._build_send_page()

        footer = QHBoxLayout()
        self.status_label = QLabel('就绪。请先导入 Excel 文件。')
        footer.addWidget(self.status_label, 1)
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setMaximumWidth(180)
        self.progress.hide()
        footer.addWidget(self.progress)
        layout.addLayout(footer)
        self.setCentralWidget(root)

    def _build_data_page(self):
        layout = QVBoxLayout(self.data_page)
        filter_box = QGroupBox('筛选条件（多条条件按 AND 同时生效）')
        filter_layout = QVBoxLayout(filter_box)
        self.filter_rows_layout = QVBoxLayout()
        self.filter_hint = QLabel('导入 Excel 或 CSV 后，可按任意实际列添加筛选条件。')
        self.filter_rows_layout.addWidget(self.filter_hint)
        filter_layout.addLayout(self.filter_rows_layout)
        actions = QHBoxLayout()
        self.add_filter_button = QPushButton('添加条件')
        self.add_filter_button.setEnabled(False)
        self.add_filter_button.clicked.connect(self.add_filter_condition)
        actions.addWidget(self.add_filter_button)
        apply_button = QPushButton('执行筛选')
        apply_button.clicked.connect(self.apply_filters)
        actions.addWidget(apply_button)
        clear_button = QPushButton('清除条件')
        clear_button.clicked.connect(self.clear_filters)
        actions.addWidget(clear_button)
        actions.addStretch(1)
        filter_layout.addLayout(actions)
        layout.addWidget(filter_box)

        tools = QHBoxLayout()
        for label, mode in (('勾选筛选结果', 'all'), ('反选筛选结果', 'invert'),
                            ('取消所有勾选', 'none')):
            button = QPushButton(label)
            button.clicked.connect(lambda checked=False, m=mode: self.select_filtered(m))
            tools.addWidget(button)
        select_button = QPushButton('勾选高亮行')
        select_button.clicked.connect(self.select_highlight)
        tools.addWidget(select_button)
        width_button = QPushButton('列宽自适应')
        width_button.setToolTip('根据表头和抽样数据调整列宽；也可直接拖动表头分隔线手动调整')
        width_button.clicked.connect(self.auto_fit_columns)
        tools.addWidget(width_button)
        tools.addStretch(1)
        tools.addWidget(QLabel('可直接拖动表头边界调整列宽'))
        layout.addLayout(tools)

        self.table = QTableView()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableView.SelectionMode.ExtendedSelection)
        self.table.setSortingEnabled(False)
        self.table.setWordWrap(False)
        self.table.setShowGrid(False)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)
        header.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        header.customContextMenuRequested.connect(self._column_menu)
        header.sectionDoubleClicked.connect(self._fit_one_column)
        self.table_model = RecordTableModel(self.table)
        self.table.setModel(self.table_model)
        self.table.setItemDelegateForColumn(0, CheckboxDelegate(self.table))
        self.table_model.checkedChanged.connect(self._on_checked_changed)
        layout.addWidget(self.table, 1)

    def _build_fields_page(self):
        layout = QVBoxLayout(self.fields_page)
        layout.addWidget(QLabel('勾选要推送的 Excel 列；列表顺序就是消息中的列顺序。'))
        area = QHBoxLayout()
        self.fields_list = QListWidget()
        self.fields_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.fields_list.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        area.addWidget(self.fields_list, 1)
        buttons = QVBoxLayout()
        for label, callback in (('上移 ↑', lambda: self.move_field(-1)),
                                ('下移 ↓', lambda: self.move_field(1)),
                                ('全部勾选', lambda: self.set_fields(True)),
                                ('全部清空', lambda: self.set_fields(False))):
            button = QPushButton(label)
            button.clicked.connect(callback)
            buttons.addWidget(button)
        buttons.addStretch(1)
        area.addLayout(buttons)
        layout.addLayout(area, 1)

        settings = QGroupBox('通知内容设置')
        form = QFormLayout(settings)
        saved_title = self.prefs.get('msg_title', '')
        if saved_title == '数据通知':
            saved_title = ''
        self.title_edit = QLineEdit(saved_title)
        self.title_edit.setPlaceholderText(
            '留空时自动生成；可用 {group_field} 和 {group_value}，例如“{group_value}活动通知”')
        form.addRow('通知标题 / 模板', self.title_edit)
        self.group_field_combo = QComboBox()
        self.group_field_combo.addItem('不分组', '')
        self.group_field_combo.setToolTip('选择任意导入列作为分组依据；映射发送时仍会按服务店分开发送。')
        form.addRow('分组字段', self.group_field_combo)
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(['图片表格', 'Markdown表格'])
        self.mode_combo.setCurrentText(self.prefs.get('mode', '图片表格'))
        form.addRow('发送格式', self.mode_combo)
        preview_button = QPushButton('校验并预览消息')
        preview_button.clicked.connect(self.preview)
        form.addRow(preview_button)
        layout.addWidget(settings)

        highlight_box = QGroupBox('重点标注条件（可选，多条条件按 AND 生效）')
        highlight_layout = QVBoxLayout(highlight_box)
        self.highlight_rows_layout = QVBoxLayout()
        self.highlight_hint = QLabel('可按任意导入列设置文本或数值条件；符合条件的记录会在消息中标注。')
        self.highlight_rows_layout.addWidget(self.highlight_hint)
        highlight_layout.addLayout(self.highlight_rows_layout)
        add_highlight = QPushButton('添加标注条件')
        add_highlight.clicked.connect(self.add_highlight_condition)
        highlight_layout.addWidget(add_highlight)
        layout.addWidget(highlight_box)

    def _build_send_page(self):
        layout = QVBoxLayout(self.send_page)
        section = QGroupBox('企业微信发送方式')
        form = QFormLayout(section)
        self.kind_combo = QComboBox()
        self.kind_combo.addItems(['群机器人Webhook', '自建应用消息'])
        self.kind_combo.setCurrentText(self.prefs.get('kind', '群机器人Webhook'))
        self.kind_combo.currentTextChanged.connect(self._toggle_sender)
        form.addRow('发送方式', self.kind_combo)
        self.webhook_edit = QLineEdit()
        self.webhook_edit.setPlaceholderText('https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=…')
        form.addRow('Webhook URL', self.webhook_edit)
        self.corpid_edit = QLineEdit()
        self.secret_edit = QLineEdit()
        self.secret_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.agentid_edit = QLineEdit()
        self.touser_edit = QLineEdit()
        form.addRow('CorpID', self.corpid_edit)
        form.addRow('Secret', self.secret_edit)
        form.addRow('AgentID', self.agentid_edit)
        form.addRow('接收人 UserID（用 | 分隔）', self.touser_edit)
        layout.addWidget(section)
        self._toggle_sender(self.kind_combo.currentText())

        mapping = QGroupBox('服务店接收目标映射（可选）')
        map_layout = QHBoxLayout(mapping)
        map_layout.addWidget(QLabel('在软件内添加、编辑和保存接收目标；凭据使用系统凭据库。'))
        map_layout.addStretch(1)
        map_button = QPushButton('管理映射…')
        map_button.clicked.connect(self.manage_mappings)
        map_layout.addWidget(map_button)
        layout.addWidget(mapping)

        actions = QHBoxLayout()
        self.send_button = QPushButton('确认并发送')
        self.send_button.clicked.connect(self.send_now)
        actions.addWidget(self.send_button)
        self.stop_button = QPushButton('停止后续发送')
        self.stop_button.setEnabled(False)
        self.stop_button.clicked.connect(self.stop_sending)
        actions.addWidget(self.stop_button)
        preview_button = QPushButton('预览消息')
        preview_button.clicked.connect(self.preview)
        actions.addWidget(preview_button)
        actions.addStretch(1)
        layout.addLayout(actions)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.document().setMaximumBlockCount(3000)
        layout.addWidget(self.log, 1)

    def _toggle_sender(self, kind):
        bot = kind == '群机器人Webhook'
        for widget in (self.webhook_edit,):
            widget.setVisible(bot)
        for widget in (self.corpid_edit, self.secret_edit, self.agentid_edit, self.touser_edit):
            widget.setVisible(not bot)
        # Hide their form labels as well by changing row visibility.
        form = self.webhook_edit.parentWidget().layout()
        if isinstance(form, QFormLayout):
            form.setRowVisible(self.webhook_edit, bot)
            for widget in (self.corpid_edit, self.secret_edit, self.agentid_edit, self.touser_edit):
                form.setRowVisible(widget, not bot)

    def choose_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择 Excel 文件', '',
            'Excel / CSV (*.xlsx *.xlsm *.xls *.csv);;所有文件 (*.*)')
        if not path:
            return
        self.file_path.setText(path)
        self.load_button.setEnabled(False)
        def load(progress, cancel_event):
            fields, records = read_excel(path, cancel_event=cancel_event)
            step = max(1, (len(records) + 499) // 500)
            sample = records[::step][:500]
            numeric_fields = {field for field in fields
                              if any(number(row.get(field)) is not None for row in sample)}
            return fields, records, numeric_fields

        self._run_task(load, self._loaded,
                       '正在解析 Excel，请稍候…', error_title='导入失败')

    def _run_task(self, function, on_success, status, error_title='执行失败'):
        task_id = self._next_task_id
        self._next_task_id += 1
        worker = Worker(task_id, function)
        self._workers[task_id] = worker
        self._callbacks[task_id] = (on_success, error_title)
        worker.signals.succeeded.connect(self._task_succeeded)
        worker.signals.failed.connect(self._task_failed)
        worker.signals.cancelled.connect(self._task_cancelled)
        worker.signals.progress.connect(self._task_progress)
        self.status_label.setText(status)
        self.progress.show()
        self.pool.start(worker)

    def _task_succeeded(self, task_id, result):
        worker = self._workers.pop(task_id, None)
        callback, _ = self._callbacks.pop(task_id, (None, None))
        if not self._workers:
            self.progress.hide()
            self.status_label.setText('就绪')
        if worker is not None and callback and not self._closing:
            callback(result)
        self.load_button.setEnabled(True)
        if self._closing and not self._workers:
            QTimer.singleShot(0, self.close)

    def _task_failed(self, task_id, error):
        self._workers.pop(task_id, None)
        _, title = self._callbacks.pop(task_id, (None, '执行失败'))
        if not self._workers:
            self.progress.hide()
            self.status_label.setText('就绪')
        self.load_button.setEnabled(True)
        if not self._closing:
            QMessageBox.warning(self, title, error)
        if self._closing and not self._workers:
            QTimer.singleShot(0, self.close)

    def _task_cancelled(self, task_id):
        self._workers.pop(task_id, None)
        self._callbacks.pop(task_id, None)
        if not self._workers:
            self.progress.hide()
            self.status_label.setText('就绪')
            if self._closing:
                QTimer.singleShot(0, self.close)

    def _task_progress(self, message):
        self.status_label.setText(message)
        self._log(message)

    def _loaded(self, result):
        fields, records, numeric_fields = result
        self.fields = list(fields)
        self.records = records
        self.filtered = []
        self.checked.clear()
        self.numeric_filter_fields = numeric_fields
        self._fill_field_list()
        self._fill_group_fields()
        self.clear_filter_conditions()
        self.clear_highlight_conditions()
        for condition in self.prefs.get('highlight_conditions', []):
            self.add_highlight_condition(condition)
        self.add_filter_button.setEnabled(bool(fields))
        self.filter_hint.setText('已识别文件列，可添加任意列的筛选条件。')
        self.filter_hint.show()
        self.table_model.checked = self.checked
        self.table_model.set_content(fields, [], self.checked)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 58)
        for col in range(1, len(fields) + 1):
            self.table.setColumnWidth(col, 150)
        self._fit_after_filter = True
        self.apply_filters()
        self._log(f'导入成功：{len(records):,} 行，{len(fields)} 列。')

    def _fill_field_list(self):
        saved = self.prefs.get('fields') if self.prefs.get('fields_confirmed', False) else []
        saved_set = set(saved or [])
        order = [field for field in self.prefs.get('order', []) if field in self.fields]
        order.extend(field for field in self.fields if field not in order)
        self.fields_list.clear()
        for field in order:
            item = QListWidgetItem(field)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsDragEnabled)
            item.setCheckState(Qt.CheckState.Checked if field in saved_set else Qt.CheckState.Unchecked)
            self.fields_list.addItem(item)

    def _fill_group_fields(self):
        self.group_field_combo.blockSignals(True)
        self.group_field_combo.clear()
        self.group_field_combo.addItem('不分组', '')
        # Store the field as item data: currentData() is used by snapshots and
        # titles, while addItems() would leave it as None for every entry.
        for field in self.fields:
            self.group_field_combo.addItem(field, field)
        selected = self.prefs.get('group_field', '')
        if not selected and self.prefs.get('group_by_shop') and '任务服务店编码' in self.fields:
            selected = '任务服务店编码'
        index = self.group_field_combo.findData(selected)
        self.group_field_combo.setCurrentIndex(max(0, index))
        self.group_field_combo.blockSignals(False)

    def add_highlight_condition(self, condition=None):
        if not self.fields:
            return
        self.highlight_hint.hide()
        row = FilterCondition(self.fields, self.numeric_filter_fields)
        row.removed.connect(self._remove_highlight_condition)
        self.highlight_rows_layout.addWidget(row)
        self.highlight_rows.append(row)
        if condition:
            index = row.field.findText(condition.get('field', ''))
            if index > 0:
                row.field.setCurrentIndex(index)
                row.operator.setCurrentText(condition.get('operator', ''))
                row.value.setText(str(condition.get('value', '')))

    def _remove_highlight_condition(self, row):
        if row in self.highlight_rows:
            self.highlight_rows.remove(row)
        row.deleteLater()
        if not self.highlight_rows:
            self.highlight_hint.show()

    def clear_highlight_conditions(self):
        for row in self.highlight_rows:
            self.highlight_rows_layout.removeWidget(row)
            row.deleteLater()
        self.highlight_rows.clear()
        if hasattr(self, 'highlight_hint'):
            self.highlight_hint.show()

    def _collect_highlight_conditions(self, strict=True):
        conditions = []
        for row in self.highlight_rows:
            try:
                condition = row.condition()
            except ValueError:
                if strict:
                    raise
                continue
            if condition:
                conditions.append(condition)
        return conditions

    def add_filter_condition(self):
        if not self.fields:
            return
        self.filter_hint.hide()
        row = FilterCondition(self.fields, self.numeric_filter_fields)
        row.removed.connect(self._remove_filter_condition)
        row.changed.connect(lambda: None)
        self.filter_rows_layout.addWidget(row)
        self.filter_rows.append(row)

    def _remove_filter_condition(self, row):
        if row in self.filter_rows:
            self.filter_rows.remove(row)
        row.deleteLater()
        if not self.filter_rows:
            self.filter_hint.show()

    def clear_filter_conditions(self):
        for row in self.filter_rows:
            self.filter_rows_layout.removeWidget(row)
            row.deleteLater()
        self.filter_rows.clear()
        if hasattr(self, 'filter_hint'):
            self.filter_hint.show()

    def clear_filters(self):
        self.clear_filter_conditions()
        self.apply_filters()

    def _collect_conditions(self):
        conditions = []
        for row in self.filter_rows:
            condition = row.condition()
            if condition:
                conditions.append(condition)
        return conditions

    def apply_filters(self):
        if not self.records:
            return
        try:
            conditions = self._collect_conditions()
        except ValueError as exc:
            QMessageBox.warning(self, '筛选条件', str(exc))
            return
        self.filter_generation += 1
        generation = self.filter_generation
        rows = self.records

        def finished(result):
            if generation != self.filter_generation:
                return
            self.filtered = result
            self.table_model.set_content(self.fields, self.filtered, self.checked)
            self._update_count()
            if self._fit_after_filter:
                self._fit_after_filter = False
                QTimer.singleShot(0, self.auto_fit_columns)

        self._run_task(lambda progress, cancel: filter_records(rows, conditions=conditions, cancel_event=cancel), finished,
                       '正在筛选数据…', error_title='筛选失败')

    def auto_fit_columns(self, column=None):
        if not self.fields:
            return
        columns = [column] if column is not None else range(1, len(self.fields) + 1)
        metrics = QFontMetrics(self.table.font())
        rows = self.filtered
        if len(rows) > 80:
            step = max(1, len(rows) // 80)
            sample = rows[::step][:80]
        else:
            sample = rows
        for col in columns:
            if col <= 0 or col > len(self.fields):
                continue
            field = self.fields[col - 1]
            width = metrics.horizontalAdvance(field) + 34
            for record in sample:
                value = text_of(record.get(field), field)
                width = max(width, metrics.horizontalAdvance(value[:120]) + 28)
            self.table.setColumnWidth(col, max(80, min(620, width)))
        self.table.setColumnWidth(0, 58)

    def _fit_one_column(self, logical_index):
        if logical_index > 0:
            self.auto_fit_columns(logical_index)

    def _column_menu(self, position):
        header = self.table.horizontalHeader()
        logical = header.logicalIndexAt(position)
        menu = QMenu(self)
        fit_one = QAction('自适应此列', menu)
        fit_one.setEnabled(logical > 0)
        fit_one.triggered.connect(lambda: self.auto_fit_columns(logical))
        menu.addAction(fit_one)
        fit_all = QAction('所有列自适应', menu)
        fit_all.triggered.connect(lambda: self.auto_fit_columns())
        menu.addAction(fit_all)
        menu.exec(header.mapToGlobal(position))

    def _update_count(self):
        if not hasattr(self, 'count_label'):
            return
        visible = len(self.filtered)
        self.count_label.setText(
            f'数据总计 {len(self.records):,} 条  |  当前筛选 {visible:,} 条  |  已勾选 {len(self.checked):,} 条')

    def _on_checked_changed(self):
        self.selection_generation += 1
        self._update_count()

    def select_filtered(self, mode):
        if not self.records:
            return
        rows = self.filtered
        checked = self.checked
        self.selection_generation += 1
        generation = self.selection_generation

        def task(progress, cancel_event):
            result = set(checked)
            keys = (row['_id'] for row in rows)
            if mode == 'all':
                result.update(keys)
            elif mode == 'none':
                result.clear()
            else:
                result.symmetric_difference_update(keys)
            return result

        def finished(result):
            if generation != self.selection_generation:
                return
            self.checked = result
            self.table_model.checked = self.checked
            self._refresh_visible_checks()

        self._run_task(task, finished, '正在更新勾选状态…', error_title='更新失败')

    def select_highlight(self):
        indexes = self.table.selectionModel().selectedRows()
        for index in indexes:
            key = self.filtered[index.row()].get('_id')
            if key:
                self.checked.add(key)
        self._refresh_visible_checks()

    def _refresh_visible_checks(self):
        if not self.filtered:
            self.table_model.checkedChanged.emit()
            return
        top = self.table.indexAt(self.table.viewport().rect().topLeft()).row()
        bottom = self.table.indexAt(self.table.viewport().rect().bottomRight()).row()
        if top < 0:
            top = 0
        if bottom < top:
            bottom = min(len(self.filtered) - 1, top + 80)
        self.table_model.refresh_checks(top, bottom)

    def move_field(self, delta):
        row = self.fields_list.currentRow()
        target = row + delta
        if row < 0 or not 0 <= target < self.fields_list.count():
            return
        item = self.fields_list.takeItem(row)
        self.fields_list.insertItem(target, item)
        self.fields_list.setCurrentRow(target)

    def set_fields(self, enabled):
        state = Qt.CheckState.Checked if enabled else Qt.CheckState.Unchecked
        for index in range(self.fields_list.count()):
            self.fields_list.item(index).setCheckState(state)

    def _snapshot(self):
        if not self.checked:
            raise ValueError('请在第 1 页勾选至少一条活动记录')
        fields = [self.fields_list.item(i).text() for i in range(self.fields_list.count())
                  if self.fields_list.item(i).checkState() == Qt.CheckState.Checked]
        if not fields:
            raise ValueError('请在第 2 页勾选至少一个推送字段')
        try:
            highlight_conditions = self._collect_highlight_conditions()
        except ValueError as exc:
            raise ValueError('请检查重点标注条件：' + str(exc)) from exc
        group_field = self.group_field_combo.currentData() or ''
        title_template = self.title_edit.text().strip()
        if not title_template:
            title_template = '{group_value}通知' if group_field else '数据通知'
        mapping_enabled = self.mapping_stores[self.kind_combo.currentText()].has_mappings()
        return {'ids': set(self.checked), 'fields': fields,
                'mode': self.mode_combo.currentText(),
                'group_field': group_field, 'route_by_store': mapping_enabled,
                'title': title_template, 'highlight_conditions': highlight_conditions,
                'records': self.records}

    def preview(self):
        try:
            snapshot = self._snapshot()
        except ValueError as exc:
            QMessageBox.warning(self, '无法预览', str(exc))
            return

        def task(progress, cancel_event):
            messages = prepare_selection(snapshot)
            return messages[0].payload, snapshot['mode'], len(messages)

        self._run_task(task, self._show_preview, '正在后台校验并生成消息预览…',
                       error_title='无法预览')

    def _show_preview(self, result):
        preview, mode, total = result
        dialog = QDialog(self)
        dialog.setWindowTitle(f'消息预览 - 自动拆分为 {total} 条消息')
        dialog.resize(1050, 740)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(f'已逐条校验，预计 {total} 条消息。这里显示第一条；只有确认发送后才会推送。'))
        if mode == '图片表格':
            from PySide6.QtWidgets import QScrollArea
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            label = QLabel()
            label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            pixmap = QPixmap()
            pixmap.loadFromData(preview)
            label.setPixmap(pixmap)
            scroll.setWidget(label)
            layout.addWidget(scroll, 1)
            save = QPushButton('保存与实际推送相同的 JPEG…')
            save.clicked.connect(lambda: self._save_preview(dialog, preview))
            layout.addWidget(save)
        else:
            text = QTextEdit()
            text.setReadOnly(True)
            text.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
            text.setPlainText(preview)
            layout.addWidget(text, 1)
        dialog.exec()

    def _save_preview(self, dialog, payload):
        path, _ = QFileDialog.getSaveFileName(dialog, '保存预览图片', '', 'JPEG 图片 (*.jpg)')
        if path:
            if not path.lower().endswith('.jpg'):
                path += '.jpg'
            Path(path).write_bytes(payload)
            QMessageBox.information(dialog, '已保存', path)

    def manage_mappings(self):
        dialog = MappingDialog(self.mapping_stores[self.kind_combo.currentText()], self.kind_combo.currentText(), self)
        if dialog.exec():
            self._log('服务店接收映射已保存。')

    def send_now(self):
        try:
            snapshot = self._snapshot()
            kind = self.kind_combo.currentText()
            if kind == '自建应用消息' and snapshot['mode'] == 'Markdown表格':
                raise ValueError('自建应用消息不支持 Markdown V2 表格。请使用图片表格。')
            options = {'kind': kind, 'webhook': self.webhook_edit.text(),
                       'corp': self.corpid_edit.text(), 'secret': self.secret_edit.text(),
                       'agent': self.agentid_edit.text(), 'user': self.touser_edit.text(),
                       'mapping_enabled': self.mapping_stores[self.kind_combo.currentText()].has_mappings()}
        except ValueError as exc:
            QMessageBox.warning(self, '发送配置', str(exc))
            return

        def prepare(progress, cancel_event):
            prepared = prepare_selection(snapshot)
            mapping = self.mapping_stores[kind].as_dict() if options['mapping_enabled'] else {}
            validate_routes([message.job for message in prepared], mapping, kind,
                            options['webhook'], options['user'])
            if kind == '自建应用消息' and not all((options['corp'].strip(),
                    options['secret'].strip(), options['agent'].strip())):
                raise ValueError('自建应用需要填写 CorpID、Secret、AgentID')
            return prepared, mapping, options, snapshot['highlight_conditions']

        self._run_task(prepare, self._confirm_and_send,
                       '正在后台校验全部消息与接收目标…', error_title='发送配置')

    def _confirm_and_send(self, plan):
        prepared, mapping, options, highlight_conditions = plan
        kind = options['kind']
        mode = prepared[0].mode
        recipients = len({mapping[item.job.store_code] for item in prepared}) if mapping else 1
        record_keys = {row['_id'] for item in prepared for row in item.job.rows}
        shop_codes = {text_of(row.get('任务服务店编码'))
                      for item in prepared for row in item.job.rows}
        confirm = (f'发送格式：{mode}（已完整校验）\n目标模式：{kind}\n'
                   f'本次勾选活动记录：{len(record_keys)} 条\n预计发送消息：{len(prepared)} 条\n'
                   f'接收目标：{recipients} 个\n\n请确认接收对象正确。发送后不可撤回。')
        if not mapping and len(shop_codes) > 1:
            confirm += '\n\n注意：未设置接收映射，多家服务店的数据将发送到同一目标。'
        answer = QMessageBox.warning(self, '正式发送前确认', confirm,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.stop_event.clear()
        self.sending = True
        self.send_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self._log(f'开始发送 {len(prepared)} 条已校验消息。')

        def process(progress, cancel_event):
            ok = 0
            failed = 0
            try:
                sender = WeComSender(kind, webhook=options['webhook'], corpid=options['corp'],
                    secret=options['secret'], agentid=options['agent'], touser=options['user'])
                for index, item in enumerate(prepared, 1):
                    if self.stop_event.is_set() or cancel_event.is_set():
                        progress('操作已停止：后续未发送的消息已取消')
                        break
                    job = item.job
                    recipient = mapping.get(job.store_code, '') if mapping else ''
                    try:
                        sender.send(item, mode, recipient, highlight_conditions=highlight_conditions)
                        ok += 1
                        progress(f'[{index}/{len(prepared)}] 发送成功 · {job.store_name or "汇总"} · 第 {job.part}/{job.parts} 页')
                    except Exception as exc:
                        failed += 1
                        error = str(exc)
                        for secret in (options['webhook'], options['secret'],
                                       recipient if kind == '群机器人Webhook' else ''):
                            if secret:
                                error = error.replace(secret, '[已隐藏]')
                        progress(f'[{index}/{len(prepared)}] 失败 · {job.store_name or "汇总"} · {error}')
                    if index < len(prepared) and kind == '群机器人Webhook':
                        self.stop_event.wait(3.2)
            except Exception as exc:
                error = str(exc)
                for secret in (options['secret'], options['webhook']):
                    if secret:
                        error = error.replace(secret, '[已隐藏]')
                progress('发送初始化失败：' + error)
            return ok, len(prepared), failed

        self._run_task(process, self._send_finished, '正在发送…', error_title='发送失败')

    def _send_finished(self, result):
        success, total, failed = result
        self.sending = False
        self.send_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self._log(f'发送结束：成功 {success}/{total} 条，失败 {failed} 条。')
        QMessageBox.information(self, '发送结束',
                                f'成功 {success}/{total} 条，失败 {failed} 条。\n详细内容请查看日志。')

    def stop_sending(self):
        self.stop_event.set()
        self.status_label.setText('正在停止…当前网络请求结束后不再发送后续消息')

    def _log(self, line):
        self.log.append(f'{datetime.now().strftime("%H:%M:%S")}  {line}')

    def closeEvent(self, event):
        if self._workers:
            answer = QMessageBox.question(
                self, '后台任务仍在执行', '正在请求停止后台任务。当前网络请求结束后会自动关闭，是否继续？',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes)
            if answer != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
            self._closing = True
            self.stop_event.set()
            for worker in self._workers.values():
                worker.cancel_event.set()
            self.status_label.setText('正在取消后台任务，完成后自动关闭…')
            event.ignore()
            return
        try:
            data = {'fields': [self.fields_list.item(i).text()
                               for i in range(self.fields_list.count())
                               if self.fields_list.item(i).checkState() == Qt.CheckState.Checked],
                    'order': [self.fields_list.item(i).text()
                              for i in range(self.fields_list.count())],
                    'fields_confirmed': True,
                    'msg_title': self.title_edit.text(), 'mode': self.mode_combo.currentText(),
                    'group_field': self.group_field_combo.currentData() or '',
                    'highlight_conditions': self._collect_highlight_conditions(strict=False),
                    'kind': self.kind_combo.currentText()}
            self.config_store.save(data)
        except (OSError, ValueError):
            pass
        event.accept()


if __name__ == '__main__':
    application = QApplication([])
    application.setApplicationName(APP_NAME)
    application.setStyle('Fusion')
    window = App()
    window.show()
    application.exec()
