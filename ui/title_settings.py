"""Visual dynamic-title settings driven by the actual imported headers."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QComboBox, QHBoxLayout, QLabel, QLineEdit,
                              QPushButton, QVBoxLayout, QWidget)

from engine import text_of
from services.title_rules import compile_title_rule, field_token


class TitleSettings(QWidget):
    def __init__(self, config=None, legacy_title='', parent=None):
        super().__init__(parent)
        if not config:
            config = ({'mode': 'custom', 'template': legacy_title}
                      if legacy_title and legacy_title not in ('数据通知', '售后活动执行进度通知')
                      else {'mode': 'auto'})
        self.fields = []
        self.sample_rows = []
        self.group_field = ''
        self.route_field = ''
        self._preferred_field = config.get('field', '')
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.mode = QComboBox()
        for label, value in (('自动跟随分组 / 接收映射', 'auto'),
                             ('按指定 Excel 列生成', 'field'),
                             ('自定义文字 / 多列组合', 'custom')):
            self.mode.addItem(label, value)
        self.mode.setCurrentIndex(max(0, self.mode.findData(config.get('mode', 'auto'))))
        self.mode.setToolTip('自动模式优先使用分组字段；未分组时使用接收映射的匹配字段。也可指定任意标题列或组合多列。')
        layout.addWidget(self.mode)

        self.dynamic_controls = QWidget()
        dynamic = QHBoxLayout(self.dynamic_controls)
        dynamic.setContentsMargins(0, 0, 0, 0)
        self.field = QComboBox()
        self.field.setMinimumWidth(180)
        self.field.addItem('选择标题来源列…', '')
        self.prefix = QLineEdit(str(config.get('prefix', '')))
        self.prefix.setPlaceholderText('前缀（可留空）')
        self.suffix = QLineEdit(str(config.get('suffix', '通知')))
        self.suffix.setPlaceholderText('后缀（可留空）')
        dynamic.addWidget(self.prefix, 1)
        dynamic.addWidget(self.field, 2)
        dynamic.addWidget(self.suffix, 1)
        layout.addWidget(self.dynamic_controls)

        self.custom_controls = QWidget()
        custom = QVBoxLayout(self.custom_controls)
        custom.setContentsMargins(0, 0, 0, 0)
        toolbar = QHBoxLayout()
        self.insert_field = QComboBox()
        self.insert_field.addItem('选择要插入的 Excel 列…', '')
        self.insert_button = QPushButton('插入列值')
        self.insert_button.clicked.connect(self._insert_field)
        toolbar.addWidget(self.insert_field, 1)
        toolbar.addWidget(self.insert_button)
        custom.addLayout(toolbar)
        self.template = QLineEdit(str(config.get('template', '')))
        self.template.setPlaceholderText('可直接写固定标题，也可插入列值并添加文字或分隔符')
        custom.addWidget(self.template)
        layout.addWidget(self.custom_controls)
        self.preview = QLabel()
        self.preview.setWordWrap(True)
        self.preview.setTextFormat(Qt.TextFormat.PlainText)
        layout.addWidget(self.preview)
        self.hint = QLabel('标题使用的列值不同时会自动分别生成消息；标题、图片和附件名称使用同一规则。')
        self.hint.setWordWrap(True)
        layout.addWidget(self.hint)
        self.mode.currentIndexChanged.connect(self._refresh)
        self.field.currentIndexChanged.connect(self._refresh)
        self.prefix.textChanged.connect(self._refresh)
        self.suffix.textChanged.connect(self._refresh)
        self.template.textChanged.connect(self._refresh)
        self._refresh()

    def set_fields(self, fields, sample_rows):
        selected = self.field.currentData() or self._preferred_field
        self.fields = list(fields)
        self.sample_rows = sample_rows[:80]
        for combo, placeholder in ((self.field, '选择标题来源列…'),
                                    (self.insert_field, '选择要插入的 Excel 列…')):
            combo.blockSignals(True)
            combo.clear()
            combo.addItem(placeholder, '')
            for field in self.fields:
                combo.addItem(field, field)
            combo.blockSignals(False)
        if selected and selected not in self.fields:
            self.field.addItem(f'{selected}（当前文件缺少）', selected)
        self.field.setCurrentIndex(max(0, self.field.findData(selected)))
        self._refresh()

    def set_context(self, group_field='', route_field=''):
        self.group_field, self.route_field = group_field, route_field
        self._refresh()

    def configuration(self):
        return {'mode': self.mode.currentData(), 'field': self.field.currentData() or '',
                'prefix': self.prefix.text(), 'suffix': self.suffix.text(),
                'template': self.template.text()}

    def rule(self):
        return compile_title_rule(self.configuration(), self.fields,
                                  self.group_field, self.route_field)

    def _insert_field(self):
        field = self.insert_field.currentData()
        if field:
            self.template.insert(field_token(field))
            self.template.setFocus()

    def _refresh(self, *args):
        mode = self.mode.currentData()
        self.dynamic_controls.setVisible(mode != 'custom')
        self.field.setVisible(mode == 'field')
        self.custom_controls.setVisible(mode == 'custom')
        self.insert_button.setEnabled(bool(self.fields))
        if not self.fields:
            self.preview.setText('导入 Excel 后，可直接选择标题来源列；标题示例会自动显示在这里。')
            return
        try:
            rule = self.rule()
            examples = []
            for row in self.sample_rows:
                values = {field: text_of(row.get(field), field) for field in rule.fields}
                title = rule.render(values)
                if title not in examples:
                    examples.append(title)
                if len(examples) == 3:
                    break
            self.preview.setText('标题示例：' + ' / '.join(value[:100] for value in examples))
            self.preview.setToolTip('\n'.join(examples))
        except ValueError as exc:
            self.preview.setText(str(exc))
            self.preview.setToolTip('')
