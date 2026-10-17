"""Dynamic titles must partition source data, not label a mixed table with its first row."""
import io
import unittest

from openpyxl import load_workbook
from engine import EXCEL_MODE, make_jobs
from services.notification_workflow import prepare_selection
from services.title_rules import compile_title_rule, field_token


class TestDynamicTitles(unittest.TestCase):
    def test_auto_uses_group_then_route_then_generic_fallback(self):
        values = {'区域': '华东', '门店': 'A店'}
        rule = compile_title_rule({'mode': 'auto'}, values, '区域', '门店')
        self.assertEqual(rule.render(values), '华东通知')
        rule = compile_title_rule({'mode': 'auto'}, values, '', '门店')
        self.assertEqual(rule.render(values), 'A店通知')
        self.assertEqual(compile_title_rule({'mode': 'auto'}, values).render(values), '数据通知')

    def test_title_column_not_in_push_fields_separates_messages_and_attachments(self):
        rows = [{'_id': str(index), '门店': 'A店' if index < 30 else 'B店',
                 '区域': '华东', '详情': f'{index} ' + '说明' * 100} for index in range(40)]
        snapshot = {'ids': {row['_id'] for row in rows}, 'records': rows,
                    'fields': ['详情'], 'group_field': '', 'route_field': '区域',
                    'title': 'legacy', 'title_rule': {'mode': 'field', 'field': '门店',
                    'prefix': '【', 'suffix': '】日报'}, 'mode': 'Markdown表格', 'include_excel': True}
        messages = prepare_selection(snapshot)
        seen = []
        for message in messages:
            job = message.job
            store = job.title_key[0]
            self.assertEqual(job.title, f'【{store}】日报')
            self.assertTrue(all(row['门店'] == store for row in job.rows))
            self.assertEqual(job.route_value, '华东')
            if message.mode == EXCEL_MODE:
                self.assertIn(store, message.filename)
                workbook = load_workbook(io.BytesIO(message.payload), read_only=True)
                self.assertEqual(next(workbook.active.values), ('详情',))
                self.assertEqual(len(list(workbook.active.values)), len(job.rows) + 1)
                workbook.close()
            else:
                seen.extend(row['_id'] for row in job.rows)
        self.assertCountEqual(seen, [row['_id'] for row in rows])
        self.assertEqual(len([m for m in messages if m.mode == EXCEL_MODE]), 2)
        for mode in ('Markdown表格', EXCEL_MODE):
            for title in ('【A店】日报', '【B店】日报'):
                pages = [m.job for m in messages if m.mode == mode and m.job.title == title]
                self.assertEqual([p.part for p in pages], list(range(1, len(pages) + 1)))
                self.assertTrue(all(p.parts == len(pages) for p in pages))

    def test_combination_titles_reference_multiple_columns_and_keep_values_literal(self):
        special = '门店{"别名"}\\值'
        fields = ['区域', special]
        template = field_token('区域') + ' / ' + field_token(special) + '日报'
        rule = compile_title_rule({'mode': 'custom', 'template': template}, fields)
        self.assertEqual(rule.fields, tuple(fields))
        self.assertEqual(rule.render({'区域': '华东', special: '{group_value}店'}),
                         '华东 / {group_value}店日报')
        rows = [{'区域': '华东', special: 'A'}, {'区域': '华东', special: 'B'},
                {'区域': '华南', special: 'A'}]
        jobs = make_jobs(rows, ['区域'], title_rule={'mode': 'custom', 'template': template})
        self.assertEqual([job.title for job in jobs], ['华东 / A日报', '华东 / B日报', '华南 / A日报'])

    def test_missing_header_or_incomplete_token_is_explicit_and_blank_values_are_labeled(self):
        with self.assertRaisesRegex(ValueError, '不在当前文件'):
            compile_title_rule({'mode': 'field', 'field': '不存在'}, ['区域'])
        with self.assertRaisesRegex(ValueError, '列标记格式不完整'):
            compile_title_rule({'mode': 'custom', 'template': '{列:"区域"'}, ['区域'])
        rule = compile_title_rule({'mode': 'field', 'field': '区域'}, ['区域'])
        self.assertEqual(rule.render({'区域': ''}), '区域未填写通知')
        with self.assertRaisesRegex(ValueError, '标题来源列'):
            make_jobs([{'区域': '华东'}, {'内容': 'Missing'}], ['内容'],
                      title_rule={'mode': 'field', 'field': '区域'})

    def test_fixed_custom_title_and_legacy_template_remain_available(self):
        rows = [{'区域': '华东'}, {'区域': '华南'}]
        jobs = make_jobs(rows, ['区域'], title_rule={'mode': 'custom', 'template': '统一公告'})
        self.assertEqual(len(jobs), 1)
        self.assertEqual(jobs[0].title, '统一公告')
        rule = compile_title_rule({'mode': 'custom', 'template': '{group_field}：{group_value}'},
                                  ['区域'], group_field='区域')
        self.assertEqual(rule.render({'区域': '华东'}), '区域：华东')

    def test_different_titles_with_same_windows_filename_get_distinct_attachments(self):
        rows = [{'_id': '1', '门店': 'A/B'}, {'_id': '2', '门店': 'A\\B'}]
        messages = prepare_selection({'ids': {'1', '2'}, 'records': rows,
            'fields': ['门店'], 'title': '', 'title_rule': {'mode': 'field', 'field': '门店'},
            'mode': EXCEL_MODE})
        self.assertEqual(len(messages), 2)
        self.assertNotEqual(messages[0].filename.casefold(), messages[1].filename.casefold())
        self.assertEqual({m.job.title for m in messages}, {'A/B通知', 'A\\B通知'})


if __name__ == '__main__':
    unittest.main()
