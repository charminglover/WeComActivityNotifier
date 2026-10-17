"""Verify arbitrary-column destinations through preparation and transport payloads."""
import unittest
from unittest.mock import patch

from engine import WeComSender
from services.notification_workflow import prepare_selection, validate_routes


class TestGenericRouting(unittest.TestCase):
    def snapshot(self, rows):
        return {'ids': {row['_id'] for row in rows}, 'records': rows,
                'fields': ['Region', 'Department', 'Description'],
                'title': '{group_value}通知', 'group_field': 'Department',
                'route_field': 'Region', 'mode': 'Markdown表格'}

    def test_route_field_independent_of_content_group_after_splitting(self):
        rows = [{'Region': region, 'Department': department,
                 'Description': '详情' * 100, '_id': str(index)}
                for index, (region, department) in enumerate(
                    [('North', 'Sales')] * 25 + [('South', 'Sales')] * 25 + [('North', 'Support')])]
        messages = prepare_selection(self.snapshot(rows))
        mapping = {region: 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=' + region
                   for region in ('North', 'South')}
        validate_routes([m.job for m in messages], mapping, '群机器人Webhook', '', '')
        seen = []
        sender = WeComSender('群机器人Webhook')
        with patch.object(sender, '_request', return_value={'errcode': 0}) as request:
            for message in messages:
                job = message.job
                self.assertEqual(job.route_field, 'Region')
                self.assertTrue(all(row['Region'] == job.route_value for row in job.rows))
                self.assertTrue(all(row['Department'] == job.group_value for row in job.rows))
                seen.extend(row['_id'] for row in job.rows)
                sender.send(message, message.mode, recipient=mapping[job.route_value])
                self.assertEqual(request.call_args.args[1], mapping[job.route_value])
        self.assertCountEqual(seen, [row['_id'] for row in rows])
        self.assertEqual(len(seen), len(set(seen)))
        self.assertTrue(all(len(m.payload.encode('utf-8')) <= 4096 for m in messages))

    def test_missing_value_or_destination_blocks_whole_batch(self):
        for region, mapping in [('South', {'North': 'user-north'}),
                                ('North', {'North': ''}), ('', {'North': 'user-north'})]:
            with self.subTest(region=region, mapping=mapping):
                rows = [{'_id': '1', 'Region': region, 'Department': 'Sales', 'Description': 'data'}]
                messages = prepare_selection(self.snapshot(rows))
                with self.assertRaisesRegex(ValueError, '已阻止整批发送'):
                    validate_routes([m.job for m in messages], mapping, '自建应用消息', '', '')

    def test_missing_source_field_blocks_preparation(self):
        rows = [{'_id': '1', 'Department': 'Sales', 'Description': 'data'}]
        with self.assertRaisesRegex(ValueError, '缺少映射匹配列'):
            prepare_selection(self.snapshot(rows))


if __name__ == '__main__':
    unittest.main()
