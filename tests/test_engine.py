import base64
import hashlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import engine
from engine import (read_excel, text_of, number, rate_number, filter_records,
                    make_jobs, fit_markdown_jobs, render_table, image_bytes,
                    markdown_table, check_webhook, WeComSender)

HERE=Path(__file__).resolve().parents[1]


class TestCore(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.fields, cls.records=read_excel(str(HERE/'演示活动数据.xlsx'))

    def test_excel_import(self):
        self.assertEqual(len(self.records),5)
        self.assertEqual(len(self.fields),35)
        self.assertEqual(self.records[0]['任务台数'],7)
        self.assertEqual(text_of(self.records[0]['任务执行率'],'任务执行率'),'100.00%')

    def test_filter(self):
        values=filter_records(self.records, metric='任务执行率',operator='<',threshold='70',
                              count_metric='任务台数',count_operator='>=',count_threshold='10')
        self.assertEqual(len(values),2)
        self.assertTrue(all(rate_number(v['任务执行率']) < 70 for v in values))

    def test_generic_headers_and_search(self):
        fields, rows = engine._parse_sheets([('Report', iter([
            ['Quarterly report'],
            ['Region', 'Revenue', 'Period'],
            ['North', 1200, '2026-Q1'],
            ['South', 900, '2026-Q1'],
        ]))])
        self.assertEqual(fields, ['Region', 'Revenue', 'Period'])
        self.assertEqual(len(rows), 2)
        self.assertEqual(len(filter_records(rows, shop='north')), 1)
        self.assertEqual(len(filter_records(rows, activity='2026-Q1')), 2)

    def test_jobs_and_paging(self):
        jobs=make_jobs(self.records,['任务服务店名称','活动名称','任务执行率'],group_by_shop=True,page_size=1)
        self.assertEqual(len(jobs),5)
        self.assertEqual(len({(j.store_code,j.store_name) for j in jobs}),3)
        self.assertEqual((jobs[0].part,jobs[0].parts),(1,2))

    def test_markdown(self):
        jobs=make_jobs(self.records,['任务服务店名称','活动名称','任务执行率'])
        result=markdown_table(jobs[0])
        self.assertIn('| 任务服务店名称 |',result)
        self.assertIn('100.00%',result)
        self.assertLessEqual(len(result.encode('utf8')),4096)

    def test_auto_md_split(self):
        data=[dict(self.records[0], **{'活动名称':'很长的标题'*120}) for _ in range(8)]
        jobs=make_jobs(data,['活动名称','任务台数'],page_size=8)
        split=fit_markdown_jobs(jobs)
        self.assertGreater(len(split),1)
        for j in split:self.assertLessEqual(len(markdown_table(j).encode('utf8')),4096)

    def test_image(self):
        j=make_jobs(self.records[:3],['任务服务店名称','活动名称','任务台数','任务车辆执行台数','任务执行率'])[0]
        image=render_table(j)
        b=image_bytes(image)
        self.assertGreater(image.width,800)
        self.assertLess(len(b),2*1024*1024)
        self.assertTrue(b.startswith(b'\xff\xd8'))

    def test_url(self):
        with self.assertRaises(ValueError):check_webhook('https://evil.example/?key=123')
        self.assertTrue(check_webhook('https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=fake'))

    @patch('requests.Session.request')
    def test_webhook_image_payload(self,request):
        response=Mock()
        response.raise_for_status.return_value=None
        response.json.return_value={'errcode':0,'errmsg':'ok'}
        request.return_value=response
        j=make_jobs(self.records[:1],['活动名称','任务执行率'])[0]
        sender=WeComSender('群机器人Webhook',webhook='https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=demo')
        sender.send(j,'图片表格')
        body=request.call_args.kwargs['json']
        image=base64.b64decode(body['image']['base64'])
        self.assertEqual(body['image']['md5'],hashlib.md5(image).hexdigest())

if __name__=='__main__': unittest.main()
