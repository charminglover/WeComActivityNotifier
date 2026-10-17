"""No live enterprise credentials are used in these regression tests."""
import base64
import hashlib
import io
import unittest
from pathlib import Path
from unittest.mock import patch, Mock
from PIL import Image
import engine
from engine import Job, prepare_messages, image_bytes, WeComSender

ROOT=Path(__file__).resolve().parents[1]

class TestPacking(unittest.TestCase):
    def setUp(self):
        self.fields, self.rows = engine.read_excel(str(ROOT/'演示活动数据.xlsx'))

    def test_autosplit_markdown_and_no_missing_rows(self):
        rows=[]
        for i in range(70):
            row=dict(self.rows[0]);row['_id']=f'row-{i}'
            row['活动名称']='我的长活动名称'*(10+i%3)+str(i)
            rows.append(row)
        jobs=engine.make_jobs(rows,['活动编码','活动名称','任务台数'], page_size=60)
        prepared=prepare_messages(jobs,'Markdown表格')
        self.assertGreater(len(prepared),1)
        self.assertEqual(sum(len(m.job.rows) for m in prepared),70)
        self.assertTrue(all(len(m.payload.encode('utf8'))<=4096 for m in prepared))
        self.assertEqual([m.job.part for m in prepared],list(range(1,len(prepared)+1)))
        self.assertTrue(all(m.job.parts==len(prepared) for m in prepared))

    def test_image_split_under_forced_tiny_byte_budget(self):
        # Simulate stricter service limit to prove automatic splitting.
        jobs=engine.make_jobs(self.rows,['任务服务店名称','活动名称','任务台数','任务车辆执行台数',
                                         '任务执行率','按期执行率'],page_size=5)
        # Keep the cap below the serialized sample size across Windows and macOS fonts.
        with patch.object(engine,'MAX_IMAGE_BYTES',12_000):
            prepared=prepare_messages(jobs,'图片表格')
        self.assertGreaterEqual(len(prepared),2)
        self.assertEqual(sum(len(m.job.rows) for m in prepared if '活动名称' in m.job.fields),5)
        for m in prepared:
            self.assertTrue(m.payload.startswith(b'\xff\xd8'))
            self.assertLessEqual(len(m.payload),22_000)
            with Image.open(io.BytesIO(m.payload)) as picture:
                self.assertEqual(picture.format,'JPEG')
                self.assertGreater(picture.width,100)

    def test_image_byte_size_and_md5(self):
        jobs=engine.make_jobs(self.rows[:2],['活动名称','任务台数'])
        msgs=prepare_messages(jobs,'图片表格')
        data=msgs[0].payload
        self.assertLessEqual(len(data),engine.MAX_IMAGE_BYTES)
        with patch('requests.Session.request') as request:
            response=Mock()
            response.raise_for_status.return_value=None
            response.json.return_value={'errcode':0, 'errmsg':'ok'}
            request.return_value=response
            sender=WeComSender('群机器人Webhook',webhook='https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=sample')
            sender.send(msgs[0], '图片表格')
            body=request.call_args.kwargs['json']
            self.assertEqual(body['msgtype'],'image')
            self.assertEqual(base64.b64decode(body['image']['base64']),data)
            self.assertEqual(body['image']['md5'],hashlib.md5(data).hexdigest())

    def test_application_image_multipart_and_media_send(self):
        jobs=engine.make_jobs(self.rows[:1],['活动名称','任务台数'])
        message=prepare_messages(jobs,'图片表格')[0]
        sender=WeComSender('自建应用消息', corpid='corpid',secret='app_secret',agentid='100',touser='user')
        with patch.object(sender,'_request',side_effect=[{'access_token':'mock-token'},
                                                         {'media_id':'mock-media-id'},
                                                         {'errcode':0,'errmsg':'ok'}]) as req:
            sender.send(message,'图片表格')
            upload=req.call_args_list[1]
            headers=upload.kwargs['headers']
            data=upload.kwargs['data']
            self.assertIn('multipart/form-data; boundary=',headers['Content-Type'])
            self.assertIn(f'filelength={len(message.payload)}'.encode(), data)
            self.assertIn(message.payload,data)
            send_body=req.call_args_list[2].kwargs['json']
            self.assertEqual(send_body['image']['media_id'],'mock-media-id')

    def test_failed_extremely_long_value_no_partial_batch(self):
        row=dict(self.rows[0], **{'活动名称': '某' * 120_000})
        jobs=[Job('标题',['活动名称'],[row])]
        with self.assertRaisesRegex(ValueError,'已取消本批发送'):
            prepare_messages(jobs,'图片表格')

    def test_per_shop_page_numbering(self):
        jobs=engine.make_jobs(self.rows,['活动名称','任务执行率'],group_by_shop=True,page_size=1)
        msgs=prepare_messages(jobs,'Markdown表格')
        for key in set((m.job.store_code,m.job.store_name) for m in msgs):
            group=[m for m in msgs if (m.job.store_code,m.job.store_name)==key]
            self.assertEqual([x.job.part for x in group],list(range(1,len(group)+1)))

if __name__ == '__main__':
    unittest.main()
