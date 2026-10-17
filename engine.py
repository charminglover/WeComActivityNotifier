"""Excel data, report generation and WeCom official HTTP transports.
No Tkinter usage here: safe to call from background workers.
"""
from __future__ import annotations

import base64
import csv
import hashlib
import io
import math
import re
import sys
import os
import time
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit

import requests
from PIL import Image, ImageDraw, ImageFont

DEFAULT_FIELDS = [
    '任务服务店名称', '活动名称', '活动编码', '活动状态',
    '任务台数', '任务车辆执行台数', '任务执行率',
    '按期完成执行数量', '按期执行率', '售前车辆数', '售前执行数量', '售前执行率',
]
KEY_STORE = '任务服务店编码'
NAME_STORE = '任务服务店名称'
ACTIVITY = '活动名称'
PERCENT_NAMES = ('执行率', '达标率', '完成率')
MAX_ROBOT_BYTES = 4096
MAX_IMAGE_BYTES = 2 * 1024 * 1024 - 1024
USER_AGENT = 'WeComActivityNotifier/1.0'


def text_of(value: Any, field: str = '') -> str:
    if value is None:
        return ''
    if isinstance(value, (datetime, date)):
        return value.strftime('%Y-%m-%d %H:%M:%S' if isinstance(value, datetime) else '%Y-%m-%d')
    if is_rate_field(field):
        n = rate_number(value)
        return f'{n:.2f}%' if n is not None else str(value)
    if isinstance(value, float):
        if math.isnan(value):
            return ''
        return str(int(value)) if value.is_integer() else f'{value:.2f}'.rstrip('0').rstrip('.')
    return str(value).strip()


def is_rate_field(field: str) -> bool:
    return any(s in field for s in PERCENT_NAMES)


def number(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (float, int)):
        return float(value) if math.isfinite(float(value)) else None
    s = str(value).replace(',', '').replace('，', '').strip()
    if not s:
        return None
    if s.endswith('%'):
        s = s[:-1].strip()
    try:
        n = float(s)
        return n if math.isfinite(n) else None
    except ValueError:
        return None


def rate_number(value: Any) -> float | None:
    n = number(value)
    if n is None:
        return None
    # Excel percentage cells usually store 75% as 0.75.
    # Strings '75%' and plain numbers 75 already represent percent values.
    if isinstance(value, (int, float)) and 0 <= n <= 1:
        return n * 100
    return n


def read_excel(path: str) -> tuple[list[str], list[dict]]:
    """Read ALL worksheets, use header row within first 15 rows, preserve source row IDs.
    This function is designed for running off the UI thread.
    """
    suffix = Path(path).suffix.lower()
    sheets = []
    if suffix in ('.xlsx', '.xlsm'):
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            for ws in wb.worksheets:
                if ws.sheet_state != 'visible':
                    continue
                sheets.append((ws.title, ws.iter_rows(values_only=True)))
            return _parse_sheets(sheets)
        finally:
            wb.close()
    if suffix == '.xls':
        import xlrd
        wb = xlrd.open_workbook(path, on_demand=True)
        try:
            for ws in wb.sheets():
                def rows_for(sheet=ws):
                    for i in range(sheet.nrows):
                        yield sheet.row_values(i)
                sheets.append((ws.name, rows_for()))
            return _parse_sheets(sheets)
        finally:
            wb.release_resources()
    if suffix == '.csv':
        for encoding in ('utf-8-sig', 'gb18030'):
            try:
                with open(path, 'r', encoding=encoding, newline='') as f:
                    rows = list(csv.reader(f))
                return _parse_sheets([('CSV', rows)])
            except UnicodeDecodeError:
                continue
        raise ValueError('CSV 编码无法识别（支持 UTF-8 / GB18030）')
    raise ValueError('支持 .xlsx、.xlsm、.xls 和 .csv')


def _parse_sheets(sheets) -> tuple[list[str], list[dict]]:
    all_fields: list[str] = []
    all_records: list[dict] = []
    for sheet_name, iterator in sheets:
        it = iter(iterator)
        ahead = []
        for _ in range(15):
            try:
                ahead.append(next(it))
            except StopIteration:
                break
        if not ahead:
            continue
        marker = ('任务服务店名称', '任务台数', '活动名称', '任务执行率')
        head_idx = max(range(len(ahead)), key=lambda i: sum(str(c or '').strip() in marker for c in ahead[i]))
        header = [str(x).strip() if x is not None else '' for x in ahead[head_idx]]
        if sum(k in header for k in marker) < 2:
            continue
        used = set()
        for i, x in enumerate(header):
            if not x:
                header[i] = f'未命名列{i+1}'
            if header[i] in used:
                header[i] = f'{header[i]}（列{i+1}）'
            used.add(header[i])
        for h in header:
            if h not in all_fields:
                all_fields.append(h)
        def make_rows():
            yield from ahead[head_idx + 1:]
            yield from it
        for row_number, values in enumerate(make_rows(), head_idx + 2):
            if not any(v is not None and str(v).strip() for v in values):
                continue
            row = {field: values[i] if i < len(values) else None for i, field in enumerate(header)}
            row['_id'] = f'{sheet_name}:{row_number}'
            row['_sheet'] = sheet_name
            row['_row'] = row_number
            all_records.append(row)
    if not all_records:
        raise ValueError('没有找到有效数据。请检查首行附近是否包含「任务服务店名称」「活动名称」「任务台数」等列。')
    return all_fields, all_records


def filter_records(records: list[dict], shop: str = '', activity: str = '', status: str = '',
                   metric: str = '', operator: str = '<', threshold: str = '',
                   count_metric: str = '', count_operator: str = '>', count_threshold: str = '') -> list[dict]:
    shop, activity = shop.strip().lower(), activity.strip().lower()
    if threshold.strip():
        target = number(threshold)
        if target is None:
            raise ValueError('完成率筛选值必须是数字，例如 80')
    else:
        target = None
    if count_threshold.strip():
        count_target = number(count_threshold)
        if count_target is None:
            raise ValueError('数量筛选值必须是数字')
    else:
        count_target = None
    compare = {'<': lambda a,b: a<b, '<=': lambda a,b: a<=b, '=':lambda a,b: a==b,
               '>=':lambda a,b: a>=b, '>':lambda a,b: a>b}
    out = []
    for r in records:
        if shop and shop not in (' '.join(str(r.get(k, '')) for k in ('任务服务店名称','任务服务店编码','任务服务店区域','任务服务店大区','任务服务店服务中心'))).lower():
            continue
        if activity and activity not in (' '.join(str(r.get(k, '')) for k in ('活动名称', '活动编码'))).lower():
            continue
        if status and status != '全部' and str(r.get('活动状态') or '').strip() != status:
            continue
        if target is not None:
            val = rate_number(r.get(metric)) if is_rate_field(metric) else number(r.get(metric))
            if val is None or not compare[operator](val, target):
                continue
        if count_target is not None:
            val = number(r.get(count_metric))
            if val is None or not compare[count_operator](val, count_target):
                continue
        out.append(r)
    return out


@dataclass
class Job:
    title: str
    fields: list[str]
    rows: list[dict]
    store_code: str = ''
    store_name: str = ''
    part: int = 1
    parts: int = 1


def make_jobs(records: list[dict], fields: list[str], group_by_shop: bool = False,
              page_size: int = 8, columns_per_page: int = 7, title: str = '售后活动执行进度通知') -> list[Job]:
    if not records or not fields:
        raise ValueError('需要至少选择一条活动记录和一项推送字段')
    page_size = max(1, min(30, int(page_size)))
    columns_per_page = max(2, min(10, int(columns_per_page)))
    groups = defaultdict(list)
    if group_by_shop:
        for row in records:
            code = text_of(row.get(KEY_STORE))
            # Use name to avoid incorrectly mixing stores if code is empty.
            name = text_of(row.get(NAME_STORE))
            groups[(code, name)].append(row)
    else:
        groups[('', '')] = records
    jobs = []
    for (code, name), entries in groups.items():
        row_parts = [entries[i:i+page_size] for i in range(0, len(entries), page_size)]
        field_parts = [fields[i:i+columns_per_page] for i in range(0, len(fields), columns_per_page)]
        part_total = len(row_parts) * len(field_parts)
        for rowpart in row_parts:
            for fieldpart in field_parts:
                jobs.append(Job(title=title, fields=fieldpart, rows=rowpart, store_code=code,
                                store_name=name, part=len(jobs) + 1 if not group_by_shop else 1, parts=part_total))
        # Correct 1-based numbering within each group.
        for idx, job in enumerate(jobs[-part_total:], 1):
            job.part = idx
    return jobs


def _get_font(size: int, bold: bool = False):
    if os.name == 'nt':
        folder = Path(os.environ.get('WINDIR', 'C:/Windows')) / 'Fonts'
        names = ['msyhbd.ttc', 'simhei.ttf', 'msyh.ttc'] if bold else ['msyh.ttc', 'simhei.ttf', 'simsun.ttc']
        paths = [folder / n for n in names]
    else:
        paths = [Path('/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'),
                 Path('/usr/share/fonts/truetype/wqy/wqy-microhei.ttc'),
                 Path('/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc')]
    for p in paths:
        if p.exists():
            try:
                return ImageFont.truetype(str(p), size)
            except OSError:
                pass
    return ImageFont.load_default(size=size)


def _wrap(draw, text: str, font, width: int, max_lines: int = 3) -> list[str]:
    text = str(text or '').replace('\n', ' ').replace('\r', '')
    pieces = []
    line = ''
    for pos, char in enumerate(text):
        if draw.textlength(line + char, font=font) > width and line:
            if len(pieces) >= max_lines - 1:
                return pieces + [_trim(draw, line + text[pos:], font, width)]
            pieces.append(line)
            line = char
        else:
            line += char
    if line:
        pieces.append(line)
    return pieces[:max_lines] or ['']


def _trim(draw, value: str, font, max_width: int) -> str:
    if draw.textlength(value, font=font) <= max_width:
        return value
    out = ''
    for c in value:
        if draw.textlength(out + c + '…', font=font) > max_width:
            break
        out += c
    return out + '…'


def _col_width(field: str) -> int:
    if field == '活动名称':
        return 360
    if '名称' in field or '备注' in field:
        return 270
    if '时间' in field:
        return 235
    if '执行台数' in field or '执行数量' in field:
        return 225
    if any(k in field for k in ('编码', '经理','督导')):
        return 200
    return 155 if any(k in field for k in ('率', '台数', '数量')) else 185


def render_table(job: Job, warning_rate: float = 80) -> Image.Image:
    columns = job.fields
    widths = [66] + [_col_width(c) for c in columns]
    full_width = sum(widths) + 72
    face = _get_font(20)
    headfont = _get_font(20, True)
    titlefont = _get_font(30, True)
    small = _get_font(18)
    canvas = Image.new('RGB', (full_width, 4000), 'white')
    d = ImageDraw.Draw(canvas)
    d.text((36, 28), job.title, fill='#202020', font=titlefont)
    subtitle = f'{job.store_name}  ' if job.store_name else ''
    subtitle += f'共 {len(job.rows)} 条活动  ·  第 {job.part}/{job.parts} 页'
    d.text((36, 82), _trim(d, subtitle, small, full_width-72), fill='#666666', font=small)
    y = 122
    x = 36
    for width, field in zip(widths, ['序号'] + columns):
        d.rectangle((x,y,x+width,y+52), fill='#E8ECF1', outline='#CCD3DD', width=1)
        d.text((x+10,y+14), _trim(d,field,headfont,width-20), fill='#222222',font=headfont)
        x += width
    y += 52
    for idx, r in enumerate(job.rows, 1):
        line_list = [['%s' % idx]]
        for field in columns:
            line_list.append(_wrap(d, text_of(r.get(field), field), face, _col_width(field)-20, 3))
        row_height = max(48, max(len(v) for v in line_list)*28+18)
        x = 36
        for ci, (width, lines) in enumerate(zip(widths, line_list)):
            shade = '#FFFFFF' if idx % 2 else '#F8F9FB'
            if ci > 0 and is_rate_field(columns[ci-1]):
                rate = rate_number(r.get(columns[ci-1]))
                if rate is not None and rate < warning_rate:
                    shade = '#FCECEA'
            d.rectangle((x,y,x+width,y+row_height), fill=shade, outline='#D9DEE6', width=1)
            for n, line in enumerate(lines):
                d.text((x+10,y+12+n*28),line,fill='#202020',font=face)
            x += width
        y += row_height
    foot = '说明：低于设定阈值的完成率以浅红色标识；每行对应 Excel 的一条活动记录。'
    d.text((36,y+20),_trim(d,foot,small,full_width-72),fill='#707070',font=small)
    return canvas.crop((0,0,full_width,y+58))


def image_bytes(image: Image.Image) -> bytes:
    for quality in (89, 80, 69, 55):
        buf = io.BytesIO()
        image.convert('RGB').save(buf, format='JPEG', quality=quality, optimize=True, progressive=True)
        value = buf.getvalue()
        if len(value) < MAX_IMAGE_BYTES:
            return value
    raise ValueError('图片超过企业微信 2MB 限制，请减小每张图片的行数或列数')


def markdown_table(job: Job) -> str:
    def esc(s):
        return str(s).replace('|','\\|').replace('\n', ' ').replace('\r',' ').replace('`', '\\`')
    lines = [f'## {esc(job.title)}', f'> {esc(job.store_name) + " · " if job.store_name else ""}共 {len(job.rows)} 条 · 第 {job.part}/{job.parts} 页',
             '', '| ' + ' | '.join(esc(f) for f in job.fields) + ' |',
             '| ' + ' | '.join('---' for _ in job.fields) + ' |']
    for r in job.rows:
        lines.append('| ' + ' | '.join(esc(text_of(r.get(f), f)) for f in job.fields) + ' |')
    result = '\n'.join(lines)
    if len(result.encode('utf8')) > MAX_ROBOT_BYTES:
        raise ValueError('Markdown 消息超出 4096 字节，请减小每页行数或列数，或使用图片模式')
    return result



def fit_markdown_jobs(jobs: list[Job]) -> list[Job]:
    """Automatically split overlong markdown_v2 messages by rows/columns."""
    def fit(job):
        try:
            markdown_table(job)
            return [job]
        except ValueError:
            if len(job.rows) > 1:
                mid = len(job.rows) // 2
                left = Job(job.title,job.fields,job.rows[:mid],job.store_code,job.store_name)
                right = Job(job.title,job.fields,job.rows[mid:],job.store_code,job.store_name)
                return fit(left) + fit(right)
            if len(job.fields) > 1:
                mid = len(job.fields) // 2
                left = Job(job.title,job.fields[:mid],job.rows,job.store_code,job.store_name)
                right = Job(job.title,job.fields[mid:],job.rows,job.store_code,job.store_name)
                return fit(left) + fit(right)
            raise ValueError('单条活动单个字段已超过企业微信 4096 字节限制，无法继续自动分页')
    expanded = []
    for job in jobs:
        expanded.extend(fit(job))
    segments = defaultdict(list)
    for job in expanded:
        segments[(job.store_code,job.store_name)].append(job)
    for group in segments.values():
        for idx,job in enumerate(group,1):
            job.part = idx
            job.parts = len(group)
    return expanded


def readable_jobs(jobs: list[Job], mode: str = '图片表格') -> tuple[int, list[str]]:
    warnings = []
    if mode == 'Markdown表格':
        for j in jobs:
            if len(markdown_table(j).encode('utf8')) > MAX_ROBOT_BYTES:
                warnings.append('消息超过 Markdown 限制')
    return len(jobs), warnings


def read_mapping(path: str, kind: str) -> dict[str,str]:
    if not path:
        return {}
    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        data = list(csv.DictReader(f))
    if not data:
        raise ValueError('映射 CSV 为空')
    dest_column = 'Webhook' if kind == '群机器人Webhook' else 'UserID'
    if not all(x in data[0] for x in (KEY_STORE,dest_column)):
        raise ValueError(f'映射表必须包含「{KEY_STORE}」和「{dest_column}」列')
    mapping = {}
    for row in data:
        code = str(row.get(KEY_STORE) or '').strip()
        value = str(row.get(dest_column) or '').strip()
        if code and value:
            if code in mapping and mapping[code] != value:
                raise ValueError(f'服务店 {code} 在映射表中存在冲突')
            mapping[code] = value
    return mapping


def check_webhook(url: str) -> str:
    url = url.strip()
    parsed = urlsplit(url)
    if parsed.scheme != 'https' or parsed.hostname != 'qyapi.weixin.qq.com' or parsed.path != '/cgi-bin/webhook/send' or 'key=' not in (parsed.query or ''):
        raise ValueError('Webhook 必须为企业微信官方 https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=...')
    return url


def _ensure_ok(response):
    response.raise_for_status()
    try:
        data = response.json()
    except ValueError:
        raise RuntimeError('企业微信返回非 JSON 响应')
    if int(data.get('errcode', -1)) != 0:
        raise RuntimeError(f'企业微信错误码 {data.get("errcode", "未知")}：{data.get("errmsg", "未知错误")}')
    return data


def _safe_error(e: Exception, secrets: list[str]) -> str:
    result = str(e)
    for secret in secrets:
        if secret:
            result = result.replace(secret, '[已隐藏]')
    result = re.sub(r'key=[^&\s]+', 'key=[已隐藏]', result)
    result = re.sub(r'access_token=[^&\s]+', 'access_token=[已隐藏]', result)
    return result


class WeComSender:
    def __init__(self, kind: str, *, webhook: str = '', corpid: str = '', secret: str = '',
                 agentid: str = '', touser: str = '', timeout: int = 15):
        self.kind = kind
        self.webhook = webhook.strip()
        self.corpid = corpid.strip()
        self.secret = secret.strip()
        self.agentid = agentid.strip()
        self.touser = touser.strip()
        self.timeout = timeout
        self.session = requests.Session()
        self._token = None
        if kind == '群机器人Webhook':
            if self.webhook:
                check_webhook(self.webhook)
        else:
            if not all((self.corpid, self.secret, self.agentid)) or not self.agentid.isdigit():
                raise ValueError('应用模式需要填写 CorpID、Secret、数字 AgentID')

    def _request(self, method, url, **kwargs):
        try:
            return _ensure_ok(self.session.request(method, url, timeout=self.timeout, **kwargs))
        except requests.RequestException as e:
            raise RuntimeError(_safe_error(e, [self.secret, self.webhook, self._token or ''])) from None

    def token(self):
        if not self._token:
            data = self._request('GET', 'https://qyapi.weixin.qq.com/cgi-bin/gettoken',
                                 params={'corpid':self.corpid, 'corpsecret':self.secret})
            self._token = data.get('access_token')
            if not self._token:
                raise RuntimeError('无法获取企业微信 access_token')
        return self._token

    def send(self, job: Job, mode: str, recipient: str = '', warning_rate: float = 80):
        if self.kind == '群机器人Webhook':
            url = check_webhook(recipient or self.webhook)
            if mode == '图片表格':
                b = image_bytes(render_table(job, warning_rate))
                body = {'msgtype':'image', 'image':{'base64':base64.b64encode(b).decode('ascii'),
                                                  'md5':hashlib.md5(b).hexdigest()}}
            elif mode == 'Markdown表格':
                body = {'msgtype':'markdown_v2', 'markdown_v2':{'content':markdown_table(job)}}
            else:
                raise ValueError('无效消息格式')
            return self._request('POST', url, json=body)
        user = recipient or self.touser
        if not user:
            raise ValueError('请配置 UserID 或导入服务店映射表')
        if mode == '图片表格':
            b = image_bytes(render_table(job, warning_rate))
            upload = self._request('POST', 'https://qyapi.weixin.qq.com/cgi-bin/media/upload',
                                   params={'access_token':self.token(), 'type':'image'},
                                   files={'media': ('activity_report.jpg', io.BytesIO(b), 'image/jpeg')})
            body = {'touser':user, 'msgtype':'image', 'agentid':int(self.agentid),
                    'image':{'media_id':upload['media_id']}}
        else:
            # App push does not document markdown_v2; use plain text so columns never render as raw Markdown.
            raw = [job.title + (f'（{job.store_name}）' if job.store_name else '')]
            for r in job.rows:
                raw.append('；'.join(f'{f}：{text_of(r.get(f),f)}' for f in job.fields))
            content = '\n'.join(raw)
            if len(content.encode('utf8')) > 2048:
                raise ValueError('应用文本消息超过 2048 字节，请缩小每条消息的行数或选择图片表格')
            body = {'touser':user,'msgtype':'text', 'agentid':int(self.agentid), 'text':{'content':content}}
        result = self._request('POST','https://qyapi.weixin.qq.com/cgi-bin/message/send',
                               params={'access_token':self.token()}, json=body)
        if result.get('invaliduser') or result.get('unlicenseduser'):
            raise RuntimeError(f'部分接收用户无效或不可达：{result.get("invaliduser", "")} {result.get("unlicenseduser", "")}')
        return result
