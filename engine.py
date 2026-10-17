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
import uuid
from collections import defaultdict
from dataclasses import dataclass
from functools import lru_cache
from datetime import date, datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any
from urllib.parse import urlsplit

if TYPE_CHECKING:
    from PIL import Image

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
MAX_IMAGE_BYTES = 1_950_000  # leave margin below the documented 2 MB (decimal)
MAX_IMAGE_WIDTH = 3000
MAX_IMAGE_HEIGHT = 6500
MAX_IMAGE_PIXELS = 16_000_000
MAX_MARKDOWN_SAFE = 3900
USER_AGENT = 'WeComActivityNotifier/2.0'


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


def read_excel(path: str, cancel_event=None) -> tuple[list[str], list[dict]]:
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
            return _parse_sheets(sheets, cancel_event)
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
            return _parse_sheets(sheets, cancel_event)
        finally:
            wb.release_resources()
    if suffix == '.csv':
        for encoding in ('utf-8-sig', 'gb18030'):
            try:
                with open(path, 'r', encoding=encoding, newline='') as f:
                    return _parse_sheets([('CSV', csv.reader(f))], cancel_event)
            except UnicodeDecodeError:
                continue
        raise ValueError('CSV 编码无法识别（支持 UTF-8 / GB18030）')
    raise ValueError('支持 .xlsx、.xlsm、.xls 和 .csv')


def _parse_sheets(sheets, cancel_event=None) -> tuple[list[str], list[dict]]:
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
        marker_scores = [sum(str(c or '').strip() in marker for c in row) for row in ahead]
        best_marker = max(marker_scores, default=0)
        if best_marker >= 2:
            head_idx = max(range(len(ahead)), key=lambda i: (marker_scores[i], sum(bool(c is not None and str(c).strip()) for c in ahead[i]), -i))
        else:
            # Generic worksheets have no WeCom-specific column names. Prefer the
            # densest early row as the header; title rows and blank rows score low.
            head_idx = max(
                range(len(ahead)),
                key=lambda i: (
                    len({str(c).strip() for c in ahead[i] if c is not None and str(c).strip()}),
                    sum(bool(c is not None and str(c).strip()) for c in ahead[i]),
                    -i,
                ),
            )
        header = [str(x).strip() if x is not None else '' for x in ahead[head_idx]]
        if sum(bool(x) for x in header) < 2:
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
            if cancel_event is not None and row_number % 512 == 0 and cancel_event.is_set():
                raise InterruptedError("导入已取消")
            if not any(v is not None and str(v).strip() for v in values):
                continue
            row = {field: values[i] if i < len(values) else None for i, field in enumerate(header)}
            row['_id'] = f'{sheet_name}:{row_number}'
            row['_sheet'] = sheet_name
            row['_row'] = row_number
            all_records.append(row)
    if not all_records:
        raise ValueError('没有找到有效表格。请确认文件前 15 行内包含至少两列表头和数据行。')
    return all_fields, all_records


def filter_records(records: list[dict], shop: str = '', activity: str = '', status: str = '',
                   metric: str = '', operator: str = '<', threshold: str = '',
                   count_metric: str = '', count_operator: str = '>', count_threshold: str = '',
                   conditions: list[dict] | None = None, cancel_event=None) -> list[dict]:
    """Filter records using generic field conditions or the legacy arguments.

    Generic conditions use ``field``, ``operator`` and ``value`` keys. Supported
    operators are text comparisons (contains/equality) and numeric comparisons.
    Legacy parameters remain for callers from older versions.
    """
    if conditions is not None:
        return _filter_by_conditions(records, conditions, cancel_event)
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
    for row_index, r in enumerate(records):
        if cancel_event is not None and row_index % 2048 == 0 and cancel_event.is_set():
            raise InterruptedError("筛选已取消")
        values = ' '.join(str(value) for key, value in r.items()
                          if not key.startswith('_') and value is not None).lower()
        if shop and shop not in values:
            continue
        if activity and activity not in values:
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


def _filter_by_conditions(records: list[dict], conditions: list[dict], cancel_event=None) -> list[dict]:
    compare = {'<': lambda a, b: a < b, '<=': lambda a, b: a <= b,
               '=': lambda a, b: a == b, '>=': lambda a, b: a >= b,
               '>': lambda a, b: a > b, '!=': lambda a, b: a != b}
    text_ops = {'包含', '不包含', '等于', '不等于'}
    compiled = []
    for condition in conditions:
        field = str(condition.get('field') or '').strip()
        op = str(condition.get('operator') or '').strip()
        expected = condition.get('value', '')
        if not field or not op:
            continue
        expected_text = str(expected if expected is not None else '').strip()
        if op in compare:
            target = rate_number(expected_text) if is_rate_field(field) else number(expected_text)
            if target is None:
                raise ValueError(f'“{field}”的筛选值必须是数字')
            compiled.append((field, op, expected_text, expected_text.casefold(), target))
        elif op in text_ops or op in ('为空', '不为空'):
            compiled.append((field, op, expected_text, expected_text.casefold(), None))
        else:
            raise ValueError(f'不支持的筛选方式：{op}')
    output = []
    for row_index, row in enumerate(records):
        if cancel_event is not None and row_index % 2048 == 0 and cancel_event.is_set():
            raise InterruptedError("筛选已取消")
        matches = True
        for field, op, expected_text, expected_folded, target in compiled:
            raw = row.get(field)
            actual_text = text_of(raw, field).strip()
            if op == '为空':
                matches = not actual_text
            elif op == '不为空':
                matches = bool(actual_text)
            elif op in text_ops:
                a, b = actual_text.casefold(), expected_folded
                matches = (b in a if op == '包含' else
                           b not in a if op == '不包含' else
                           a == b if op == '等于' else a != b)
            elif op in compare:
                actual = rate_number(raw) if is_rate_field(field) else number(raw)
                matches = actual is not None and compare[op](actual, target)
            if not matches:
                break
        if matches:
            output.append(row)
    return output


@dataclass
class Job:
    title: str
    fields: list[str]
    rows: list[dict]
    store_code: str = ''
    store_name: str = ''
    part: int = 1
    parts: int = 1
    group_field: str = ''
    group_value: str = ''


def make_jobs(records: list[dict], fields: list[str], group_by_shop: bool = False,
              page_size: int = 8, columns_per_page: int = 7, title: str = '数据通知',
              group_by_field: str = '', route_by_store: bool = False) -> list[Job]:
    """Build message pages, optionally grouped by any imported column.

    ``group_by_shop`` is retained for callers of earlier versions. Store routing
    can be kept separate from the displayed grouping when per-store mappings exist.
    """
    if not records or not fields:
        raise ValueError('需要至少选择一条活动记录和一项推送字段')
    if group_by_shop and not group_by_field:
        group_by_field = KEY_STORE
        route_by_store = True
    page_size = max(1, min(30, int(page_size)))
    columns_per_page = max(2, min(10, int(columns_per_page)))
    groups = defaultdict(list)
    for row in records:
        code = text_of(row.get(KEY_STORE)) if route_by_store else ''
        name = text_of(row.get(NAME_STORE)) if route_by_store else ''
        group_value = text_of(row.get(group_by_field), group_by_field) if group_by_field else ''
        groups[(code, name, group_value)].append(row)
    jobs = []
    for (code, name, group_value), entries in groups.items():
        row_parts = [entries[i:i+page_size] for i in range(0, len(entries), page_size)]
        field_parts = [fields[i:i+columns_per_page] for i in range(0, len(fields), columns_per_page)]
        part_total = len(row_parts) * len(field_parts)
        group_title = title.replace('{group_field}', group_by_field).replace('{group_value}', group_value)
        start = len(jobs)
        for rowpart in row_parts:
            for fieldpart in field_parts:
                jobs.append(Job(title=group_title, fields=fieldpart, rows=rowpart, store_code=code,
                                store_name=name, parts=part_total, group_field=group_by_field,
                                group_value=group_value))
        for idx, job in enumerate(jobs[start:], 1):
            job.part = idx
    return jobs


@lru_cache(maxsize=16)
def _get_font(size: int, bold: bool = False):
    from PIL import ImageFont

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


def _wrap(draw, text: str, font, width: int, max_lines: int | None = None) -> list[str]:
    """Binary-search per-line wrapping avoids quadratic text measurement on long cells.
    Never silently omit Excel text; oversized rows are rejected by the page planner.
    """
    value = str(text or '').replace('\n', ' ').replace('\r', ' ')
    if not value:
        return ['']
    if max_lines is None and draw.textlength(value, font=font) > width * 170:
        raise ValueError('单元格文本过长，无法在一张图片完整呈现；请取消超长备注等字段')
    result = []
    pos = 0
    while pos < len(value):
        low, high, best = pos + 1, len(value), pos + 1
        while low <= high:
            mid = (low + high) // 2
            if draw.textlength(value[pos:mid], font=font) <= width:
                best = mid
                low = mid + 1
            else:
                high = mid - 1
        result.append(value[pos:best])
        pos = best
        if len(result) > 170:
            raise ValueError('单元格文本过长，无法在一张图片完整呈现；请取消超长备注等字段')
        if max_lines is not None and len(result) >= max_lines and pos < len(value):
            # Only used for non-data decoration, never for tabular data.
            result[-1] = _trim(draw, result[-1] + value[pos:], font, width)
            break
    return result


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


class ImagePageOverflow(ValueError):
    """Page needs to be split before allocating a large bitmap."""


def condition_matches(row: dict, condition: dict) -> bool:
    field = str(condition.get('field') or '')
    op = str(condition.get('operator') or '')
    expected = str(condition.get('value') or '').strip()
    raw = row.get(field)
    actual_text = text_of(raw, field).strip()
    if op == '为空':
        return not actual_text
    if op == '不为空':
        return bool(actual_text)
    if op in ('包含', '不包含', '等于', '不等于'):
        actual, target = actual_text.casefold(), expected.casefold()
        return (target in actual if op == '包含' else target not in actual if op == '不包含'
                else actual == target if op == '等于' else actual != target)
    if op in ('<', '<=', '=', '!=', '>=', '>'):
        left = rate_number(raw) if is_rate_field(field) else number(raw)
        right = rate_number(expected) if is_rate_field(field) else number(expected)
        if left is None or right is None:
            return False
        return {'<': left < right, '<=': left <= right, '=': left == right,
                '!=': left != right, '>=': left >= right, '>': left > right}[op]
    return False


def matches_conditions(row: dict, conditions: list[dict] | None) -> bool:
    return bool(conditions) and all(condition_matches(row, condition) for condition in conditions)


def render_table(job: Job, warning_rate: float | None = None,
                 highlight_conditions: list[dict] | None = None) -> Image.Image:
    from PIL import Image, ImageDraw

    columns = job.fields
    widths = [66] + [_col_width(c) for c in columns]
    full_width = sum(widths) + 72
    if full_width > MAX_IMAGE_WIDTH:
        raise ImagePageOverflow('图片列宽过大')
    face = _get_font(20)
    headfont = _get_font(20, True)
    titlefont = _get_font(30, True)
    small = _get_font(18)
    # Create only a tiny measurement surface; allocate the report at its actual size.
    d = ImageDraw.Draw(Image.new('RGB', (1, 1)))
    wrapped = []
    total_height = 122 + 52 + 58
    for row in job.rows:
        lines = [['']]
        for f in columns:
            lines.append(_wrap(d, text_of(row.get(f), f), face, _col_width(f) - 20))
        h = max(48, max(len(v) for v in lines) * 28 + 18)
        total_height += h
        wrapped.append((lines, h))
    if total_height > MAX_IMAGE_HEIGHT or full_width * total_height > MAX_IMAGE_PIXELS:
        raise ImagePageOverflow('图片尺寸过大，需要自动分页')
    canvas = Image.new('RGB', (full_width, total_height), 'white')
    d = ImageDraw.Draw(canvas)
    d.text((36, 28), job.title, fill='#202020', font=titlefont)
    if job.group_field:
        subtitle = f'按{job.group_field}：{job.group_value}  ·  '
    else:
        subtitle = f'{job.store_name}  ' if job.store_name else ''
    subtitle += f'共 {len(job.rows)} 条活动  ·  第 {job.part}/{job.parts} 页'
    d.text((36, 82), _trim(d, subtitle, small, full_width - 72), fill='#666666', font=small)
    y = 122
    x = 36
    for width, field in zip(widths, ['序号'] + columns):
        d.rectangle((x, y, x + width, y + 52), fill='#E8ECF1', outline='#CCD3DD', width=1)
        d.text((x + 10, y + 14), _trim(d, field, headfont, width - 20), fill='#222222', font=headfont)
        x += width
    y += 52
    for idx, (row, (lines_by_col, row_height)) in enumerate(zip(job.rows, wrapped), 1):
        lines_by_col[0] = [str(idx)]
        x = 36
        for ci, (width, lines) in enumerate(zip(widths, lines_by_col)):
            shade = '#FFFFFF' if idx % 2 else '#F8F9FB'
            if matches_conditions(row, highlight_conditions):
                shade = '#FFF0CC'
            elif warning_rate is not None and ci > 0 and is_rate_field(columns[ci - 1]):
                rate = rate_number(row.get(columns[ci - 1]))
                if rate is not None and rate < warning_rate:
                    shade = '#FCECEA'
            d.rectangle((x, y, x + width, y + row_height), fill=shade, outline='#D9DEE6', width=1)
            for n, line in enumerate(lines):
                d.text((x + 10, y + 12 + n * 28), line, fill='#202020', font=face)
            x += width
        y += row_height
    if highlight_conditions:
        d.text((36, y + 20), '浅黄色记录符合设置的重点标注条件', fill='#707070', font=small)
    elif warning_rate is not None:
        d.text((36, y + 20), f'低于 {warning_rate:g}% 的完成率以浅红色标识', fill='#707070', font=small)
    return canvas


def image_bytes(image: Image.Image) -> bytes:
    """Produce a valid JPG payload under the group robot's documented 2 MB cap.
    Exact encoding is checked before transport. No truncated images are used.
    """
    from PIL import Image

    if image.width <= 0 or image.height <= 0:
        raise ValueError('图片尺寸无效')
    for scale in (1.0, 0.90, 0.78, 0.66):
        scaled = image if scale == 1 else image.resize((max(1, int(image.width * scale)),
                                                         max(1, int(image.height * scale))), Image.Resampling.LANCZOS)
        for quality in (88, 78, 66, 52):
            buf = io.BytesIO()
            scaled.convert('RGB').save(buf, format='JPEG', quality=quality, optimize=True)
            data = buf.getvalue()
            if len(data) <= MAX_IMAGE_BYTES:
                with Image.open(io.BytesIO(data)) as im:
                    im.verify()
                return data
    raise ImagePageOverflow('单页图片即使压缩仍超出 2 MB，需要自动分页')


def markdown_table(job: Job, highlight_conditions: list[dict] | None = None) -> str:
    def esc(s):
        return str(s).replace('|','\\|').replace('\n', ' ').replace('\r',' ').replace('`', '\\`')
    subtitle = (f'按{job.group_field}：{job.group_value} · ' if job.group_field else
                f'{job.store_name} · ' if job.store_name else '')
    lines = [f'## {esc(job.title)}', f'> {esc(subtitle)}共 {len(job.rows)} 条 · 第 {job.part}/{job.parts} 页',
             '', '| ' + ' | '.join(esc(f) for f in job.fields) + ' |',
             '| ' + ' | '.join('---' for _ in job.fields) + ' |']
    if highlight_conditions:
        lines.insert(2, '> 标记 ◆ 的记录符合重点标注条件')
    for r in job.rows:
        values = [text_of(r.get(f), f) for f in job.fields]
        if values and matches_conditions(r, highlight_conditions):
            values[0] = '◆ ' + values[0]
        lines.append('| ' + ' | '.join(esc(value) for value in values) + ' |')
    result = '\n'.join(lines)
    if len(result.encode('utf8')) > MAX_ROBOT_BYTES:
        raise ValueError('Markdown 消息超出 4096 字节，请减小每页行数或列数，或使用图片模式')
    return result



@dataclass
class PreparedMessage:
    job: Job
    mode: str
    payload: bytes | str


def _split_to_fit(job: Job, mode: str, warning_rate: float | None = None,
                  highlight_conditions: list[dict] | None = None):
    """Pack by actual serialized bytes rather than estimated rows/columns."""
    try:
        if mode == '图片表格':
            payload = image_bytes(render_table(job, warning_rate, highlight_conditions))
        elif mode == 'Markdown表格':
            payload = markdown_table(job, highlight_conditions)
            if len(payload.encode('utf-8')) > MAX_MARKDOWN_SAFE:
                raise ValueError('Markdown 安全分页阈值')
        else:
            raise ValueError('未知消息格式')
        return [PreparedMessage(job, mode, payload)]
    except (ValueError, ImagePageOverflow):
        if len(job.rows) > 1:
            middle = len(job.rows) // 2
            left = Job(job.title, job.fields, job.rows[:middle], job.store_code, job.store_name, group_field=job.group_field, group_value=job.group_value)
            right = Job(job.title, job.fields, job.rows[middle:], job.store_code, job.store_name, group_field=job.group_field, group_value=job.group_value)
            return _split_to_fit(left, mode, warning_rate, highlight_conditions) + _split_to_fit(right, mode, warning_rate, highlight_conditions)
        if len(job.fields) > 1:
            middle = len(job.fields) // 2
            left = Job(job.title, job.fields[:middle], job.rows, job.store_code, job.store_name, group_field=job.group_field, group_value=job.group_value)
            right = Job(job.title, job.fields[middle:], job.rows, job.store_code, job.store_name, group_field=job.group_field, group_value=job.group_value)
            return _split_to_fit(left, mode, warning_rate, highlight_conditions) + _split_to_fit(right, mode, warning_rate, highlight_conditions)
        field = job.fields[0]
        raise ValueError(f'一条活动中的「{field}」内容单独成页仍超过推送限制，已取消本批发送；请检查该单元格内容') from None


def prepare_messages(jobs: list[Job], mode: str = '图片表格', warning_rate: float | None = None,
                     max_messages: int = 500, highlight_conditions: list[dict] | None = None) -> list[PreparedMessage]:
    """Safely render/validate the entire batch BEFORE confirming any network send."""
    prepared: list[PreparedMessage] = []
    for job in jobs:
        prepared.extend(_split_to_fit(job, mode, warning_rate, highlight_conditions))
        if len(prepared) > max_messages:
            raise ValueError(f'本批超过 {max_messages} 条消息，已取消预处理；请缩小筛选范围')
    # Page counters change after a split; validate again using the final counters.
    # Iterate until splitting is stable (the counter itself can affect size).
    for _ in range(20):
        original_counters = {id(m.job): (m.job.part, m.job.parts) for m in prepared}
        groups = defaultdict(list)
        for msg in prepared:
            groups[(msg.job.store_code, msg.job.store_name, msg.job.group_field, msg.job.group_value)].append(msg)
        for group in groups.values():
            for idx, msg in enumerate(group, 1):
                msg.job.part, msg.job.parts = idx, len(group)
        final: list[PreparedMessage] = []
        split_again = False
        for msg in prepared:
            try:
                # Avoid duplicate image encoding when make_jobs' counters were already correct.
                if original_counters[id(msg.job)] != (msg.job.part, msg.job.parts):
                    if mode == '图片表格':
                        msg.payload = image_bytes(render_table(msg.job, warning_rate, highlight_conditions))
                    else:
                        msg.payload = markdown_table(msg.job, highlight_conditions)
                final.append(msg)
            except (ImagePageOverflow, ValueError):
                replacement = _split_to_fit(msg.job, mode, warning_rate, highlight_conditions)
                if len(replacement) == 1:
                    raise ValueError('页码变化造成单页超限；该活动字段无法进一步拆分') from None
                final.extend(replacement)
                split_again = True
        if len(final) > max_messages:
            raise ValueError(f'本批超过 {max_messages} 条消息，已取消发送')
        if not split_again:
            return final
        prepared = final
    raise ValueError('分页校验未能收敛，已取消发送，请减少超长字段')


def fit_markdown_jobs(jobs: list[Job]) -> list[Job]:
    """Backward compatible wrapper. New UI uses prepare_messages directly."""
    return [m.job for m in prepare_messages(jobs, 'Markdown表格')]


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


def image_upload_form(data: bytes) -> tuple[dict[str, str], bytes]:
    """WeCom multipart upload with filename, MIME type, and filelength metadata."""
    boundary = 'wecom_' + uuid.uuid4().hex
    head = (f'--{boundary}\r\n'
            f'Content-Disposition: form-data; name="media"; filename="activity_report.jpg"; filelength={len(data)}\r\n'
            f'Content-Type: image/jpeg\r\n\r\n').encode('ascii')
    body = head + data + f'\r\n--{boundary}--\r\n'.encode('ascii')
    return {'Content-Type': f'multipart/form-data; boundary={boundary}'}, body


class WeComSender:
    def __init__(self, kind: str, *, webhook: str = '', corpid: str = '', secret: str = '',
                 agentid: str = '', touser: str = '', timeout: int = 15):
        import requests

        self.kind = kind
        self.webhook = webhook.strip()
        self.corpid = corpid.strip()
        self.secret = secret.strip()
        self.agentid = agentid.strip()
        self.touser = touser.strip()
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update({'User-Agent': USER_AGENT})
        self._token = None
        if kind == '群机器人Webhook':
            if self.webhook:
                check_webhook(self.webhook)
        else:
            if not all((self.corpid, self.secret, self.agentid)) or not self.agentid.isdigit():
                raise ValueError('应用模式需要填写 CorpID、Secret、数字 AgentID')

    def _request(self, method, url, **kwargs):
        import requests

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

    def send(self, job: Job | PreparedMessage, mode: str, recipient: str = '', warning_rate: float | None = None,
             highlight_conditions: list[dict] | None = None):
        prepared = job if isinstance(job, PreparedMessage) else None
        if prepared:
            job, mode = prepared.job, prepared.mode
        if self.kind == '群机器人Webhook':
            url = check_webhook(recipient or self.webhook)
            if mode == '图片表格':
                b = prepared.payload if prepared else image_bytes(render_table(job, warning_rate, highlight_conditions))
                body = {'msgtype':'image', 'image':{'base64':base64.b64encode(b).decode('ascii'),
                                                  'md5':hashlib.md5(b).hexdigest()}}
            elif mode == 'Markdown表格':
                body = {'msgtype':'markdown_v2', 'markdown_v2':{'content':prepared.payload if prepared else markdown_table(job, highlight_conditions)}}
            else:
                raise ValueError('无效消息格式')
            return self._request('POST', url, json=body)
        user = recipient or self.touser
        if not user:
            raise ValueError('请配置 UserID 或导入服务店映射表')
        if mode == '图片表格':
            b = prepared.payload if prepared else image_bytes(render_table(job, warning_rate, highlight_conditions))
            upload_headers, upload_data = image_upload_form(b)
            upload = self._request('POST', 'https://qyapi.weixin.qq.com/cgi-bin/media/upload',
                                   params={'access_token':self.token(), 'type':'image'},
                                   data=upload_data, headers=upload_headers)
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
