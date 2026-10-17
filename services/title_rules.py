"""Compile visual title settings into literal and Excel-field segments."""
from __future__ import annotations

import json
import re
from dataclasses import dataclass

_TOKEN = re.compile(r'\{列:|\{group_field\}|\{group_value\}|\{route_field\}|\{route_value\}')


def field_token(field: str) -> str:
    # JSON quoting supports headers containing braces, quotes, or backslashes.
    return '{列:' + json.dumps(field, ensure_ascii=False) + '}'


@dataclass(frozen=True)
class TitleRule:
    segments: tuple[tuple[str, str], ...]
    fields: tuple[str, ...]

    def render(self, values: dict[str, str]) -> str:
        title = ''.join(value if kind == 'text' else
                        (values.get(value, '').strip() or f'{value}未填写')
                        for kind, value in self.segments).strip()
        if not title:
            raise ValueError('生成的标题为空，请设置标题列或自定义文字。')
        return title


def compile_title_rule(config: dict, available_fields, group_field='', route_field='') -> TitleRule:
    available = set(available_fields)
    segments = []
    dependencies = []

    def add_field(field):
        if not field:
            raise ValueError('请选择用于生成标题的 Excel 列。')
        if field not in available:
            raise ValueError(f'标题引用的列“{field}”不在当前文件中，请重新选择。')
        segments.append(('field', field))
        if field not in dependencies:
            dependencies.append(field)

    mode = config.get('mode', 'auto')
    if mode in ('auto', 'field'):
        segments.append(('text', str(config.get('prefix', ''))))
        source = (group_field or route_field) if mode == 'auto' else str(config.get('field', ''))
        if source or mode == 'field':
            add_field(source)
        else:
            segments.append(('text', '数据'))
        segments.append(('text', str(config.get('suffix', '通知'))))
    elif mode == 'custom':
        template = str(config.get('template', ''))
        if not template.strip():
            raise ValueError('请填写自定义标题，或用“插入列值”组合动态标题。')
        position = 0
        while match := _TOKEN.search(template, position):
            segments.append(('text', template[position:match.start()]))
            token = match.group()
            if token == '{列:':
                try:
                    field, end = json.JSONDecoder().raw_decode(template, match.end())
                except ValueError as exc:
                    raise ValueError('标题中的列标记格式不完整，请重新使用“插入列值”。') from exc
                if not isinstance(field, str) or template[end:end + 1] != '}':
                    raise ValueError('标题中的列标记格式不完整，请重新使用“插入列值”。')
                add_field(field)
                position = end + 1
            else:
                source = group_field if token.startswith('{group_') else route_field
                if not source:
                    name = '分组字段' if token.startswith('{group_') else '映射匹配字段'
                    raise ValueError(f'标题使用了{token}，请先设置{name}或直接插入 Excel 列。')
                if token.endswith('_value}'):
                    add_field(source)
                else:
                    segments.append(('text', source))
                position = match.end()
        segments.append(('text', template[position:]))
    else:
        raise ValueError('未知的标题生成方式，请重新选择。')
    return TitleRule(tuple(segments), tuple(dependencies))
