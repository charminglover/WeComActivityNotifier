"""UI-independent orchestration for selecting, preparing, and validating notices."""
from __future__ import annotations

from engine import EXCEL_MODE, check_webhook, make_jobs, prepare_messages


def prepare_selection(snapshot: dict):
    selected_ids = set(snapshot['ids'])
    selected = [row for row in snapshot['records'] if row.get('_id') in selected_ids]
    options = {'title': snapshot['title'], 'group_by_field': snapshot.get('group_field', ''),
               'route_field': snapshot.get('route_field', ''),
               'title_rule': snapshot.get('title_rule')}
    attachment_only = snapshot['mode'] == EXCEL_MODE
    jobs = make_jobs(selected, snapshot['fields'], page_size=None if attachment_only else 12,
                     columns_per_page=None if attachment_only else 6, **options)
    messages = prepare_messages(jobs, snapshot['mode'],
                                highlight_conditions=snapshot.get('highlight_conditions'))
    if snapshot.get('include_excel') and not attachment_only:
        attachments = make_jobs(selected, snapshot['fields'], page_size=None,
                                columns_per_page=None, **options)
        messages.extend(prepare_messages(attachments, EXCEL_MODE,
                                         highlight_conditions=snapshot.get('highlight_conditions')))
        if len(messages) > 500:
            raise ValueError('本批消息与附件合计超过 500 条，请缩小所选数据范围')
    return messages


def validate_routes(jobs, mapping: dict[str, str], kind: str, webhook: str, user: str) -> None:
    if mapping:
        missing = sorted({(job.route_field, job.route_value) for job in jobs
                          if not job.route_value or not mapping.get(job.route_value)})
        if missing:
            details = '、'.join(f'{field or "匹配字段"}={value or "（空值）"}'
                               for field, value in missing[:8])
            raise ValueError('以下匹配值没有有效接收目标，已阻止整批发送：' + details)
        if kind == '群机器人Webhook':
            for target in {mapping[job.route_value] for job in jobs}:
                check_webhook(target)
    elif kind == '群机器人Webhook':
        check_webhook(webhook)
    elif not user.strip():
        raise ValueError('应用消息需要填写接收人 UserID，或者配置接收目标映射')
