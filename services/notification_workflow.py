"""UI-independent orchestration for selecting, preparing, and validating notices."""
from __future__ import annotations

from engine import check_webhook, make_jobs, prepare_messages


def prepare_selection(snapshot: dict):
    selected_ids = set(snapshot['ids'])
    selected = [row for row in snapshot['records'] if row.get('_id') in selected_ids]
    jobs = make_jobs(selected, snapshot['fields'], page_size=12, columns_per_page=6,
                     title=snapshot['title'], group_by_field=snapshot.get('group_field', ''),
                     route_by_store=snapshot.get('route_by_store', False))
    return prepare_messages(jobs, snapshot['mode'],
                            highlight_conditions=snapshot.get('highlight_conditions'))


def validate_routes(jobs, mapping: dict[str, str], kind: str, webhook: str, user: str) -> None:
    if mapping:
        missing = sorted({(job.store_code, job.store_name) for job in jobs
                          if not job.store_code or job.store_code not in mapping})
        if missing:
            details = '、'.join(f'{code or "（无编码）"} {name}' for code, name in missing[:8])
            raise ValueError('以下服务店没有接收映射，已阻止整批发送：' + details)
        if kind == '群机器人Webhook':
            for value in mapping.values():
                check_webhook(value)
    elif kind == '群机器人Webhook':
        check_webhook(webhook)
    elif not user.strip():
        raise ValueError('应用消息需要填写接收人 UserID，或者使用按店的映射')
