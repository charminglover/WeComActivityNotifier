"""Store routing metadata in JSON and destinations in the OS credential vault."""
from __future__ import annotations

import csv
import json
import uuid
from pathlib import Path
from typing import Protocol

from services.config_store import app_data_dir

SERVICE_NAME = 'WeComActivityNotifier'


class SecretBackend(Protocol):
    def get_password(self, service: str, username: str) -> str | None: ...
    def set_password(self, service: str, username: str, password: str) -> None: ...
    def delete_password(self, service: str, username: str) -> None: ...


class MappingStore:
    """Mappings contain an id/code in JSON; webhook/UserID values use OS keychain."""

    def __init__(self, path: Path | None = None, secret_backend: SecretBackend | None = None,
                 kind: str = 'default'):
        suffix = 'webhook' if kind == '群机器人Webhook' else 'userid' if kind == '自建应用消息' else 'default'
        self.path = path or (app_data_dir() / f'mappings-{suffix}.json')
        self.service_name = f'{SERVICE_NAME}:{suffix}'
        self.secrets = secret_backend

    def _backend(self):
        if self.secrets is None:
            import keyring
            self.secrets = keyring
        return self.secrets

    def _read(self) -> list[dict]:
        try:
            data = json.loads(self.path.read_text(encoding='utf-8'))
            return data.get('mappings', []) if isinstance(data, dict) else []
        except (OSError, ValueError):
            return []

    def has_mappings(self) -> bool:
        return bool(self._read())

    def as_dict(self) -> dict[str, str]:
        return {row['code']: row['target'] for row in self.all() if row['target']}

    def all(self) -> list[dict]:
        rows = []
        for item in self._read():
            target = self._backend().get_password(self.service_name, item['id'])
            rows.append({'id': item['id'], 'code': item.get('code', ''), 'target': target or ''})
        return rows

    def save(self, mappings: list[dict]) -> None:
        normalized = []
        seen = set()
        for mapping in mappings:
            code = str(mapping.get('code', '')).strip()
            target = str(mapping.get('target', '')).strip()
            item_id = str(mapping.get('id') or uuid.uuid4().hex)
            if not code or not target:
                raise ValueError('每条映射都必须填写服务店编码和接收目标。')
            if code in seen:
                raise ValueError(f'服务店编码重复：{code}')
            seen.add(code)
            self._backend().set_password(self.service_name, item_id, target)
            normalized.append({'id': item_id, 'code': code})
        old_ids = {item['id'] for item in self._read()}
        new_ids = {item['id'] for item in normalized}
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix('.tmp')
        temporary.write_text(json.dumps({'version': 1, 'mappings': normalized}, ensure_ascii=False, indent=2), encoding='utf-8')
        temporary.replace(self.path)
        for removed_id in old_ids - new_ids:
            try:
                self._backend().delete_password(self.service_name, removed_id)
            except Exception:
                pass

    def import_csv(self, path: str, kind: str) -> list[dict]:
        expected_target = 'Webhook' if kind == '群机器人Webhook' else 'UserID'
        with open(path, 'r', encoding='utf-8-sig', newline='') as stream:
            rows = list(csv.reader(stream))
        if not rows:
            return []
        header = [str(value).strip().casefold() for value in rows[0]]
        code_index = next((i for i, value in enumerate(header) if '编码' in value or value in ('code', 'store_code')), 0)
        target_index = next((i for i, value in enumerate(header) if expected_target.casefold() in value or value == 'target'), 1)
        output = []
        for row in rows[1:]:
            if max(code_index, target_index) < len(row):
                code, target = row[code_index].strip(), row[target_index].strip()
                if code and target:
                    output.append({'code': code, 'target': target})
        return output
