"""Persist arbitrary-column routing metadata and OS-vault destination values."""
from __future__ import annotations

import csv
import json
import os
import tempfile
import uuid
from pathlib import Path
from typing import Protocol

from engine import KEY_STORE, check_webhook
from services.config_store import app_data_dir

SERVICE_NAME = 'WeComActivityNotifier'


class SecretBackend(Protocol):
    def get_password(self, service: str, username: str) -> str | None: ...
    def set_password(self, service: str, username: str, password: str) -> None: ...
    def delete_password(self, service: str, username: str) -> None: ...


class MappingStore:
    """Version 2 uses an explicit source field and arbitrary matching values.

    Version 1 mappings are read as KEY_STORE mappings until explicitly changed.
    Destination values remain in the OS vault but are displayed normally by the UI.
    """

    def __init__(self, path: Path | None = None, secret_backend: SecretBackend | None = None,
                 kind: str = 'default'):
        suffix = 'webhook' if kind == '群机器人Webhook' else 'userid' if kind == '自建应用消息' else 'default'
        self.path = path or (app_data_dir() / f'mappings-{suffix}.json')
        self.service_name = f'{SERVICE_NAME}:{suffix}'
        self.kind = kind
        self.secrets = secret_backend

    def _backend(self):
        if self.secrets is None:
            import keyring
            self.secrets = keyring
        return self.secrets

    def _read_document(self) -> dict:
        try:
            data = json.loads(self.path.read_text(encoding='utf-8'))
        except FileNotFoundError:
            return {}
        except ValueError as exc:
            raise ValueError('映射配置无法解析，请备份并检查配置文件。') from exc
        if not isinstance(data, dict) or not isinstance(data.get('mappings'), list):
            raise ValueError('映射配置格式无效。')
        return data

    def settings(self) -> dict:
        """Read only metadata, without touching the credential backend."""
        data = self._read_document()
        rows = data.get('mappings', [])
        field = data.get('field', KEY_STORE if rows and data.get('version', 1) == 1 else '')
        return {'field': str(field or ''), 'count': len(rows)}

    def has_mappings(self) -> bool:
        return bool(self.settings()['count'])

    def as_dict(self) -> dict[str, str]:
        # Missing credentials must reach validation, not silently fall back to
        # the default destination.
        return {row['value']: row['target'] for row in self.all()}

    def all(self) -> list[dict]:
        rows = []
        for item in self._read_document().get('mappings', []):
            target = self._backend().get_password(self.service_name, item['id'])
            rows.append({'id': item['id'], 'value': str(item.get('value', item.get('code', ''))),
                         'target': target or ''})
        return rows

    def save(self, mappings: list[dict], field: str = '') -> None:
        field = field.strip()
        if mappings and not field:
            raise ValueError('请选择映射所匹配的 Excel 列。')
        normalized = []
        seen = set()
        for mapping in mappings:
            value = str(mapping.get('value', '')).strip()
            target = str(mapping.get('target', '')).strip()
            if not value or not target:
                raise ValueError('每条映射都必须填写匹配值和接收目标。')
            if value in seen:
                raise ValueError(f'匹配值重复：{value}')
            if self.kind == '群机器人Webhook':
                check_webhook(target)
            seen.add(value)
            normalized.append({'id': uuid.uuid4().hex, 'value': value, 'target': target})

        old_ids = {item['id'] for item in self._read_document().get('mappings', [])}
        written = []
        temporary = None
        try:
            # Commit a fresh credential revision before atomically replacing
            # metadata, preserving the previous routing if a save fails.
            for item in normalized:
                self._backend().set_password(self.service_name, item['id'], item['target'])
                written.append(item['id'])
            document = {'version': 2, 'field': field,
                        'mappings': [{k: item[k] for k in ('id', 'value')} for item in normalized]}
            self.path.parent.mkdir(parents=True, exist_ok=True)
            fd, temporary = tempfile.mkstemp(prefix='.mappings-', suffix='.tmp', dir=self.path.parent)
            with os.fdopen(fd, 'w', encoding='utf-8') as stream:
                json.dump(document, stream, ensure_ascii=False, indent=2)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary, self.path)
        except Exception:
            self._delete_credentials(written)
            raise
        finally:
            if temporary and os.path.exists(temporary):
                os.unlink(temporary)
        self._delete_credentials(old_ids)

    def _delete_credentials(self, ids):
        for item_id in ids:
            try:
                self._backend().delete_password(self.service_name, item_id)
            except Exception:
                pass

    def import_csv(self, path: str, kind: str) -> list[dict]:
        expected_target = 'Webhook' if kind == '群机器人Webhook' else 'UserID'
        with open(path, 'r', encoding='utf-8-sig', newline='') as stream:
            rows = list(csv.reader(stream))
        if not rows:
            return []
        header = [str(value).strip().casefold() for value in rows[0]]
        value_index = next((i for i, value in enumerate(header)
                            if value in ('匹配值', 'value', 'code', 'store_code') or '编码' in value), 0)
        target_index = next((i for i, value in enumerate(header)
                             if expected_target.casefold() in value or value in ('target', '接收目标')), 1)
        output = []
        for row in rows[1:]:
            if max(value_index, target_index) < len(row):
                value, target = row[value_index].strip(), row[target_index].strip()
                if value and target:
                    output.append({'value': value, 'target': target})
        return output
