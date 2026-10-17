"""Versioned, atomic persistence for non-sensitive user preferences."""
from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path

APP_DIR_NAME = 'WeComActivityNotifier'


def app_data_dir() -> Path:
    if os.name == 'nt':
        return Path(os.environ.get('APPDATA', str(Path.home()))) / APP_DIR_NAME
    if sys.platform == 'darwin':
        return Path.home() / 'Library' / 'Application Support' / APP_DIR_NAME
    return Path(os.environ.get('XDG_CONFIG_HOME', str(Path.home() / '.config'))) / APP_DIR_NAME


class ConfigStore:
    def __init__(self, path: Path | None = None):
        self.path = path or (app_data_dir() / 'preferences.json')

    def load(self) -> dict:
        try:
            value = json.loads(self.path.read_text(encoding='utf-8'))
            return value if isinstance(value, dict) else {}
        except (OSError, ValueError):
            return {}

    def save(self, values: dict) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(prefix='.preferences-', suffix='.tmp', dir=self.path.parent)
        try:
            with os.fdopen(fd, 'w', encoding='utf-8') as stream:
                json.dump({'version': 1, **values}, stream, ensure_ascii=False, indent=2)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary, self.path)
        finally:
            try:
                os.unlink(temporary)
            except FileNotFoundError:
                pass
