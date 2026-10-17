import json
import tempfile
import unittest
import sys
import uuid
from pathlib import Path
from unittest.mock import patch

from services.config_store import ConfigStore
from services.mapping_store import MappingStore


class FakeKeyring:
    def __init__(self):
        self.values = {}
    def get_password(self, service, username):
        return self.values.get((service, username))
    def set_password(self, service, username, password):
        self.values[(service, username)] = password
    def delete_password(self, service, username):
        self.values.pop((service, username), None)


class TestStores(unittest.TestCase):
    def test_preferences_are_versioned_json_and_atomic(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ConfigStore(Path(directory) / 'preferences.json')
            store.save({'mode': 'Markdown表格', 'fields': ['Name']})
            self.assertEqual(store.load()['mode'], 'Markdown表格')
            self.assertEqual(json.loads(store.path.read_text())['version'], 1)

    def test_mapping_target_stays_out_of_json(self):
        with tempfile.TemporaryDirectory() as directory:
            backend = FakeKeyring()
            path = Path(directory) / 'mappings.json'
            store = MappingStore(path, backend, kind='群机器人Webhook')
            secret_url = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=very-secret'
            store.save([{'value': 'North', 'target': secret_url}], field='Region')
            reopened = MappingStore(path, backend, kind='群机器人Webhook')
            self.assertEqual(reopened.settings(), {'field': 'Region', 'count': 1})
            self.assertEqual(reopened.as_dict(), {'North': secret_url})
            self.assertNotIn('very-secret', path.read_text(encoding='utf-8'))
            credentials_before = dict(backend.values)
            with self.assertRaises(ValueError):
                store.save([{'value': 'North', 'target': secret_url},
                            {'value': 'North', 'target': secret_url}], field='Region')
            self.assertEqual(backend.values, credentials_before)

    def test_failed_save_keeps_previously_committed_routes(self):
        with tempfile.TemporaryDirectory() as directory:
            backend = FakeKeyring()
            store = MappingStore(Path(directory) / 'mappings.json', backend)
            store.save([{'value': 'North', 'target': 'old-user'}], field='Region')
            before = store.path.read_bytes()
            original_write = backend.set_password
            calls = []

            def failing_write(service, username, password):
                calls.append(username)
                if len(calls) == 2:
                    raise RuntimeError('Credential write failed')
                original_write(service, username, password)

            with patch.object(backend, 'set_password', side_effect=failing_write):
                with self.assertRaises(RuntimeError):
                    store.save([{'value': 'East', 'target': 'new-user'},
                                {'value': 'West', 'target': 'other-user'}], field='Department')
            self.assertEqual(store.path.read_bytes(), before)
            self.assertEqual(store.as_dict(), {'North': 'old-user'})
            self.assertEqual(len(backend.values), 1)

    def test_legacy_mapping_migrates_without_losing_target(self):
        with tempfile.TemporaryDirectory() as directory:
            backend = FakeKeyring()
            store = MappingStore(Path(directory) / 'mappings.json', backend)
            store.path.write_text(json.dumps({'version': 1, 'mappings': [
                {'id': 'legacy', 'code': 'S001'}]}), encoding='utf-8')
            backend.set_password(store.service_name, 'legacy', 'legacy-user')
            self.assertEqual(store.settings()['field'], '任务服务店编码')
            rows = store.all()
            self.assertEqual(rows[0]['value'], 'S001')
            store.save(rows, field='Department')
            self.assertEqual(store.settings()['field'], 'Department')
            self.assertEqual(store.as_dict(), {'S001': 'legacy-user'})
            self.assertEqual(json.loads(store.path.read_text())['version'], 2)

    @unittest.skipUnless(sys.platform == 'win32', 'Requires Windows credential vault')
    def test_real_windows_vault_save_reopen_edit_delete(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'mappings.json'
            store = MappingStore(path, kind='自建应用消息')
            store.service_name += ':test:' + uuid.uuid4().hex
            try:
                store.save([{'value': 'North', 'target': 'test-user'}], field='Region')
                reopened = MappingStore(path, kind='自建应用消息')
                reopened.service_name = store.service_name
                self.assertEqual(reopened.as_dict(), {'North': 'test-user'})
                reopened.save([{'value': 'East', 'target': 'edited-user'}], field='Department')
                self.assertEqual(store.as_dict(), {'East': 'edited-user'})
            finally:
                store.save([], field='')
            self.assertFalse(store.has_mappings())


if __name__ == '__main__':
    unittest.main()
