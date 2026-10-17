import json
import tempfile
import unittest
from pathlib import Path

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
            store.save([{'code': 'S001', 'target': secret_url}])
            self.assertEqual(store.as_dict(), {'S001': secret_url})
            self.assertNotIn('very-secret', path.read_text(encoding='utf-8'))
            with self.assertRaises(ValueError):
                store.save([{'code': 'S001', 'target': secret_url},
                            {'code': 'S001', 'target': secret_url}])


if __name__ == '__main__':
    unittest.main()
