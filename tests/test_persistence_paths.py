"""在真正的用户数据目录验证保存，覆盖 Windows LocalAppData 重定向。"""

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from models import Config, USER_DIR
from session_store import SessionStore


class PersistencePathTests(unittest.TestCase):
    def test_user_directory_creation_and_replacement(self):
        USER_DIR.mkdir(parents=True, exist_ok=True)
        # 故意保留逻辑路径，不在测试侧 resolve；不改动用户配置和任务。
        with tempfile.TemporaryDirectory(prefix='storage-test-', dir=USER_DIR) as folder:
            base = Path(folder)

            class TestConfig(Config):
                CONFIG_FILE = base / 'config.json'

            store = SessionStore(base / 'session.json')
            for quality in ('vbr', '320k'):
                settings = {'quality': quality, 'output_format': 'mp3'}
                self.assertTrue(TestConfig.save(settings), TestConfig.last_error)
                self.assertEqual(TestConfig.load(), settings)
                self.assertTrue(store.save([], settings), store.last_error)
                self.assertEqual(store.load(), ([], settings))
            self.assertEqual(sorted(p.name for p in base.iterdir()),
                             ['config.json', 'session.json'])


if __name__ == '__main__':
    unittest.main()
