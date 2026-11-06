import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from client_diagnostics import assess_client, connection_error, inspect_session, REQUIRED_EXPORTS
from decryptor import QQMusicDecryptor


class ClientDiagnosticsTests(unittest.TestCase):
    def test_preflight_distinguishes_architecture_module_and_exports(self):
        self.assertIn('架构不兼容', assess_client({'arch': 'x64'}))
        self.assertIn('尚未加载', assess_client({'arch': 'ia32', 'module': None}))
        self.assertIn('缺少必要接口', assess_client({'arch': 'ia32', 'module': 'dll', 'missing_exports': ['Open']}))
        self.assertEqual(assess_client({'arch': 'ia32', 'module': 'dll', 'missing_exports': []}), '')

    def test_preflight_unloads_probe_and_all_hook_exports_are_checked(self):
        script = Mock()
        script.exports_sync.inspect.return_value = dict(arch='ia32', module='dll', missing_exports=[], executable='QQMusic.exe')
        session = Mock(create_script=Mock(return_value=script))
        with patch('client_diagnostics.file_version', return_value='22.05'):
            report = inspect_session(session)
        self.assertEqual(report['client_version'], '22.05')
        script.unload.assert_called_once()
        hook = (Path(__file__).resolve().parents[1] / 'src' / 'hook_qqmusic_v17.js').read_text(encoding='utf-8')
        self.assertTrue(all(name in hook for name in REQUIRED_EXPORTS))

    def test_incompatible_client_is_detached_before_loading_hook(self):
        session = Mock()
        fake_frida = SimpleNamespace(attach=Mock(return_value=session))
        decryptor = QQMusicDecryptor()
        with patch.dict(sys.modules, {'frida': fake_frida}), \
             patch('decryptor.inspect_session', return_value={'error': '客户端版本不兼容'}):
            with self.assertRaisesRegex(RuntimeError, '版本不兼容'):
                decryptor.connect()
        session.create_script.assert_not_called()
        session.detach.assert_called_once()
        self.assertFalse(decryptor._connected)

    def test_connection_errors_are_actionable(self):
        missing = type('ProcessNotFoundError', (Exception,), {})('not found')
        denied = type('PermissionDeniedError', (Exception,), {})('denied')
        self.assertIn('未运行', connection_error(missing))
        self.assertIn('运行权限', connection_error(denied))


if __name__ == '__main__':
    unittest.main()
