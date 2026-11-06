import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from models import FileItem
from session_store import SessionStore, file_signature


class SessionTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix='qqmusic_session_')
        self.addCleanup(temporary.cleanup)
        self.base = Path(temporary.name).resolve()
        self.store = SessionStore(self.base / 'session.json')
        self.source = self.base / 'source.mp3'
        self.source.write_bytes(b'original')
        self.output = self.base / 'output.mp3'
        self.output.write_bytes(b'converted')
        self.item = FileItem('1', self.source, self.source.name, self.source.stat().st_size,
                             'done', 'standard_mp3', output_path=self.output,
                             source_signature=file_signature(self.source),
                             output_signature=file_signature(self.output))

    def save(self):
        self.assertTrue(self.store.save([self.item], {'quality': '320k', 'output_format': 'mp3',
                                                     'output_path': str(self.base)}))

    def test_completed_record_restores_with_options(self):
        self.save()
        items, settings = self.store.load()
        self.assertEqual(items[0].status, 'done')
        self.assertEqual(items[0].output_path, self.output)
        self.assertEqual(settings['quality'], '320k')
        self.assertEqual(settings['output_format'], 'mp3')

    def test_missing_or_modified_output_requires_reprocessing(self):
        self.save()
        self.output.write_bytes(b'modified output')
        items, _ = self.store.load()
        self.assertEqual(items[0].status, 'waiting')
        self.output.unlink()
        items, _ = self.store.load()
        self.assertEqual(items[0].status, 'waiting')

    def test_changed_source_is_not_treated_as_completed(self):
        self.save()
        current = self.source.stat()
        os.utime(self.source, ns=(current.st_atime_ns, current.st_mtime_ns + 2000000000))
        items, _ = self.store.load()
        self.assertEqual(items[0].status, 'waiting')
        self.assertNotEqual(items[0].source_signature, self.item.source_signature)

    def test_interrupted_item_waits_and_missing_source_has_reason(self):
        self.item.status = 'converting'
        self.save()
        items, _ = self.store.load()
        self.assertEqual(items[0].status, 'waiting')
        self.source.unlink()
        items, _ = self.store.load()
        self.assertEqual(items[0].status, 'error')
        self.assertIn('源文件', items[0].error_msg)
        self.assertEqual(items[0].output_path, self.output)

    def test_broken_records_do_not_prevent_other_items_restoring(self):
        self.save()
        data = json.loads(self.store.path.read_text(encoding='utf-8'))
        data['files'].insert(0, {'bad': 'record'})
        data['files'].append(None)
        self.store.path.write_text(json.dumps(data), encoding='utf-8')
        items, _ = self.store.load()
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0].status, 'done')
        self.assertTrue(self.store.last_error)

    def test_failed_save_preserves_previous_snapshot(self):
        self.save()
        previous = self.store.path.read_bytes()
        with patch('session_store.os.replace', side_effect=PermissionError('locked')):
            self.assertFalse(self.store.save([], {}))
        self.assertEqual(self.store.path.read_bytes(), previous)
        self.assertEqual(len(list(self.base.iterdir())), 3)


if __name__ == '__main__':
    unittest.main()
