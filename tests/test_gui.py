import sys
import tempfile
import threading
import time
import unittest
import wave
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch, Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from gui import ConverterGUI, create_root
from models import Config
from session_store import SessionStore
from converter import AudioConverter


class GuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as folder:
            source, output = Path(folder) / 'test.wav', Path(folder) / 'test.mp3'
            with wave.open(str(source), 'wb') as stream:
                stream.setparams((1, 2, 44100, 0, 'NONE', 'not compressed'))
                stream.writeframes(b'\x00\x00' * 11025)
            AudioConverter().wav_to_mp3(source, output, 'vbr')
            cls.audio_bytes = output.read_bytes()

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='qqmusic_gui_')
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.config_patch = patch.object(Config, 'CONFIG_FILE', self.base / 'config.json')
        self.config_patch.start()
        self.addCleanup(self.config_patch.stop)
        self.messages = patch('gui.messagebox')
        self.messages.start()
        self.addCleanup(self.messages.stop)
        self.root = create_root()
        self.root.withdraw()
        self.app = ConverterGUI(self.root, SessionStore(self.base / 'session.json'))
        self.root.withdraw()
        self.app.output_path_var.set(str(self.base / 'out'))
        self.root.update_idletasks()
        self.app.batch_converter.log_path = self.base / 'error.log'
        self.callback_errors = []
        self.root.report_callback_exception = lambda *args: self.callback_errors.append(args)
        self.addCleanup(lambda: self.app._destroy())

    def pump_until(self, predicate, timeout=5):
        end = time.monotonic() + timeout
        while not predicate() and time.monotonic() < end:
            self.root.update()
            time.sleep(0.01)
        self.assertTrue(predicate(), 'GUI operation timed out')
        self.assertEqual(self.callback_errors, [])

    def source(self, name='song.mp3'):
        path = self.base / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(self.audio_bytes)
        return path

    def test_scan_deduplicates_and_conversion_uses_snapshot(self):
        path = self.source('中文 空格/song.mp3')
        self.app._start_scan([str(path.parent), str(path)])
        self.pump_until(lambda: not self.app.is_scanning)
        self.assertEqual(len(self.app.files), 1)
        self.app.output_format_var.set('keep')
        self.app.start_conversion()
        self.assertTrue(self.app.is_converting)
        self.app.remove_selected()
        self.pump_until(lambda: not self.app.is_converting)
        item = next(iter(self.app.files.values()))
        self.assertEqual(item.status, 'done', item.error_msg)
        self.assertEqual(item.output_path.read_bytes(), self.audio_bytes)

    def test_empty_scanning_and_completed_queue_disable_start(self):
        self.assertTrue(self.app.convert_btn.instate(['disabled']))
        self.assertFalse(self.app.progress_bar.winfo_manager())
        self.assertEqual(self.app.empty_hint.winfo_manager(), 'place')
        self.app._start_scan([str(self.source())])
        self.assertTrue(self.app.convert_btn.instate(['disabled']))
        self.pump_until(lambda: not self.app.is_scanning)
        self.assertTrue(self.app.convert_btn.instate(['!disabled']))
        self.assertFalse(self.app.empty_hint.winfo_manager())
        self.app.start_conversion()
        self.assertEqual(self.app.progress_bar.winfo_manager(), 'grid')
        self.pump_until(lambda: not self.app.is_converting)
        self.assertTrue(self.app.convert_btn.instate(['disabled']))
        self.assertFalse(self.app.progress_bar.winfo_manager())
        self.assertFalse(self.app.stage_bar.winfo_manager())
        self.app.tree.selection_set(next(iter(self.app.files)))
        self.app.remove_selected()
        self.assertEqual(self.app.empty_hint.winfo_manager(), 'place')

    def test_invalid_output_path_shows_error_without_starting_worker(self):
        self.app.add_file_item(self.source())
        self.app.output_path_var.set(str(self.base / 'invalid') + '\x00')
        with patch('gui.messagebox.showerror') as error:
            self.app.start_conversion()
            error.assert_called_once()
        self.assertFalse(self.app.is_converting)
        self.assertIsNone(self.app.worker_thread)
        self.assertEqual(next(iter(self.app.files.values())).status, 'waiting')

    def test_open_output_rejects_empty_and_invalid_paths(self):
        for raw in ('  ', str(self.base / 'invalid') + '\x00'):
            with self.subTest(path=repr(raw)), patch('gui.os.startfile') as startfile, \
                 patch('gui.messagebox.showerror') as error:
                self.app.output_path_var.set(raw)
                self.app.open_output_folder()
                startfile.assert_not_called()
                error.assert_called_once()

    def test_successful_retry_clears_only_its_previous_save_error(self):
        def fail_save(*args):
            self.app.session_store.last_error = '任务记录保存失败: locked'
            return False
        with patch.object(self.app.session_store, 'save', side_effect=fail_save):
            self.assertFalse(self.app._save_session())
        self.assertIn('保存失败', self.app.status_bar.cget('text'))
        self.assertTrue(self.app._save_session())
        self.assertEqual(self.app.status_bar.cget('text'), '任务记录已保存')
        with patch.object(self.app.session_store, 'save', side_effect=fail_save):
            self.app._save_session()
        self.app.status_bar.config(text='已处理 1/2')
        self.assertTrue(self.app._save_session())
        self.assertEqual(self.app.status_bar.cget('text'), '已处理 1/2')

    def test_quality_controls_and_retry_only_failed(self):
        self.app.output_format_var.set('keep')
        self.assertTrue(all(w.instate(['disabled']) for w in self.app.quality_frame.winfo_children()))
        self.app.output_format_var.set('mp3')
        self.assertTrue(all(w.instate(['!disabled']) for w in self.app.quality_frame.winfo_children()))
        for name in ('failed.mp3', 'waiting.mp3', 'done.mp3'):
            self.app.add_file_item(self.source(name))
        items = list(self.app.files.values())
        items[0].status = 'error'
        items[0].error_msg = 'previous failure'
        items[2].status = 'done'
        with patch.object(self.app, 'start_conversion') as start:
            self.app.retry_failed()
        start.assert_called_once_with(only_ids={items[0].id})
        self.assertEqual([i.status for i in items], ['waiting', 'waiting', 'done'])
        self.assertEqual(items[0].error_msg, '')

    def test_close_waits_for_current_file_and_cleanup(self):
        self.app.add_file_item(self.source())
        entered, release = threading.Event(), threading.Event()
        original = self.app.batch_converter.process_single
        def blocked(*args, **kwargs):
            entered.set()
            release.wait(3)
            return original(*args, **kwargs)
        with patch.object(self.app.batch_converter, 'process_single', side_effect=blocked), \
             patch.object(self.app, '_destroy') as destroy:
            self.app.start_conversion()
            self.assertTrue(entered.wait(2))
            self.app.on_close()
            self.assertTrue(self.app.stop_event.is_set())
            destroy.assert_not_called()
            release.set()
            self.pump_until(lambda: destroy.called)
            self.assertFalse(self.app.worker_thread.is_alive())

    def test_drag_drop_handles_spaces_and_unicode(self):
        if not self.root.dnd_available:
            self.skipTest('tkinterdnd2 is not installed')
        paths = [self.source('中文 文件.mp3'), self.source('a b/other.mp3')]
        event = SimpleNamespace(data=' '.join('{' + str(p) + '}' for p in paths))
        self.assertEqual(self.app.on_drop(event), 'copy')
        self.pump_until(lambda: not self.app.is_scanning)
        self.assertEqual({item.path for item in self.app.files.values()}, {p.resolve() for p in paths})

    def test_reprocess_selected_preserves_old_output_and_other_waiting_items(self):
        self.app.output_format_var.set('keep')
        self.app.add_file_item(self.source())
        self.app.add_file_item(self.source('other.mp3'))
        first, other = self.app.files.values()
        first.status = 'done'
        out = self.base / 'out'
        out.mkdir()
        original = out / 'song.mp3'
        original.write_bytes(b'previous output')
        self.app.tree.selection_set(first.id)
        self.app.reprocess_selected()
        self.pump_until(lambda: not self.app.is_converting)
        self.assertEqual(first.status, 'done')
        self.assertEqual(other.status, 'waiting')
        self.assertEqual(original.read_bytes(), b'previous output')
        self.assertNotEqual(first.output_path, original)
        self.assertEqual(first.output_path.read_bytes(), self.audio_bytes)

    def test_relaunch_restores_queue_and_does_not_start_automatically(self):
        self.app.output_format_var.set('keep')
        self.app.add_file_item(self.source())
        self.app.start_conversion()
        self.pump_until(lambda: not self.app.is_converting)
        self.app.add_file_item(self.source('waiting.mp3'))
        self.app.quality_var.set('192k')
        self.app._destroy()
        self.root = create_root()
        self.root.withdraw()
        self.app = ConverterGUI(self.root, SessionStore(self.base / 'session.json'))
        self.root.update_idletasks()
        self.assertEqual([item.status for item in self.app.files.values()], ['done', 'waiting'])
        self.assertEqual(self.app.quality_var.get(), '192k')
        self.assertFalse(self.app.is_converting)
        self.assertIsNone(self.app.worker_thread)

    def test_details_show_paths_and_audio_properties(self):
        source = self.base / 'tone.wav'
        with wave.open(str(source), 'wb') as stream:
            stream.setparams((1, 2, 44100, 0, 'NONE', 'not compressed'))
            stream.writeframes(b'\x00\x00' * 4410)
        self.app.add_file_item(source)
        item = next(iter(self.app.files.values()))
        self.app.show_file_details(item.id)
        window, text = next(iter(self.app._detail_requests.values()))
        window.withdraw()
        self.pump_until(lambda: not self.app._detail_requests)
        report = text.get('1.0', 'end')
        self.assertIn(str(source.resolve()), report)
        self.assertIn('44100 Hz', report)
        self.assertIn('声道: 1', report)
        window.destroy()

    def test_stage_progress_and_elapsed_time_reach_the_ui(self):
        self.app.add_file_item(self.source())
        item = next(iter(self.app.files.values()))
        self.app.events.put(('status', item.id, '处理中', 'blue'))
        self.app.events.put(('stage', item.id, '编码', 42))
        self.pump_until(lambda: self.app._active_stage is not None)
        self.assertEqual(float(self.app.stage_bar['value']), 42)
        label = self.app.status_label.cget('text')
        self.assertIn('编码', label)
        self.assertIn('已用', label)
        self.assertIn('42%', label)

    def test_second_instance_notification_wakes_existing_window(self):
        self.app.instance_guard = Mock(activation_requested=Mock(return_value=True))
        with patch.object(self.root, 'deiconify') as show, patch.object(self.root, 'lift') as lift, \
             patch.object(self.root, 'focus_force') as focus:
            self.pump_until(lambda: show.called)
            lift.assert_called()
            focus.assert_called()
        self.app.instance_guard = None

    def test_environment_diagnostics_render_in_report_window(self):
        with patch('gui.environment_report', return_value='客户端版本: 22.05\n必要接口检查通过'):
            self.app.show_environment_diagnostics()
            window, text = next(iter(self.app._detail_requests.values()))
            window.withdraw()
            self.pump_until(lambda: not self.app._detail_requests)
            self.assertIn('22.05', text.get('1.0', 'end'))
            window.destroy()

    def test_full_validation_result_and_setting_are_saved(self):
        if not self.app.converter.ffmpeg_ok:
            self.skipTest('ffmpeg unavailable')
        self.app.verify_full_var.set(True)
        self.app.add_file_item(self.source())
        self.app.start_conversion()
        self.pump_until(lambda: not self.app.is_converting)
        item = next(iter(self.app.files.values()))
        self.assertEqual(item.status, 'done', item.error_msg)
        self.assertIn('完整解码校验通过', item.validation_msg)
        restored, settings = self.app.session_store.load()
        self.assertTrue(settings['verify_full'])
        self.assertEqual(restored[0].validation_msg, item.validation_msg)


if __name__ == '__main__':
    unittest.main()
