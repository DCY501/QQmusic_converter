import math
import struct
import sys
import tempfile
import threading
import unittest
import wave
import shutil
import subprocess
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from unittest.mock import patch, Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from converter import AudioConverter
from models import FileItem
from worker import BatchConverter


class OfflineDecryptor:
    def __init__(self):
        self.connections = 0
        self.disconnected = False

    def connect(self):
        self.connections += 1
        raise RuntimeError("QQ Music is offline")

    def disconnect(self):
        self.disconnected = True


class ConversionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="qqmusic_test_")
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.output = self.base / "out"
        self.output.mkdir()
        self.decryptor = OfflineDecryptor()
        self.converter = AudioConverter()
        # 此组主要验证调度/提交；损坏音频和真实校验在 test_audio_validation 中覆盖。
        self.worker = BatchConverter(self.decryptor, self.converter, self.base / "error.log",
                                     validator=Mock(validate=Mock(return_value='test validation')))

    def item(self, name="a/song.mp3", kind="standard_mp3", data=b"original"):
        path = self.base / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return FileItem(name, path, path.name, len(data), "waiting", kind)

    def test_same_names_preserve_existing_files(self):
        existing = self.output / "song.mp3"
        existing.write_bytes(b"existing")
        items = [self.item(data=b"first"), self.item("b/song.mp3", data=b"second")]
        stats = self.worker.process_batch(items, self.output, "keep", "vbr")
        self.assertEqual(stats["done"], 2)
        self.assertEqual(existing.read_bytes(), b"existing")
        self.assertEqual(items[0].output_path.read_bytes(), b"first")
        self.assertEqual(items[1].output_path.read_bytes(), b"second")
        self.assertEqual(len(list(self.output.iterdir())), 3)

    def test_concurrent_publish_does_not_overwrite(self):
        items = [self.item(f"{i}/song.mp3", data=str(i).encode()) for i in range(8)]
        with ThreadPoolExecutor(max_workers=4) as pool:
            results = list(pool.map(lambda x: self.worker.process_single(x, self.output, "keep", "vbr"), items))
        self.assertEqual(len({r.output_path for r in results}), 8)
        self.assertEqual({p.read_bytes() for p in self.output.iterdir()}, {str(i).encode() for i in range(8)})

    def test_source_is_destination_is_skipped(self):
        item = self.item("out/song.mp3")
        result = self.worker.process_single(item, self.output, "keep", "vbr")
        self.assertEqual(result.status, "skipped")
        self.assertEqual(item.path.read_bytes(), b"original")

    def test_offline_decryption_does_not_block_plain_files(self):
        items = [self.item("a/one.mflac", "qmc_mflac"), self.item(),
                 self.item("a/two.mflac", "qmc_mflac")]
        stats = self.worker.process_batch(items, self.output, "keep", "vbr")
        self.assertEqual((stats["done"], stats["error"]), (1, 2))
        self.assertEqual(self.decryptor.connections, 1)
        self.assertTrue(self.decryptor.disconnected)
        self.assertIn("QQ Music is offline", (self.base / "error.log").read_text(encoding="utf-8"))

    def test_failure_discards_partial_output_and_preserves_old_output(self):
        item = self.item("a/song.wav", "standard_wav")
        existing = self.output / "song.mp3"
        existing.write_bytes(b"keep me")
        def fail(source, destination, kind, quality, **kwargs):
            destination.write_bytes(b"partial")
            raise RuntimeError("encoder diagnostic")
        with patch.object(self.converter, "convert_to_mp3", side_effect=fail):
            stats = self.worker.process_batch([item], self.output, "mp3", "vbr")
        self.assertEqual(stats["error"], 1)
        self.assertEqual(item.error_msg, "encoder diagnostic")
        self.assertEqual(list(self.output.iterdir()), [existing])
        self.assertEqual(existing.read_bytes(), b"keep me")

    def test_stop_leaves_remaining_files_waiting(self):
        items = [self.item(), self.item("b/other.mp3")]
        stop = threading.Event()
        stats = self.worker.process_batch(items, self.output, "keep", "vbr",
                                          on_progress=lambda *_: stop.set(), stop_check=stop.is_set)
        self.assertTrue(stats["stopped"])
        self.assertEqual(stats["remaining"], 1)
        self.assertEqual([i.status for i in items], ["done", "waiting"])
        resumed = self.worker.process_batch([items[1]], self.output, "keep", "vbr")
        self.assertEqual(resumed["done"], 1)

    def test_unsupported_mp4_is_skipped_without_connecting(self):
        self.converter.ffmpeg_ok = False
        item = self.item("a/song.mmp4", "qmc_mmp4")
        stats = self.worker.process_batch([item], self.output, "mp3", "vbr")
        self.assertEqual(stats["skipped"], 1)
        self.assertEqual(self.decryptor.connections, 0)

    def test_mp3_copy_does_not_require_lame(self):
        self.converter.lame_ok = False
        stats = self.worker.process_batch([self.item()], self.output, "mp3", "vbr")
        self.assertEqual(stats["done"], 1)

    def test_tool_failure_and_timeout_have_diagnostics(self):
        with self.assertRaisesRegex(RuntimeError, "specific error"):
            self.converter._run([sys.executable, "-c", "import sys; sys.stderr.write('specific error'); sys.exit(3)"])
        with self.assertRaisesRegex(RuntimeError, "0.1"):
            self.converter._run([sys.executable, "-c", "import time; time.sleep(10)"], timeout=0.1)

    def test_hook_checks_partial_reads_and_releases_resources(self):
        node = shutil.which('node')
        if not node:
            self.skipTest('node is not installed')
        result = subprocess.run([node, str(Path(__file__).with_name('test_hook.js'))],
                                capture_output=True, text=True, encoding='utf-8')
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_tag_failure_preserves_encoded_audio_with_warning(self):
        item = self.item('a/song.wav', 'standard_wav')
        def convert(source, destination, kind, quality, **kwargs):
            destination.write_bytes(b'encoded audio')
        def bad_tags(source, destination):
            destination.write_bytes(b'broken tag output')
            raise RuntimeError('invalid cover')
        with patch.object(self.converter, 'convert_to_mp3', side_effect=convert), \
             patch('worker.copy_tags', side_effect=bad_tags):
            stats = self.worker.process_batch([item], self.output, 'mp3', 'vbr')
        self.assertEqual((stats['done'], stats['warnings']), (1, 1))
        self.assertEqual(item.output_path.read_bytes(), b'encoded audio')
        self.assertIn('invalid cover', item.warning_msg)

    def test_real_ogg_decode_with_unicode_and_spaces_in_temp_directory(self):
        if not (self.converter.ffmpeg_ok and self.converter.oggdec_ok):
            self.skipTest('ffmpeg or oggdec unavailable')
        ogg = self.base / 'song.ogg'
        self.converter._run([str(self.converter.ffmpeg_path), '-hide_banner', '-y',
                             '-f', 'lavfi', '-i', 'sine=frequency=440:duration=0.1', str(ogg)])
        unusual_temp = self.base / '中文 临时目录'
        unusual_temp.mkdir()
        output = self.base / 'decoded.wav'
        with patch('tempfile.tempdir', str(unusual_temp)):
            self.converter.ogg_to_wav(ogg, output)
        self.assertTrue(output.read_bytes().startswith(b'RIFF'))
        self.assertEqual(list(unusual_temp.iterdir()), [])

    def test_real_wav_to_mp3(self):
        from audio_validation import AudioValidator
        self.worker.validator = AudioValidator(self.converter)
        if not self.converter.lame_ok:
            self.skipTest("bundled lame unavailable")
        item = self.item("中文 空格/song.wav", "standard_wav")
        with wave.open(str(item.path), "wb") as stream:
            stream.setparams((1, 2, 44100, 0, "NONE", "not compressed"))
            stream.writeframes(b"".join(struct.pack("<h", int(8000 * math.sin(2 * math.pi * 440 * i / 44100))) for i in range(11025)))
        stats = self.worker.process_batch([item], self.output, "mp3", "vbr")
        self.assertEqual(stats["done"], 1, item.error_msg)
        self.assertGreater(item.output_path.stat().st_size, 100)

    def test_progress_stream_handles_carriage_returns_and_errors(self):
        values = []
        self.converter._run([sys.executable, '-u', '-c',
                             "import sys; sys.stderr.write('song (99%).wav\\r[10%]\\r1/10 ( 25%)|\\rinput.flac: 60% complete\\r'); sys.stderr.flush()"],
                            on_progress=values.append)
        self.assertEqual(values, [10, 25, 60, 100])
        values = []
        with self.assertRaisesRegex(RuntimeError, 'stream error'):
            self.converter._run([sys.executable, '-u', '-c',
                                 "import sys; print('stream error'); sys.exit(2)"], on_progress=values.append)
        self.assertNotIn(100, values)
        with self.assertRaisesRegex(RuntimeError, '超时'):
            self.converter._run([sys.executable, '-c', 'import time; time.sleep(5)'],
                                timeout=0.1, on_progress=values.append)

    def test_m4a_conversion_keeps_tags_and_emits_stage_progress(self):
        from audio_validation import AudioValidator
        self.worker.validator = AudioValidator(self.converter)
        if not (self.converter.ffmpeg_ok and self.converter.lame_ok):
            self.skipTest('ffmpeg or lame unavailable')
        from mutagen.mp4 import MP4
        from mutagen.mp3 import MP3
        source = self.base / 'tagged.m4a'
        self.converter._run([str(self.converter.ffmpeg_path), '-hide_banner', '-y', '-f', 'lavfi',
                             '-i', 'sine=frequency=440:duration=1', '-c:a', 'aac', str(source)])
        audio = MP4(source)
        audio['\u00a9nam'] = ['M4A 中文歌名']
        audio['\u00a9ART'] = ['测试歌手']
        audio['trkn'] = [(2, 8)]
        audio.save()
        item = FileItem('m4a', source, source.name, source.stat().st_size, 'waiting', 'standard_m4a')
        stages = []
        stats = self.worker.process_batch([item], self.output, 'mp3', 'vbr',
                                          on_stage=lambda _, stage, percent: stages.append((stage, percent)))
        self.assertEqual(stats['done'], 1, item.error_msg)
        self.assertEqual(stats['warnings'], 0, item.warning_msg)
        result = MP3(item.output_path)
        self.assertEqual(str(result.tags['TIT2']), 'M4A 中文歌名')
        self.assertEqual(str(result.tags['TRCK']), '2/8')
        self.assertIn(('解码', 100), stages)
        self.assertIn(('编码', 100), stages)
        self.assertEqual(stages[-1][0], '保存结果')


if __name__ == "__main__":
    unittest.main()
