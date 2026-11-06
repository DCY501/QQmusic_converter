import math
import struct
import sys
import tempfile
import unittest
import wave
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from audio_validation import AudioValidator
from converter import AudioConverter
from models import FileItem
from worker import BatchConverter


class AudioValidationTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix='qqmusic_validation_')
        self.addCleanup(temporary.cleanup)
        self.base = Path(temporary.name)
        self.converter = AudioConverter()
        self.validator = AudioValidator(self.converter)

    def wav(self, name, seconds=1):
        path = self.base / name
        with wave.open(str(path), 'wb') as stream:
            stream.setparams((1, 2, 44100, 0, 'NONE', 'not compressed'))
            stream.writeframes(b''.join(struct.pack('<h', int(7000 * math.sin(i / 20)))
                                       for i in range(int(seconds * 44100))))
        return path

    def test_rejects_invalid_format_and_wrong_duration(self):
        source = self.wav('source.wav', 2)
        bad = self.base / 'bad.mp3'
        bad.write_bytes(b'invalid audio')
        with self.assertRaisesRegex(RuntimeError, '校验失败'):
            self.validator.validate(source, bad)
        bad.write_bytes(source.read_bytes())
        with self.assertRaisesRegex(RuntimeError, '格式不符'):
            self.validator.validate(source, bad)
        short = self.wav('short.wav', 0.5)
        with self.assertRaisesRegex(RuntimeError, '时长异常'):
            self.validator.validate(source, short)

    def test_basic_and_full_validation_succeed_on_real_mp3(self):
        source = self.wav('source.wav')
        output = self.base / 'output.mp3'
        self.converter.wav_to_mp3(source, output, 'vbr')
        self.assertIn('基础校验通过', self.validator.validate(source, output))
        if self.converter.ffmpeg_ok:
            self.assertIn('完整解码校验通过', self.validator.validate(source, output, full=True))

    def test_full_validation_detects_damaged_flac_frames(self):
        if not (self.converter.ffmpeg_ok and self.converter.flac_ok):
            self.skipTest('ffmpeg/flac unavailable')
        source = self.wav('source.wav', 2)
        output = self.base / 'damaged.flac'
        self.converter._run([str(self.converter.flac_path), '-f', '-o', str(output), str(source)])
        data = bytearray(output.read_bytes())
        # 保留 FLAC 元数据，让基础检查仍能读取正确时长；破坏后面的音频帧。
        offset = 4
        while True:
            last = bool(data[offset] & 0x80)
            length = int.from_bytes(data[offset + 1:offset + 4], 'big')
            offset += 4 + length
            if last:
                break
        self.assertGreater(len(data) - offset, 100)
        data[offset + 20:offset + 80] = b'\xff' * 60
        output.write_bytes(data)
        self.assertIn('基础校验通过', self.validator.validate(source, output))
        with self.assertRaises(RuntimeError):
            self.validator.validate(source, output, full=True)

    def test_validation_failure_does_not_publish_or_overwrite(self):
        source = self.wav('source.wav')
        out = self.base / 'out'
        out.mkdir()
        existing = out / 'source.mp3'
        existing.write_bytes(b'previous file')
        item = FileItem('1', source, source.name, source.stat().st_size, 'waiting', 'standard_wav')
        worker = BatchConverter(None, self.converter, self.base / 'error.log')
        with patch.object(worker.validator, 'validate', side_effect=RuntimeError('校验失败: damaged')):
            stats = worker.process_batch([item], out, 'mp3', 'vbr')
        self.assertEqual(stats['error'], 1)
        self.assertEqual(item.validation_msg, '')
        self.assertEqual(existing.read_bytes(), b'previous file')
        self.assertEqual(list(out.iterdir()), [existing])

    def test_requesting_full_validation_without_ffmpeg_fails_explicitly(self):
        source = self.wav('source.wav')
        item = FileItem('1', source, source.name, source.stat().st_size, 'waiting', 'standard_wav')
        self.converter.ffmpeg_ok = False
        worker = BatchConverter(None, self.converter, self.base / 'error.log')
        stats = worker.process_batch([item], self.base / 'out', 'mp3', 'vbr', verify_full=True)
        self.assertEqual(stats['error'], 1)
        self.assertIn('ffmpeg.exe', item.error_msg)
        self.assertIsNone(item.output_path)


if __name__ == '__main__':
    unittest.main()
