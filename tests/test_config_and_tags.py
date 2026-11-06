import base64
import sys
import tempfile
import unittest
import wave
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from converter import AudioConverter
from models import Config, FileItem
from scanner import scan_paths
from worker import BatchConverter


class ConfigAndTagsTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='qqmusic_config_')
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)

    def test_invalid_config_and_atomic_save_failure(self):
        path = self.base / 'config.json'
        with patch.object(Config, 'CONFIG_FILE', path):
            path.write_text('[]', encoding='utf-8')
            self.assertEqual(Config.load(), {})
            self.assertTrue(Config.last_error)
            path.write_text('{"quality": "bad", "output_format": "keep"}', encoding='utf-8')
            self.assertEqual(Config.load(), {'output_format': 'keep'})
            old = path.read_bytes()
            with patch('models.os.replace', side_effect=PermissionError('locked')):
                self.assertFalse(Config.save({'quality': 'vbr'}))
            self.assertEqual(path.read_bytes(), old)
            self.assertEqual(list(self.base.iterdir()), [path])
            self.assertTrue(Config.save({'quality': 'vbr'}))
            self.assertEqual(Config.load(), {'quality': 'vbr'})

    def test_scanner_excludes_output_and_work_directories(self):
        output = self.base / 'out'
        output.mkdir()
        (output / 'copied.mp3').write_bytes(b'audio')
        work = self.base / '.qqmusic-work'
        work.mkdir()
        (work / 'partial.mp3').write_bytes(b'partial')
        song = self.base / 'song.mp3'
        song.write_bytes(b'audio')
        found, errors = [], []
        scan_paths([self.base], lambda *args: found.append(args), errors.append, lambda: False, output)
        self.assertEqual([row[0] for row in found], [song.resolve()])
        self.assertEqual(errors, [])

    def test_ffmpeg_next_to_executable_takes_priority(self):
        app, data = self.base / 'app', self.base / 'bundle' / 'src'
        app.mkdir()
        data.mkdir(parents=True)
        tool = app / 'ffmpeg.exe'
        tool.write_bytes(b'placeholder')
        with patch('converter.APP_DIR', app), patch('converter.DATA_DIR', data):
            converter = AudioConverter()
        self.assertTrue(converter.ffmpeg_ok)
        self.assertEqual(converter.ffmpeg_path, tool)

    def test_real_flac_tags_and_cover_survive_conversion(self):
        try:
            from mutagen.flac import FLAC, Picture
            from mutagen.mp3 import MP3
        except ImportError:
            self.skipTest('mutagen is not installed')
        converter = AudioConverter()
        if not (converter.flac_ok and converter.lame_ok):
            self.skipTest('audio tools unavailable')
        wav = self.base / 'source.wav'
        with wave.open(str(wav), 'wb') as stream:
            stream.setparams((1, 2, 44100, 0, 'NONE', 'not compressed'))
            stream.writeframes(b'\x00\x00' * 11025)
        flac = self.base / '歌曲.flac'
        converter._run([str(converter.flac_path), '-f', '-o', str(flac), str(wav)])
        source = FLAC(flac)
        source['title'] = '测试歌曲'
        source['artist'] = ['歌手甲', '歌手乙']
        source['album'] = '测试专辑'
        picture = Picture()
        picture.type, picture.mime = 3, 'image/png'
        picture.data = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j5V0AAAAASUVORK5CYII=')
        source.add_picture(picture)
        source.save()
        item = FileItem('1', flac, flac.name, flac.stat().st_size, 'waiting', 'standard_flac')
        worker = BatchConverter(None, converter, self.base / 'error.log')
        stats = worker.process_batch([item], self.base / 'out', 'mp3', 'vbr')
        self.assertEqual(stats['done'], 1, item.error_msg)
        self.assertEqual(stats['warnings'], 0, item.warning_msg)
        result = MP3(item.output_path)
        self.assertEqual(str(result.tags['TIT2']), '测试歌曲')
        self.assertEqual(str(result.tags['TALB']), '测试专辑')
        self.assertEqual(result.tags.getall('APIC')[0].data, picture.data)


if __name__ == '__main__':
    unittest.main()
