"""供打包验证使用的自检；不连接 QQ 音乐，不读写用户配置。"""

import json
import tempfile
import wave
import uuid
from pathlib import Path


def run_self_check(report_path):
    report = {'ok': False}
    root = None
    try:
        import frida
        import mutagen
        from mutagen.flac import FLAC
        from mutagen.mp3 import MP3
        from mutagen.mp4 import MP4
        from gui import create_root
        from models import DATA_DIR, USER_DIR, Config, FileItem
        from converter import AudioConverter
        from worker import BatchConverter
        from session_store import SessionStore
        from file_details import describe_item
        from single_instance import SingleInstance

        root = create_root()
        root.withdraw()
        root.update_idletasks()
        report['drag_drop'] = root.dnd_available
        report['frida'] = frida.__version__
        report['mutagen'] = mutagen.version_string
        report['hook_script'] = (DATA_DIR / 'hook_qqmusic_v17.js').is_file()
        converter = AudioConverter()
        report['tools'] = dict(flac=converter.flac_ok, oggdec=converter.oggdec_ok,
                               lame=converter.lame_ok)
        name = 'Local\\QQMusicSelfCheck-' + uuid.uuid4().hex
        first, second = SingleInstance(name), SingleInstance(name)
        try:
            report['single_instance'] = (first.acquire() and not second.acquire() and
                                         second.notify_existing() and first.activation_requested())
        finally:
            first.close()
            second.close()
        # 在真实用户目录下使用独立临时文件，验证重定向环境中的首次保存与覆盖。
        USER_DIR.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix='storage-check-', dir=USER_DIR) as folder:
            class CheckConfig(Config):
                CONFIG_FILE = Path(folder) / 'config.json'

            check_session = SessionStore(Path(folder) / 'session.json')
            report['user_directory_storage'] = True
            for quality in ('vbr', '320k'):
                settings = {'quality': quality, 'output_format': 'mp3'}
                if not CheckConfig.save(settings):
                    raise OSError(CheckConfig.last_error)
                if not check_session.save([], settings):
                    raise OSError(check_session.last_error)
                report['user_directory_storage'] &= (
                    CheckConfig.load() == settings and check_session.load() == ([], settings))
        with tempfile.TemporaryDirectory(prefix='qqmusic_selfcheck_') as folder:
            work = Path(folder)
            source = work / 'input.wav'
            with wave.open(str(source), 'wb') as stream:
                stream.setparams((1, 2, 44100, 0, 'NONE', 'not compressed'))
                stream.writeframes(b'\x00\x00' * 11025)
            item = FileItem('test', source, source.name, source.stat().st_size,
                            'waiting', 'standard_wav')
            worker = BatchConverter(None, converter, work / 'error.log')
            stats = worker.process_batch([item], work / 'output', 'mp3', 'vbr')
            report['wav_to_mp3'] = item.status == 'done'
            report['basic_validation'] = item.validation_msg.startswith('基础校验通过')
            report['details'] = item.error_msg or item.warning_msg
            flac_path = work / 'tagged.flac'
            converter._run([str(converter.flac_path), '-f', '-o', str(flac_path), str(source)])
            tagged = FLAC(flac_path)
            tagged['title'] = '打包标签自检'
            tagged.save()
            flac_item = FileItem('flac', flac_path, flac_path.name, flac_path.stat().st_size,
                                 'waiting', 'standard_flac')
            flac_stats = worker.process_batch([flac_item], work / 'output', 'mp3', 'vbr')
            report['flac_to_mp3_with_tags'] = (
                flac_item.status == 'done' and not flac_stats['warnings'] and
                str(MP3(flac_item.output_path).tags.get('TIT2')) == '打包标签自检')
            session = SessionStore(work / 'session.json')
            session.save([item, flac_item], {'output_format': 'mp3', 'quality': 'vbr'})
            restored, _ = session.load()
            report['session_restore'] = len(restored) == 2 and all(row.status == 'done' for row in restored)
            report['file_details'] = '采样率:' in describe_item(flac_item)
            report['ffmpeg_available'] = converter.ffmpeg_ok
            if converter.ffmpeg_ok:
                m4a_path = work / 'tagged.m4a'
                converter._run([str(converter.ffmpeg_path), '-nostdin', '-hide_banner', '-y',
                                '-i', str(source), '-c:a', 'aac', str(m4a_path)])
                m4a = MP4(m4a_path)
                m4a['\u00a9nam'] = ['M4A 打包自检']
                m4a.save()
                m4a_item = FileItem('m4a', m4a_path, m4a_path.name, m4a_path.stat().st_size,
                                    'waiting', 'standard_m4a')
                stages = []
                m4a_stats = worker.process_batch([m4a_item], work / 'output', 'mp3', 'vbr',
                    on_stage=lambda _, stage, percent: stages.append((stage, percent)), verify_full=True)
                report['m4a_to_mp3_with_tags'] = (m4a_item.status == 'done' and not m4a_stats['warnings'] and
                    str(MP3(m4a_item.output_path).tags.get('TIT2')) == 'M4A 打包自检')
                report['stage_progress'] = ('解码', 100) in stages and ('编码', 100) in stages
                report['full_validation'] = m4a_item.validation_msg.startswith('完整解码校验通过')
            report['ok'] = (report['drag_drop'] and report['hook_script'] and
                            all(report['tools'].values()) and report['wav_to_mp3'] and
                            report['flac_to_mp3_with_tags'] and not stats['warnings'] and
                            report['session_restore'] and report['file_details'] and
                            report['single_instance'] and report['basic_validation'] and
                            report['user_directory_storage'] and
                            (not converter.ffmpeg_ok or (report['m4a_to_mp3_with_tags'] and report['stage_progress']
                                                        and report['full_validation'])))
    except Exception as exc:
        report['error'] = str(exc)
    finally:
        if root is not None:
            root.destroy()
    Path(report_path).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    return 0 if report['ok'] else 1
