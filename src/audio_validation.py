"""发布输出前校验格式、音频参数与时长；可选完整解码。"""

import math


class AudioValidator:
    def __init__(self, converter):
        self.converter = converter

    @staticmethod
    def probe(path):
        import mutagen
        try:
            audio = mutagen.File(path)
            if audio is None:
                raise ValueError('无法识别音频格式')
            duration = float(audio.info.length)
            if not math.isfinite(duration) or duration <= 0:
                raise ValueError('音频时长无效')
            if getattr(audio.info, 'sample_rate', 0) <= 0 or getattr(audio.info, 'channels', 0) <= 0:
                raise ValueError('音频参数无效')
            return type(audio).__name__, duration
        except Exception as exc:
            raise RuntimeError(f'音频校验失败（{path.name}）: {exc}') from exc

    def validate(self, source, output, full=False):
        expected = {'.mp3': {'MP3'}, '.flac': {'FLAC'}, '.ogg': {'OggVorbis', 'OggOpus'},
                    '.wav': {'WAVE'}, '.m4a': {'MP4'}, '.mp4': {'MP4'}}
        kind, duration = self.probe(output)
        if kind not in expected.get(output.suffix.lower(), set()):
            raise RuntimeError(f'输出格式不符: {output.suffix} 文件实际识别为 {kind}')
        _, source_duration = self.probe(source)
        # 容许 MP3 帧边界/编码延迟造成的小量差异，不掩盖明显截断。
        tolerance = max(0.25, source_duration * 0.001)
        if abs(duration - source_duration) > tolerance:
            raise RuntimeError(f'输出时长异常: 源音频 {source_duration:.3f} 秒，输出 {duration:.3f} 秒')
        if full:
            if not self.converter.ffmpeg_ok:
                raise RuntimeError('完整解码校验需要 ffmpeg.exe，输出未提交')
            self.converter._run([str(self.converter.ffmpeg_path), '-nostdin', '-v', 'error',
                                 '-xerror', '-err_detect', 'explode', '-i', str(output),
                                 '-map', '0:a:0', '-f', 'null', '-'],
                                timeout=max(300, int(duration * 3)))
            return f'完整解码校验通过；{kind}，{duration:.2f} 秒'
        return f'基础校验通过（格式、参数、时长）；{kind}，{duration:.2f} 秒'
