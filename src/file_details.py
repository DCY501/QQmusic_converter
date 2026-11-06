"""本地音频信息读取。加密源文件只展示文件信息，不自动连接客户端。"""

from models import TYPE_LABELS, detect_file_type


def describe_audio(path):
    try:
        if not path.is_file():
            return '文件不存在或无法访问'
        kind = detect_file_type(path)
        if kind.startswith('qmc_'):
            return '加密源文件，转换后可查看输出音频信息'
        import mutagen
        audio = mutagen.File(path)
        if audio is None:
            return '无法读取音频信息'
        info = audio.info
        seconds = getattr(info, 'length', 0)
        lines = [f'时长: {int(seconds // 60):02d}:{seconds % 60:04.1f}']
        for attr, label, unit in [('sample_rate', '采样率', ' Hz'), ('channels', '声道', ''),
                                  ('bitrate', '码率', ' bps')]:
            value = getattr(info, attr, None)
            if value:
                lines.append(f'{label}: {value}{unit}')
        tags = audio.tags or {}
        for label, keys in [('歌名', ('TIT2', 'title', '\u00a9nam')),
                            ('歌手', ('TPE1', 'artist', '\u00a9ART')),
                            ('专辑', ('TALB', 'album', '\u00a9alb')),
                            ('专辑歌手', ('TPE2', 'albumartist', 'aART')),
                            ('日期', ('TDRC', 'date', '\u00a9day')),
                            ('曲目', ('TRCK', 'tracknumber', 'trkn'))]:
            for key in keys:
                value = tags.get(key)
                if value:
                    value = ', '.join(map(str, value)) if isinstance(value, list) else str(value)
                    lines.append(f'{label}: {value[:1000]}')
                    break
        return '\n'.join(lines)
    except Exception as exc:
        return f'音频信息读取失败: {exc}'


def describe_item(item):
    labels = {'waiting': '等待中', 'converting': '处理中', 'done': '已完成',
              'error': '失败', 'skipped': '已跳过'}
    lines = [f'文件: {item.name}', f'类型: {TYPE_LABELS.get(item.file_type, item.file_type)}',
             f'状态: {labels.get(item.status, item.status)}', f'大小: {item.size_mb:.2f} MB',
             '', f'源路径: {item.path}', describe_audio(item.path)]
    if item.output_path:
        lines.extend(['', f'输出路径: {item.output_path}', describe_audio(item.output_path)])
    if item.error_msg:
        lines.extend(['', f'原因: {item.error_msg}'])
    if item.warning_msg:
        lines.extend(['', f'提示: {item.warning_msg}'])
    lines.extend(['', f'校验: {item.validation_msg or "暂无有效校验记录"}'])
    return '\n'.join(lines)
