"""转码后迁移常用文字标签和内嵌封面；不修改源文件。"""

import base64
from copy import deepcopy


def copy_tags(source, destination):
    import mutagen
    from mutagen.flac import Picture
    from mutagen.mp4 import MP4, MP4Cover
    from mutagen.id3 import (ID3, APIC, TIT2, TPE1, TPE2, TALB, TCOM,
                             TCON, TDRC, TRCK, TPOS, USLT, COMM)

    audio = mutagen.File(source)
    if audio is None:
        raise ValueError('无法识别源文件的标签格式')
    tags = ID3()
    if isinstance(audio, MP4):
        source_tags = audio.tags or {}
        fields = {'\u00a9nam': TIT2, '\u00a9ART': TPE1, 'aART': TPE2,
                  '\u00a9alb': TALB, '\u00a9wrt': TCOM, '\u00a9gen': TCON, '\u00a9day': TDRC}
        for key, frame in fields.items():
            if source_tags.get(key):
                tags.add(frame(encoding=3, text=source_tags[key]))
        for key, frame in [('trkn', TRCK), ('disk', TPOS)]:
            if source_tags.get(key):
                number, total = source_tags[key][0]
                tags.add(frame(encoding=3, text=[f'{number}/{total}' if total else str(number)]))
        for key, frame in [('\u00a9lyr', USLT), ('\u00a9cmt', COMM)]:
            if source_tags.get(key):
                tags.add(frame(encoding=3, lang='und', desc='', text='\n'.join(source_tags[key])))
        for index, cover in enumerate(source_tags.get('covr', [])):
            mime = 'image/png' if cover.imageformat == MP4Cover.FORMAT_PNG else 'image/jpeg'
            tags.add(APIC(encoding=3, mime=mime, type=3, desc=str(index), data=bytes(cover)))
    elif isinstance(audio.tags, ID3):
        for name in ('TIT2', 'TPE1', 'TPE2', 'TALB', 'TCOM', 'TCON', 'TDRC',
                     'TRCK', 'TPOS', 'USLT', 'COMM', 'APIC'):
            for frame in audio.tags.getall(name):
                tags.add(deepcopy(frame))
    else:
        fields = {'title': TIT2, 'artist': TPE1, 'albumartist': TPE2,
                  'album': TALB, 'composer': TCOM, 'genre': TCON,
                  'date': TDRC, 'tracknumber': TRCK, 'discnumber': TPOS}
        source_tags = audio.tags or {}
        for key, frame in fields.items():
            values = source_tags.get(key)
            if values:
                tags.add(frame(encoding=3, text=values))
        for key, frame in [('lyrics', USLT), ('comment', COMM)]:
            values = source_tags.get(key)
            if values:
                tags.add(frame(encoding=3, lang='und', desc='', text='\n'.join(values)))
        pictures = list(getattr(audio, 'pictures', []))
        for encoded in source_tags.get('metadata_block_picture', []):
            pictures.append(Picture(base64.b64decode(encoded)))
        for index, picture in enumerate(pictures):
            tags.add(APIC(encoding=3, mime=picture.mime, type=picture.type,
                          desc=f'{index}: {picture.desc}', data=picture.data))
    if tags:
        tags.save(destination, v2_version=3)
