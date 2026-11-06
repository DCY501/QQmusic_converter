# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files

project = Path(SPECPATH)
resources = [
    'hook_qqmusic_v17.js', 'flac.exe', 'oggdec.exe', 'lame.exe',
    'lame_enc.dll', 'libFLAC.dll', 'libFLAC++.dll',
]
datas = [(str(project / 'src' / name), 'src') for name in resources]
datas += collect_data_files('tkinterdnd2')

a = Analysis(
    [str(project / 'src' / 'qqmusic_converter_gui.py')],
    pathex=[str(project / 'src')], binaries=[], datas=datas,
    hiddenimports=['frida', 'tkinterdnd2', 'mutagen.flac', 'mutagen.oggvorbis',
                   'mutagen.wave', 'mutagen.mp3', 'mutagen.id3', 'mutagen.mp4'],
    hookspath=[], hooksconfig={}, runtime_hooks=[], excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name='QQMusicConverter', debug=False, bootloader_ignore_signals=False,
    strip=False, upx=False, console=False,
)
