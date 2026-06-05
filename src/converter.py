#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
音频转换器 - 封装小工具链
"""

import subprocess
import tempfile
import shutil
from pathlib import Path
from typing import List, Tuple

from models import DATA_DIR


class AudioConverter:
    """音频转换器 - 封装小工具链"""

    def __init__(self):
        self.flac_path = DATA_DIR / "flac.exe"
        self.oggdec_path = DATA_DIR / "oggdec.exe"
        self.lame_path = DATA_DIR / "lame.exe"
        # ffmpeg 作为备选解码器（处理 flac.exe 无法解码的损坏/特殊 FLAC）
        self.ffmpeg_path = DATA_DIR.parent / "ffmpeg.exe"
        self._check_tools()

    def _check_tools(self):
        """检查工具是否存在"""
        self.flac_ok = self.flac_path.exists()
        self.oggdec_ok = self.oggdec_path.exists()
        self.lame_ok = self.lame_path.exists()
        self.ffmpeg_ok = self.ffmpeg_path.exists()

    def _run(self, cmd: List[str], timeout: int = 300) -> Tuple[bool, str]:
        """运行命令"""
        try:
            result = subprocess.run(
                cmd, capture_output=True, text=True,
                encoding='utf-8', errors='ignore',
                timeout=timeout, creationflags=subprocess.CREATE_NO_WINDOW
            )
            return result.returncode == 0, result.stderr or result.stdout
        except Exception as e:
            return False, str(e)

    def flac_to_wav(self, input_path: Path, output_path: Path) -> bool:
        """FLAC → WAV（flac.exe 失败时尝试 ffmpeg 备选）"""
        if not self.flac_ok:
            return False
        ok, _ = self._run([str(self.flac_path), "-d", "-f", "-o", str(output_path), str(input_path)])
        if ok and output_path.exists():
            return True
        # flac.exe 失败时，尝试用 ffmpeg 作为备选（更容错，可跳过损坏帧）
        if self.ffmpeg_ok:
            ok, _ = self._run([
                str(self.ffmpeg_path), "-hide_banner", "-y",
                "-i", str(input_path),
                "-ar", "44100", "-ac", "2",
                str(output_path)
            ])
            return ok and output_path.exists()
        return False

    def ogg_to_wav(self, input_path: Path, output_path: Path) -> bool:
        """OGG → WAV (oggdec 不支持空格路径，需临时复制)"""
        if not self.oggdec_ok:
            return False

        # oggdec 不支持空格路径，临时复制到纯 ASCII 路径
        tmp_dir = Path(tempfile.gettempdir()) / "qqmusic_ogg"
        tmp_dir.mkdir(exist_ok=True)
        tmp_ogg = tmp_dir / "tmp.ogg"
        tmp_wav = tmp_dir / "tmp.wav"

        try:
            shutil.copy2(str(input_path), str(tmp_ogg))
            ok, _ = self._run([str(self.oggdec_path), str(tmp_ogg)])
            if ok and tmp_wav.exists():
                if output_path.exists():
                    output_path.unlink()
                shutil.move(str(tmp_wav), str(output_path))
                return True
            return False
        finally:
            # 清理临时文件
            for f in [tmp_ogg, tmp_wav]:
                if f.exists():
                    try:
                        f.unlink()
                    except:
                        pass

    def wav_to_mp3(self, input_path: Path, output_path: Path, quality: str) -> bool:
        """WAV → MP3"""
        if not self.lame_ok:
            return False

        if quality == 'vbr':
            cmd = [str(self.lame_path), '-V', '0', str(input_path), str(output_path)]
        elif quality == '320k':
            cmd = [str(self.lame_path), '-b', '320', str(input_path), str(output_path)]
        else:  # 192k
            cmd = [str(self.lame_path), '-b', '192', str(input_path), str(output_path)]

        ok, _ = self._run(cmd)
        return ok and output_path.exists()

    def convert_to_mp3(self, input_path: Path, output_path: Path, file_type: str, quality: str) -> bool:
        """通用转 MP3 入口"""
        tmp_dir = Path(tempfile.gettempdir()) / "qqmusic_conv"
        tmp_dir.mkdir(exist_ok=True)
        tmp_wav = tmp_dir / f"{input_path.stem}.wav"

        try:
            # Step 1: 解码为 WAV
            if file_type in ['standard_flac', 'qmc_mflac_decrypted']:
                if not self.flac_to_wav(input_path, tmp_wav):
                    return False
            elif file_type in ['standard_ogg', 'qmc_mgg_decrypted']:
                if not self.ogg_to_wav(input_path, tmp_wav):
                    return False
            elif file_type == 'standard_wav':
                tmp_wav = input_path  # 直接就是 WAV
            elif file_type == 'standard_mp3':
                # MP3 → MP3，直接复制
                shutil.copy2(str(input_path), str(output_path))
                return True
            else:
                return False

            # Step 2: WAV → MP3
            if file_type == 'standard_wav':
                return self.wav_to_mp3(input_path, output_path, quality)
            else:
                ok = self.wav_to_mp3(tmp_wav, output_path, quality)
                return ok
        finally:
            # 清理临时 WAV
            if tmp_wav != input_path and tmp_wav.exists():
                try:
                    tmp_wav.unlink()
                except:
                    pass
