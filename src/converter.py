#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""音频工具链。失败抛出含工具诊断的异常，由批处理逐项处理。"""

import subprocess
import tempfile
import shutil
import re
import threading
import time
from collections import deque
from queue import Queue, Empty
from pathlib import Path
from typing import List

from models import APP_DIR, DATA_DIR


class AudioConverter:
    def __init__(self):
        self.flac_path = DATA_DIR / "flac.exe"
        self.oggdec_path = DATA_DIR / "oggdec.exe"
        self.lame_path = DATA_DIR / "lame.exe"
        self._check_tools()

    def _check_tools(self):
        self.flac_ok = self.flac_path.is_file()
        self.oggdec_ok = self.oggdec_path.is_file()
        self.lame_ok = self.lame_path.is_file()
        candidates = [APP_DIR / "ffmpeg.exe", DATA_DIR / "ffmpeg.exe",
                      DATA_DIR.parent / "ffmpeg.exe"]
        self.ffmpeg_path = next((p for p in candidates if p.is_file()), candidates[0])
        self.ffmpeg_ok = self.ffmpeg_path.is_file()

    def _run(self, cmd: List[str], timeout: int = 300, cwd=None, on_progress=None):
        tool = Path(cmd[0]).name
        if on_progress is not None:
            return self._run_with_progress(cmd, timeout, cwd, on_progress)
        try:
            result = subprocess.run(
                cmd, capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=timeout, cwd=cwd,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(f"{tool} 超时（{timeout} 秒），已结束该工具进程") from exc
        except OSError as exc:
            raise RuntimeError(f"无法运行 {tool}: {exc}") from exc
        if result.returncode:
            detail = (result.stderr or result.stdout or "未返回诊断信息").strip()
            raise RuntimeError(f"{tool} 失败（退出码 {result.returncode}）:\n{detail[-6000:]}")

    def _run_with_progress(self, cmd, timeout, cwd, on_progress):
        tool = Path(cmd[0]).name
        try:
            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                       text=True, encoding='utf-8', errors='replace', cwd=cwd,
                                       creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        except OSError as exc:
            raise RuntimeError(f'无法运行 {tool}: {exc}') from exc
        events, diagnostic = Queue(), deque(maxlen=100)
        def read_lines():
            try:
                for line in process.stdout:
                    events.put(line)
            finally:
                events.put(None)
        reader = threading.Thread(target=read_lines, daemon=True)
        reader.start()
        started, duration, last_percent = time.monotonic(), None, -1
        eof = False
        try:
            while not eof or process.poll() is None:
                if time.monotonic() - started > timeout:
                    raise RuntimeError(f'{tool} 超时（{timeout} 秒），已结束该工具进程')
                try:
                    line = events.get(timeout=0.1)
                except Empty:
                    continue
                if line is None:
                    eof = True
                    continue
                diagnostic.append(line.rstrip())
                match = re.search(r'^\s*Duration: (\d+):(\d+):(\d+(?:\.\d+)?)', line)
                if match:
                    duration = int(match[1]) * 3600 + int(match[2]) * 60 + float(match[3])
                percent = None
                if line.startswith('out_time_us=') and duration:
                    try:
                        percent = float(line.split('=', 1)[1]) / 10000 / duration
                    except ValueError:
                        pass
                elif tool.lower() != 'ffmpeg.exe':
                    # 只解析工具的进度行，避免文件名或元数据中的百分号被当成进度。
                    match = re.search(r':\s*(\d+(?:\.\d+)?)%\s*complete\s*$'
                                      r'|^\s*\[\s*(\d+(?:\.\d+)?)%\s*\]'
                                      r'|^\s*\d+/\d+\s*\(\s*(\d+(?:\.\d+)?)%\)', line)
                    if match:
                        percent = float(next(value for value in match.groups() if value is not None))
                if percent is not None:
                    percent = max(0, min(99, int(percent)))
                    if percent > last_percent:
                        on_progress(percent)
                        last_percent = percent
            if process.returncode:
                detail = '\n'.join(diagnostic)[-6000:]
                raise RuntimeError(f'{tool} 失败（退出码 {process.returncode}）:\n{detail}')
            on_progress(100)
        finally:
            if process.poll() is None:
                process.kill()
            process.wait()
            reader.join()
            process.stdout.close()

    def mp3_unavailable_reason(self, file_type):
        if file_type in ('qmc_mmp4', 'qmc_mmp4_decrypted', 'standard_m4a') and not self.ffmpeg_ok:
            return '此格式转 MP3 需要 ffmpeg.exe，请将它放到程序同级目录后重新处理'
        return ''

    def media_to_wav(self, input_path, output_path, on_progress=None):
        if not self.ffmpeg_ok:
            raise RuntimeError('缺少 ffmpeg.exe，无法解码 MP4/M4A')
        self._run([str(self.ffmpeg_path), '-nostdin', '-hide_banner', '-y',
                   '-i', str(input_path), '-map', '0:a:0', '-vn', '-ac', '2',
                   '-c:a', 'pcm_s16le', '-progress', 'pipe:1', '-nostats', str(output_path)],
                  on_progress=on_progress)
        self._verify_output(output_path, 'FFmpeg 解码')

    @staticmethod
    def _verify_output(path: Path, stage: str):
        if not path.is_file() or path.stat().st_size == 0:
            raise RuntimeError(f"{stage} 未生成有效输出")

    def flac_to_wav(self, input_path: Path, output_path: Path, on_progress=None):
        errors = []
        if self.flac_ok:
            try:
                self._run([str(self.flac_path), "-d", "-f", "-o", str(output_path), str(input_path)], on_progress=on_progress)
                self._verify_output(output_path, "FLAC 解码")
                return
            except RuntimeError as exc:
                errors.append(str(exc))
        if self.ffmpeg_ok:
            try:
                self._run([str(self.ffmpeg_path), "-nostdin", "-hide_banner", "-y",
                           "-i", str(input_path), "-vn", '-progress', 'pipe:1', '-nostats', str(output_path)], on_progress=on_progress)
                self._verify_output(output_path, "FFmpeg 解码")
                return
            except RuntimeError as exc:
                errors.append(str(exc))
        raise RuntimeError("\n\n".join(errors) or "缺少 FLAC 解码工具：flac.exe 或 ffmpeg.exe")

    def ogg_to_wav(self, input_path: Path, output_path: Path, on_progress=None):
        if not self.oggdec_ok:
            raise RuntimeError("缺少 OGG 解码工具：oggdec.exe")
        with tempfile.TemporaryDirectory(prefix="qqmusic_ogg_") as folder:
            tmp_dir = Path(folder)
            shutil.copy2(input_path, tmp_dir / "input.ogg")
            # 给旧版工具传相对 ASCII 文件名，工作目录可以包含中文或空格。
            self._run([str(self.oggdec_path.resolve()), "input.ogg"], cwd=tmp_dir, on_progress=on_progress)
            self._verify_output(tmp_dir / "input.wav", "OGG 解码")
            shutil.move(str(tmp_dir / "input.wav"), str(output_path))

    def wav_to_mp3(self, input_path: Path, output_path: Path, quality: str, on_progress=None):
        if not self.lame_ok:
            raise RuntimeError("缺少 MP3 编码工具：lame.exe")
        options = {"vbr": ["-V", "0"], "320k": ["-b", "320"], "192k": ["-b", "192"]}
        if quality not in options:
            raise ValueError(f"未知 MP3 音质: {quality}")
        self._run([str(self.lame_path), *options[quality], str(input_path), str(output_path)], on_progress=on_progress)
        self._verify_output(output_path, "MP3 编码")

    def convert_to_mp3(self, input_path: Path, output_path: Path, file_type: str, quality: str, on_stage=None):
        def stage(label, percent=None):
            if on_stage:
                on_stage(label, percent)
        if file_type == "standard_mp3":
            stage('复制')
            shutil.copy2(input_path, output_path)
            return
        with tempfile.TemporaryDirectory(prefix="qqmusic_conv_") as folder:
            wav = Path(folder) / "decoded.wav"
            stage('解码')
            progress = (lambda percent: stage('解码', percent)) if on_stage else None
            if file_type in ("standard_flac", "qmc_mflac_decrypted"):
                self.flac_to_wav(input_path, wav, progress)
            elif file_type in ("standard_ogg", "qmc_mgg_decrypted"):
                self.ogg_to_wav(input_path, wav, progress)
            elif file_type in ('standard_m4a', 'qmc_mmp4_decrypted'):
                self.media_to_wav(input_path, wav, progress)
            elif file_type == "standard_wav":
                wav = input_path
            else:
                raise RuntimeError(f"不支持将 {file_type} 转为 MP3")
            stage('编码')
            progress = (lambda percent: stage('编码', percent)) if on_stage else None
            self.wav_to_mp3(wav, output_path, quality, progress)
