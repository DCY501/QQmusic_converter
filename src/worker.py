#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""批处理、输出提交与错误记录，不依赖 GUI。"""

import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional

from models import USER_DIR, DECRYPTED_EXTENSIONS, FileItem, ProcessingResult
from metadata import copy_tags
from session_store import file_signature
from decryptor import QQMusicDecryptor
from converter import AudioConverter
from audio_validation import AudioValidator


class BatchConverter:
    def __init__(self, decryptor: QQMusicDecryptor, converter: AudioConverter,
                 log_path: Optional[Path] = None, validator=None):
        self.decryptor = decryptor
        self.converter = converter
        self.log_path = log_path if log_path is not None else USER_DIR / "error.log"
        self.validator = validator if validator is not None else AudioValidator(converter)

    @staticmethod
    def _publish(staged: Path, desired: Path) -> Path:
        """仅提交完整文件；并发实例也不能覆盖已有输出。"""
        index = 1
        while True:
            candidate = desired if index == 1 else desired.with_name(
                f"{desired.stem} ({index}){desired.suffix}")
            if candidate.exists():
                index += 1
                continue
            try:
                if os.name == "nt":
                    # Windows rename 在目标已存在时失败，不能用 replace。
                    staged.rename(candidate)
                else:
                    os.link(staged, candidate)
                    staged.unlink()
                return candidate
            except FileExistsError:
                index += 1

    def process_single(self, item: FileItem, output_dir: Path,
                       output_format: str, quality: str, on_stage=None, verify_full=False) -> ProcessingResult:
        def stage(label, percent=None):
            if on_stage:
                on_stage(label, percent)
        if output_format not in ("keep", "mp3"):
            raise ValueError(f"未知输出格式: {output_format}")
        if verify_full and not self.converter.ffmpeg_ok:
            raise RuntimeError('完整解码校验需要 ffmpeg.exe，请安装后重试，或取消此选项使用基础校验')
        if output_format == 'mp3':
            reason = self.converter.mp3_unavailable_reason(item.file_type)
            if reason:
                return ProcessingResult('skipped', message=reason)

        extensions = DECRYPTED_EXTENSIONS
        suffix = ".mp3" if output_format == "mp3" else extensions.get(
            item.file_type, item.path.suffix)
        desired = output_dir / (item.path.stem + suffix)
        if desired.resolve() == item.path.resolve():
            return ProcessingResult("skipped", message="源文件已在输出位置，无需复制")

        output_dir.mkdir(parents=True, exist_ok=True)
        # 与最终文件同一卷，确保提交是原子操作；失败只留下原文件。
        with tempfile.TemporaryDirectory(prefix=".qqmusic-", dir=output_dir) as folder:
            work_dir = Path(folder)
            current_path, current_type = item.path, item.file_type
            if item.file_type.startswith("qmc_"):
                stage('解密')
                current_path = work_dir / ("decoded" + extensions[item.file_type])
                self.decryptor.decrypt(item.path, current_path)
                current_type = item.file_type + "_decrypted"
            staged = work_dir / ("result" + suffix)
            if output_format == "keep":
                stage('复制')
                shutil.copy2(current_path, staged)
            else:
                self.converter.convert_to_mp3(current_path, staged, current_type, quality, on_stage=on_stage)
            if not staged.is_file() or staged.stat().st_size == 0:
                raise RuntimeError("处理未生成有效输出文件")
            warning = ""
            if output_format == "mp3" and current_type != "standard_mp3":
                stage('保留标签与封面')
                # 标签失败不能破坏已编码的音频；在独立副本写入后再替换。
                tagged = work_dir / "tagged.mp3"
                try:
                    shutil.copy2(staged, tagged)
                    copy_tags(current_path, tagged)
                    tagged.replace(staged)
                except Exception as exc:
                    warning = f"音频转换成功，但标签/封面未保留: {exc}"
            stage('完整解码校验' if verify_full else '基础校验')
            validation = self.validator.validate(current_path, staged, full=verify_full)
            stage('保存结果')
            return ProcessingResult("done", output_path=self._publish(staged, desired), warning=warning, validation=validation)

    def _log_error(self, item: FileItem) -> str:
        try:
            self.log_path.parent.mkdir(parents=True, exist_ok=True)
            with self.log_path.open("a", encoding="utf-8") as stream:
                stream.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {item.path}\n"
                             f"  类型: {item.file_type}\n  错误: {item.error_msg}\n\n")
            return ""
        except OSError as exc:
            return f"无法写入日志 {self.log_path}: {exc}"

    def process_batch(self, files: List[FileItem], output_dir: Path,
                      output_format: str, quality: str,
                      on_status: Optional[Callable] = None,
                      on_progress: Optional[Callable] = None,
                      on_error: Optional[Callable] = None,
                      stop_check: Optional[Callable] = None,
                      on_stage: Optional[Callable] = None, verify_full=False) -> Dict:
        stats = dict(done=0, error=0, skipped=0, warnings=0, stopped=False, remaining=0,
                     log_path=str(self.log_path), log_error="")
        connection_attempted = False
        connection_error = ""
        try:
            for index, item in enumerate(files):
                if stop_check and stop_check():
                    stats.update(stopped=True, remaining=len(files) - index)
                    break
                item.status, item.error_msg, item.output_path = "converting", "", None
                item.warning_msg = ""
                item.source_signature = file_signature(item.path)
                item.output_signature = None
                item.validation_msg = ''
                if on_status:
                    on_status(item, "处理中", "blue")
                try:
                    if not item.path.is_file():
                        raise FileNotFoundError(f'源文件不存在或无法访问: {item.path}')
                    needs_decrypt = item.file_type.startswith("qmc_") and not (
                        output_format == "mp3" and self.converter.mp3_unavailable_reason(item.file_type))
                    if needs_decrypt:
                        if not connection_attempted:
                            connection_attempted = True
                            if on_status:
                                on_status(None, "正在连接 QQ 音乐…", "blue")
                            if on_stage:
                                on_stage(item, '连接 QQ 音乐', None)
                            try:
                                self.decryptor.connect()
                            except Exception as exc:
                                connection_error = str(exc)
                        if connection_error:
                            raise RuntimeError(connection_error)
                    stage_callback = (lambda label, percent=None: on_stage(item, label, percent)) if on_stage else None
                    if stage_callback or verify_full:
                        result = self.process_single(item, output_dir, output_format, quality,
                                                     on_stage=stage_callback, verify_full=verify_full)
                    else:
                        result = self.process_single(item, output_dir, output_format, quality)
                except Exception as exc:
                    result = ProcessingResult("error", message=str(exc))

                item.status = result.status
                item.error_msg = result.message
                item.output_path = result.output_path
                item.output_signature = file_signature(result.output_path)
                item.warning_msg = result.warning
                item.validation_msg = result.validation
                if result.warning:
                    stats['warnings'] += 1
                stats[result.status] += 1
                if result.status == "error":
                    stats["log_error"] = self._log_error(item) or stats["log_error"]
                    if on_error:
                        on_error(item, result.message)
                if on_status:
                    label, color = {"done": ("已完成", "green"),
                                    "error": ("失败", "red"),
                                    "skipped": ("已跳过", "gray")}[result.status]
                    if result.warning:
                        label, color = "完成（有提示）", "orange"
                    on_status(item, label, color)
                if on_progress:
                    on_progress(index + 1, len(files))
        finally:
            if connection_attempted:
                self.decryptor.disconnect()
        return stats
