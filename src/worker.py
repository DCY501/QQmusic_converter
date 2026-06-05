#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量转换工作器 - 纯逻辑，不依赖 GUI
"""

import shutil
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Callable

from models import FileItem, TYPE_LABELS
from decryptor import QQMusicDecryptor
from converter import AudioConverter


class BatchConverter:
    """批量音频转换工作器"""

    def __init__(self, decryptor: QQMusicDecryptor, converter: AudioConverter):
        self.decryptor = decryptor
        self.converter = converter

    def _determine_output_path(self, file_item: FileItem, output_dir: Path,
                                output_format: str) -> Tuple[Optional[Path], str]:
        """
        确定单个文件的输出路径。
        返回 (output_path, status_or_empty)，status 非空表示需要跳过/报错。
        """
        if output_format == 'keep':
            if file_item.file_type == 'qmc_mflac':
                ext = '.flac'
            elif file_item.file_type == 'qmc_mgg':
                ext = '.ogg'
            elif file_item.file_type == 'qmc_mmp4':
                ext = '.mp4'
            else:
                ext = file_item.path.suffix
            return output_dir / (file_item.path.stem + ext), ""
        else:
            # 统一转 MP3
            if file_item.file_type == 'qmc_mmp4':
                return None, "skipped"
            return output_dir / (file_item.path.stem + '.mp3'), ""

    def _decrypt_if_needed(self, file_item: FileItem, output_dir: Path) -> Tuple[Path, str, bool]:
        """
        如果是加密文件则解密。
        返回 (current_path, decrypted_type, success)
        """
        if not file_item.file_type.startswith('qmc_'):
            return file_item.path, file_item.file_type, True

        tmp_suffix = file_item.path.suffix.replace('.', '_')
        decrypted_path = output_dir / (file_item.path.stem + f"{tmp_suffix}.decrypted.tmp")

        if self.decryptor.decrypt(file_item.path, decrypted_path):
            if file_item.file_type == 'qmc_mflac':
                dec_type = 'qmc_mflac_decrypted'
            elif file_item.file_type == 'qmc_mgg':
                dec_type = 'qmc_mgg_decrypted'
            elif file_item.file_type == 'qmc_mmp4':
                dec_type = 'qmc_mmp4_decrypted'
            else:
                dec_type = file_item.file_type
            return decrypted_path, dec_type, True
        else:
            return file_item.path, file_item.file_type, False

    def _convert_or_copy(self, file_item: FileItem, current_path: Path, output_path: Path,
                         decrypted_type: str, output_format: str, quality: str) -> Tuple[bool, str]:
        """
        根据输出格式执行转换或复制。
        返回 (success, error_msg)
        """
        if output_format == 'keep':
            # 保持原格式
            if current_path != file_item.path:
                if output_path.exists():
                    output_path.unlink()
                shutil.move(str(current_path), str(output_path))
            else:
                shutil.copy2(str(current_path), str(output_path))
            return True, ""
        else:
            # 转 MP3
            if self.converter.convert_to_mp3(current_path, output_path, decrypted_type, quality):
                # 清理临时解密文件
                if current_path != file_item.path and current_path.exists():
                    try:
                        current_path.unlink()
                    except:
                        pass
                return True, ""
            else:
                return False, "转码失败"

    def process_single(self, file_item: FileItem, output_dir: Path,
                       output_format: str, quality: str) -> Tuple[bool, str]:
        """
        处理单个文件。
        返回 (success, error_msg)
        """
        # 1. 确定输出路径
        output_path, special_status = self._determine_output_path(
            file_item, output_dir, output_format
        )

        if special_status == "skipped":
            return False, ".mmp4 (E-AC3格式) 暂不支持转 MP3，请选择保持原格式"

        # 2. 解密（如果需要）
        current_path, decrypted_type, ok = self._decrypt_if_needed(file_item, output_dir)
        if not ok:
            return False, "FRIDA 解密失败"

        # 3. 转换或复制
        ok, err = self._convert_or_copy(
            file_item, current_path, output_path, decrypted_type, output_format, quality
        )
        if not ok:
            # 尝试清理临时解密文件
            if current_path != file_item.path and current_path.exists():
                try:
                    current_path.unlink()
                except:
                    pass
            return False, err

        return True, ""

    def process_batch(self, files: List[FileItem], output_dir: Path,
                      output_format: str, quality: str,
                      on_status: Optional[Callable] = None,
                      on_progress: Optional[Callable] = None,
                      on_error: Optional[Callable] = None,
                      stop_check: Optional[Callable] = None) -> Dict:
        """
        批量处理文件。

        回调签名：
            on_status(file_item: FileItem, status_text: str, color: str)
            on_progress(current_index: int, total: int)
            on_error(file_item: FileItem, error_msg: str)
            stop_check() -> bool

        返回统计：{'done': int, 'error': int, 'skipped': int}
        """
        total = len(files)
        stats = {'done': 0, 'error': 0, 'skipped': 0}

        # 如果有加密文件，先连接 FRIDA
        has_encrypted = any(f.file_type.startswith('qmc_') for f in files)
        if has_encrypted:
            if on_status:
                on_status(None, "正在连接 QQ 音乐...", "blue")
            if not self.decryptor.connect():
                # 全部标记为失败
                for f in files:
                    f.status = 'error'
                    f.error_msg = "无法 attach 到 QQMusic.exe"
                    if on_error:
                        on_error(f, f.error_msg)
                    if on_status:
                        on_status(f, '失败', 'red')
                    stats['error'] += 1
                return stats

        for i, file_item in enumerate(files):
            if stop_check and stop_check():
                break

            file_item.status = 'converting'
            if on_status:
                on_status(file_item, '处理中', 'blue')

            try:
                ok, err = self.process_single(file_item, output_dir, output_format, quality)
                if ok:
                    file_item.status = 'done'
                    if on_status:
                        on_status(file_item, '已完成', 'green')
                    stats['done'] += 1
                else:
                    if err == ".mmp4 (E-AC3格式) 暂不支持转 MP3，请选择保持原格式":
                        file_item.status = 'skipped'
                        file_item.error_msg = err
                        if on_status:
                            on_status(file_item, '不支持转MP3', 'gray')
                        stats['skipped'] += 1
                    else:
                        file_item.status = 'error'
                        file_item.error_msg = err
                        if on_error:
                            on_error(file_item, err)
                        if on_status:
                            on_status(file_item, '失败', 'red')
                        stats['error'] += 1
            except Exception as e:
                file_item.status = 'error'
                file_item.error_msg = str(e)
                if on_error:
                    on_error(file_item, str(e))
                if on_status:
                    on_status(file_item, '失败', 'red')
                stats['error'] += 1

            if on_progress:
                on_progress(i + 1, total)

        # 断开 FRIDA
        if has_encrypted:
            self.decryptor.disconnect()

        return stats
