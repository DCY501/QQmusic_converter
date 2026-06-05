#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据模型、配置管理与文件类型检测
"""

import sys
import json
from pathlib import Path
from dataclasses import dataclass
from typing import Dict


def get_app_dir() -> Path:
    """获取程序所在目录（兼容开发环境和打包环境）"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


APP_DIR = get_app_dir()

# UI 显示用的类型名称映射
TYPE_LABELS = {
    'qmc_mflac': '加密FLAC',
    'qmc_mgg': '加密OGG',
    'qmc_mmp4': '加密MP4',
    'qmc_mflac_decrypted': '加密FLAC',
    'qmc_mgg_decrypted': '加密OGG',
    'qmc_mmp4_decrypted': '加密MP4',
    'standard_flac': 'FLAC',
    'standard_ogg': 'OGG',
    'standard_wav': 'WAV',
    'standard_mp3': 'MP3',
    'standard_m4a': 'M4A',
}


@dataclass
class FileItem:
    """文件项数据类"""
    id: str
    path: Path
    name: str
    size: int
    status: str  # 'waiting', 'converting', 'done', 'error', 'skipped'
    file_type: str  # 检测到的文件类型
    error_msg: str = ""

    @property
    def size_mb(self) -> float:
        return self.size / (1024 * 1024)


class Config:
    """配置管理"""
    CONFIG_FILE = APP_DIR / "config.json"

    @classmethod
    def load(cls) -> Dict:
        if cls.CONFIG_FILE.exists():
            try:
                with open(cls.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {}

    @classmethod
    def save(cls, config: Dict):
        try:
            with open(cls.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except:
            pass


def detect_file_type(file_path: Path) -> str:
    """检测文件类型"""
    ext = file_path.suffix.lower()
    try:
        with open(file_path, 'rb') as f:
            header = f.read(12)
    except:
        return 'error'

    # QQ音乐加密格式
    if ext in ['.mflac', '.mflac0']:
        return 'qmc_mflac'
    if ext in ['.mgg', '.mgg0', '.mgg1', '.mggl']:
        return 'qmc_mgg'
    if ext == '.mmp4':
        return 'qmc_mmp4'

    # 标准音频格式 (通过文件头确认)
    if ext == '.flac' and len(header) >= 4 and header[:4] == b'fLaC':
        return 'standard_flac'
    if ext == '.ogg' and len(header) >= 4 and header[:4] == b'OggS':
        return 'standard_ogg'
    if ext == '.wav' and len(header) >= 12 and header[:4] == b'RIFF' and header[8:12] == b'WAVE':
        return 'standard_wav'
    if ext in ['.mp3']:
        # MP3 文件头比较多样，通过后缀判断
        return 'standard_mp3'
    if ext in ['.m4a', '.mp4']:
        # MP4/M4A 格式
        if len(header) >= 8:
            # ftyp box
            if header[4:8] == b'ftyp':
                return 'standard_m4a'
        return 'standard_m4a'

    return 'unsupported'
