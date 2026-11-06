#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据模型、配置管理与文件类型检测
"""

import sys
import json
import os
import tempfile
from pathlib import Path
from dataclasses import dataclass
from typing import Dict, Optional


def get_app_dir() -> Path:
    """获取应用目录（EXE 所在目录或源码目录）。"""
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后：exe 所在目录（如桌面、dist/）
        return Path(sys.executable).parent
    # 开发环境：src/ 目录
    return Path(__file__).parent


def get_data_dir() -> Path:
    """获取程序资源目录（工具、脚本从这里读取）"""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # PyInstaller onefile 模式：资源在临时解压目录的 src/ 子目录下
        return Path(sys._MEIPASS) / "src"
    # 开发环境 或 onedir 模式
    return Path(__file__).parent


APP_DIR = get_app_dir()
DATA_DIR = get_data_dir()
USER_DIR = Path(os.environ.get('LOCALAPPDATA', str(Path.home() / 'AppData' / 'Local'))) / 'QQMusicConverter'

EXTENSION_TYPES = {
    '.mflac': 'qmc_mflac', '.mflac0': 'qmc_mflac',
    '.mgg': 'qmc_mgg', '.mgg0': 'qmc_mgg', '.mgg1': 'qmc_mgg', '.mggl': 'qmc_mgg',
    '.mmp4': 'qmc_mmp4', '.flac': 'standard_flac', '.ogg': 'standard_ogg',
    '.wav': 'standard_wav', '.mp3': 'standard_mp3',
    '.m4a': 'standard_m4a', '.mp4': 'standard_m4a',
}
DECRYPTED_EXTENSIONS = {'qmc_mflac': '.flac', 'qmc_mgg': '.ogg', 'qmc_mmp4': '.mp4'}

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
    output_path: Optional[Path] = None
    warning_msg: str = ""
    source_signature: Optional[Dict] = None
    output_signature: Optional[Dict] = None
    validation_msg: str = ""

    @property
    def size_mb(self) -> float:
        return self.size / (1024 * 1024)


@dataclass(frozen=True)
class ProcessingResult:
    """单个文件的结果；状态与面向用户的说明分开。"""
    status: str
    output_path: Optional[Path] = None
    message: str = ""
    warning: str = ""
    validation: str = ""


class Config:
    """配置管理"""
    CONFIG_FILE = USER_DIR / "config.json"
    LEGACY_FILE = APP_DIR / "config.json"
    last_error = ""

    @classmethod
    def load(cls) -> Dict:
        cls.last_error = ""
        path = cls.CONFIG_FILE if cls.CONFIG_FILE.exists() else cls.LEGACY_FILE
        try:
            config = json.loads(path.read_text(encoding='utf-8-sig'))
            if not isinstance(config, dict):
                raise ValueError('配置内容必须是对象')
            clean = {}
            output = config.get('output_path')
            if isinstance(output, str) and output.strip():
                output_path = Path(output).expanduser()
                if not output_path.is_absolute():
                    output_path = APP_DIR / output_path
                clean['output_path'] = str(output_path.resolve())
            if config.get('output_format') in ('keep', 'mp3'):
                clean['output_format'] = config['output_format']
            if config.get('quality') in ('vbr', '320k', '192k'):
                clean['quality'] = config['quality']
            if isinstance(config.get('verify_full'), bool):
                clean['verify_full'] = config['verify_full']
            return clean
        except FileNotFoundError:
            return {}
        except (OSError, ValueError) as exc:
            cls.last_error = f'配置读取失败，已使用默认设置: {exc}'
            return {}

    @classmethod
    def save(cls, config: Dict):
        cls.last_error = ""
        tmp_path = None
        try:
            cls.CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
            # Windows 应用目录虚拟化可能把 LocalAppData 重定向到另一磁盘。
            # 先解析已存在的父目录，让临时文件和目标使用同一个实际路径。
            target = cls.CONFIG_FILE.parent.resolve() / cls.CONFIG_FILE.name
            with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8',
                                             dir=target.parent, delete=False) as f:
                tmp_path = Path(f.name)
                json.dump(config, f, ensure_ascii=False, indent=2)
            os.replace(tmp_path, target)
            return True
        except (OSError, ValueError, TypeError) as exc:
            cls.last_error = f'配置保存失败: {exc}'
            return False
        finally:
            if tmp_path and tmp_path.exists():
                try:
                    tmp_path.unlink()
                except OSError:
                    pass


def detect_file_type(file_path: Path) -> str:
    """检测文件类型"""
    ext = file_path.suffix.lower()
    file_type = EXTENSION_TYPES.get(ext)
    if file_type is None:
        return 'unsupported'
    try:
        with open(file_path, 'rb') as f:
            header = f.read(12)
    except OSError:
        return 'error'

    # QQ音乐加密格式
    if file_type.startswith('qmc_'):
        return file_type

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
        return 'unsupported'

    return 'unsupported'
