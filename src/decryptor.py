#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""QQ 音乐 FRIDA 会话；异常保留原因，临时文件按调用隔离。"""

import shutil
import tempfile
import ctypes
from pathlib import Path

from models import DATA_DIR
from client_diagnostics import inspect_session, connection_error


class QQMusicDecryptor:
    def __init__(self):
        self.session = None
        self.script = None
        self._connected = False
        self.last_diagnostic = {}

    def connect(self):
        if self._connected:
            return
        self.last_diagnostic = {}
        try:
            import frida
            self.session = frida.attach("QQMusic.exe")
            self.last_diagnostic = inspect_session(self.session)
            if self.last_diagnostic['error']:
                raise RuntimeError(self.last_diagnostic['error'])
            js_code = (DATA_DIR / "hook_qqmusic_v17.js").read_text(encoding="utf-8")
            self.script = self.session.create_script(js_code)
            self.script.load()
            self._connected = True
        except Exception as exc:
            self.disconnect()
            message = self.last_diagnostic.get('error') or connection_error(exc)
            raise RuntimeError(message) from exc

    def decrypt(self, input_path: Path, output_path: Path):
        if not self._connected:
            raise RuntimeError("尚未连接 QQ 音乐")
        with tempfile.TemporaryDirectory(prefix="qqmusic_decrypt_") as folder:
            tmp_path = Path(folder) / "decoded.bin"
            # 老版 Frida File 写路径需 ASCII；中文用户名下优先使用 Windows 短路径。
            script_path = tmp_path
            if not folder.isascii():
                buffer = ctypes.create_unicode_buffer(32768)
                length = ctypes.windll.kernel32.GetShortPathNameW(str(folder), buffer, len(buffer))
                if not length or length >= len(buffer) or not buffer.value.isascii():
                    raise RuntimeError('临时目录含中文且无法获取短路径，请将 TEMP 指向可写的英文目录后重启程序')
                script_path = Path(buffer.value) / 'decoded.bin'
            try:
                result = self.script.exports_sync.decrypt(str(input_path.resolve()), str(script_path))
                if not result or not result.get("success"):
                    raise RuntimeError(f"解密脚本未返回成功结果: {result}")
                if not tmp_path.is_file() or tmp_path.stat().st_size != result.get("size"):
                    raise RuntimeError("解密输出大小与读取结果不一致")
                shutil.move(str(tmp_path), str(output_path))
            except Exception as exc:
                detail = str(exc)
                label = '文件无法解密，请确认此文件可由当前账号和客户端使用' if 'Open 失败' in detail else '解密失败'
                raise RuntimeError(f"{label}: {exc}") from exc

    def disconnect(self):
        if self.session:
            try:
                self.session.detach()
            except Exception:
                pass
        self.session = None
        self.script = None
        self._connected = False
