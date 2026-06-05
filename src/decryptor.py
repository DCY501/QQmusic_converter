#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QQ音乐 FRIDA Hook 解密器
"""

import tempfile
import hashlib
import shutil
from pathlib import Path

from models import APP_DIR


class QQMusicDecryptor:
    """QQ音乐 FRIDA hook 解密器"""

    def __init__(self):
        self.session = None
        self.script = None
        self._connected = False

    def is_qqmusic_running(self) -> bool:
        """检查 QQ 音乐是否在运行"""
        try:
            import frida
            frida.attach("QQMusic.exe")
            return True
        except:
            return False

    def connect(self) -> bool:
        """连接 QQ 音乐进程"""
        if self._connected:
            return True
        try:
            import frida
            self.session = frida.attach("QQMusic.exe")
            js_path = APP_DIR / "hook_qqmusic_v17.js"
            with open(js_path, "r", encoding="utf-8") as f:
                js_code = f.read()
            self.script = self.session.create_script(js_code)
            self.script.load()
            self._connected = True
            return True
        except Exception as e:
            print(f"[FRIDA] 连接失败: {e}")
            return False

    def decrypt(self, input_path: Path, output_path: Path) -> bool:
        """解密单个文件"""
        if not self._connected:
            raise RuntimeError("未连接 QQ 音乐")

        # 必须使用绝对路径，因为 FRIDA hook 在 QQ 音乐进程内执行，
        # 其工作目录与 Python 脚本不同，相对路径会解析失败
        input_path = input_path.resolve()

        # FRIDA File API 不支持中文路径，用纯 ASCII 临时文件
        tmp_dir = Path(tempfile.gettempdir()) / "qqmusic_decrypt"
        tmp_dir.mkdir(exist_ok=True)
        tmp_name = hashlib.md5(str(input_path).encode()).hexdigest()
        tmp_path = tmp_dir / tmp_name

        try:
            result = self.script.exports_sync.decrypt(str(input_path), str(tmp_path))
            if result and result.get("success"):
                if output_path.exists():
                    output_path.unlink()
                shutil.move(str(tmp_path), str(output_path))
                return True
            return False
        except Exception as e:
            print(f"[FRIDA] 解密出错: {e}")
            # 清理临时文件
            if tmp_path.exists():
                tmp_path.unlink()
            return False

    def disconnect(self):
        """断开连接"""
        if self.session:
            try:
                self.session.detach()
            except:
                pass
        self.session = None
        self.script = None
        self._connected = False
