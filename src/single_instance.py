"""Windows 登录会话内单实例；内核对象随进程退出释放，不依赖残留锁文件。"""

import ctypes
import hashlib
import os
import time
from ctypes import wintypes

from models import USER_DIR


class SingleInstance:
    def __init__(self, name=None):
        if os.name != 'nt':
            raise RuntimeError('此桌面程序仅支持 Windows')
        suffix = hashlib.sha256(str(USER_DIR).lower().encode('utf-8')).hexdigest()[:20]
        self.name = name or 'Local\\QQMusicConverter-' + suffix
        self.default_name = name is None
        self.kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        self.kernel.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
        self.kernel.CreateMutexW.restype = wintypes.HANDLE
        self.kernel.CreateEventW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.BOOL, wintypes.LPCWSTR]
        self.kernel.CreateEventW.restype = wintypes.HANDLE
        self.kernel.OpenEventW.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.LPCWSTR]
        self.kernel.OpenEventW.restype = wintypes.HANDLE
        self.kernel.SetEvent.argtypes = [wintypes.HANDLE]
        self.kernel.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        self.kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        self.mutex = self.event = None

    def acquire(self):
        if self.mutex:
            return True
        ctypes.set_last_error(0)
        handle = self.kernel.CreateMutexW(None, False, self.name)
        if not handle:
            raise ctypes.WinError(ctypes.get_last_error())
        if ctypes.get_last_error() == 183:  # ERROR_ALREADY_EXISTS
            self.kernel.CloseHandle(handle)
            return False
        self.mutex = handle
        self.event = self.kernel.CreateEventW(None, False, False, self.name + '-activate')
        if not self.event:
            error = ctypes.WinError(ctypes.get_last_error())
            self.close()
            raise error
        return True

    def notify_existing(self, timeout=3):
        # 首个实例可能仍在创建窗口；事件会保留通知直到它开始轮询。
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            handle = self.kernel.OpenEventW(0x0002, False, self.name + '-activate')
            if handle:
                try:
                    if self.default_name:
                        user = ctypes.WinDLL('user32', use_last_error=True)
                        user.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
                        user.FindWindowW.restype = wintypes.HWND
                        user.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
                        window = user.FindWindowW(None, 'QQ音乐解密/转码工具')
                        if window:
                            pid = wintypes.DWORD()
                            user.GetWindowThreadProcessId(window, ctypes.byref(pid))
                            user.AllowSetForegroundWindow(pid.value)
                    return bool(self.kernel.SetEvent(handle))
                finally:
                    self.kernel.CloseHandle(handle)
            time.sleep(0.05)
        return False

    def activation_requested(self):
        return bool(self.event) and self.kernel.WaitForSingleObject(self.event, 0) == 0

    def close(self):
        for handle in (self.event, self.mutex):
            if handle:
                self.kernel.CloseHandle(handle)
        self.event = self.mutex = None
