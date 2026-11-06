"""一次遍历目录，逐项发送检测结果；不访问 Tk。"""

import os
from pathlib import Path

from models import EXTENSION_TYPES, detect_file_type


def scan_paths(paths, on_file, on_error, stop_check, excluded_dir=None):
    excluded_dir = Path(excluded_dir).resolve() if excluded_dir else None
    seen = set()

    def inspect(path):
        if stop_check() or path.suffix.lower() not in EXTENSION_TYPES:
            return
        try:
            path = path.resolve()
            key = os.path.normcase(str(path))
            if key in seen:
                return
            seen.add(key)
            kind = detect_file_type(path)
            if kind in ('error', 'unsupported'):
                on_error(f"无法读取或识别: {path}")
                return
            on_file(path, kind, path.stat().st_size)
        except OSError as exc:
            on_error(f"{path}: {exc}")

    for raw in paths:
        if stop_check():
            return
        path = Path(raw).resolve()
        if path.is_dir():
            for folder, dirs, names in os.walk(path, onerror=lambda exc: on_error(str(exc))):
                if stop_check():
                    return
                dirs[:] = [d for d in dirs if not d.startswith('.qqmusic-') and
                           (Path(folder) / d).resolve() != excluded_dir]
                for name in names:
                    if stop_check():
                        return
                    inspect(Path(folder) / name)
        else:
            inspect(path)
