"""保存最近一次窗口的任务列表；恢复时验证文件，绝不自动开始转换。"""

import json
import os
import tempfile
from pathlib import Path

from models import USER_DIR, FileItem, detect_file_type


def file_signature(path):
    if path is None:
        return None
    try:
        stat = Path(path).stat()
        if not Path(path).is_file():
            return None
        return {'size': stat.st_size, 'mtime_ns': stat.st_mtime_ns}
    except OSError:
        return None


class SessionStore:
    def __init__(self, path=None):
        self.path = Path(path) if path is not None else USER_DIR / 'session.json'
        self.last_error = ''

    def save(self, items, settings):
        records = []
        for item in items:
            records.append(dict(path=str(item.path), size=item.size, file_type=item.file_type,
                                status=item.status, error_msg=item.error_msg,
                                warning_msg=item.warning_msg,
                                validation_msg=item.validation_msg,
                                output_path=str(item.output_path) if item.output_path else None,
                                source_signature=item.source_signature,
                                output_signature=item.output_signature))
        temporary = None
        self.last_error = ''
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            # 使用实际目录，避免 Windows 虚拟化路径触发跨盘替换错误。
            target = self.path.parent.resolve() / self.path.name
            with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8',
                                             dir=target.parent, delete=False) as stream:
                temporary = Path(stream.name)
                json.dump(dict(version=1, settings=settings, files=records), stream,
                          ensure_ascii=False, indent=2)
            os.replace(temporary, target)
            return True
        except (OSError, ValueError, TypeError) as exc:
            self.last_error = f'任务记录保存失败: {exc}'
            return False
        finally:
            if temporary and temporary.exists():
                try:
                    temporary.unlink()
                except OSError:
                    pass

    def load(self):
        self.last_error = ''
        try:
            data = json.loads(self.path.read_text(encoding='utf-8'))
            if not isinstance(data, dict) or data.get('version') != 1 or not isinstance(data.get('files'), list):
                raise ValueError('不支持的任务记录格式')
            settings = data.get('settings', {})
            if not isinstance(settings, dict):
                settings = {}
            clean = {}
            if isinstance(settings.get('output_path'), str) and settings['output_path'].strip():
                clean['output_path'] = str(Path(settings['output_path']).resolve())
            if settings.get('output_format') in ('keep', 'mp3'):
                clean['output_format'] = settings['output_format']
            if settings.get('quality') in ('vbr', '320k', '192k'):
                clean['quality'] = settings['quality']
            if isinstance(settings.get('verify_full'), bool):
                clean['verify_full'] = settings['verify_full']
            items, seen = [], set()
            for record in data['files']:
                try:
                    item = self._restore_item(record, str(len(items)))
                    key = os.path.normcase(str(item.path))
                    if key not in seen:
                        items.append(item)
                        seen.add(key)
                except (KeyError, TypeError, ValueError, OSError):
                    self.last_error = '部分损坏的任务记录已忽略'
            return items, clean
        except FileNotFoundError:
            return [], {}
        except (OSError, ValueError, TypeError) as exc:
            self.last_error = f'无法恢复上次任务: {exc}'
            return [], {}

    @staticmethod
    def _restore_item(record, item_id):
        path = Path(record['path']).resolve()
        kind = record['file_type']
        if not isinstance(kind, str):
            raise ValueError('无效类型')
        output = Path(record['output_path']).resolve() if record.get('output_path') else None
        item = FileItem(item_id, path, path.name, int(record['size']),
                        record.get('status', 'waiting'), kind,
                        error_msg=str(record.get('error_msg', '')),
                        warning_msg=str(record.get('warning_msg', '')),
                        validation_msg=str(record.get('validation_msg', '')),
                        output_path=output, source_signature=record.get('source_signature'),
                        output_signature=record.get('output_signature'))
        current = file_signature(path)
        if current is None:
            item.status, item.error_msg = 'error', '源文件不存在或无法访问，请重新添加文件'
            return item
        item.size = current['size']
        if current != item.source_signature:
            item.status, item.error_msg = 'waiting', ''
            item.warning_msg = '源文件已变化，需要重新处理'
            item.file_type = detect_file_type(path)
            if item.file_type in ('error', 'unsupported'):
                item.status, item.error_msg = 'error', '源文件已变化，且无法识别当前格式'
        elif item.status == 'done' and (item.output_signature is None or
                                        file_signature(output) != item.output_signature):
            item.status, item.error_msg = 'waiting', ''
            item.warning_msg = '输出文件已丢失或变化，需要重新处理'
        elif item.status not in ('done', 'error', 'skipped', 'waiting'):
            item.status, item.error_msg = 'waiting', ''
            item.warning_msg = '上次处理未完成，将从该文件重新开始'
        item.source_signature = current
        if item.status != 'done':
            item.validation_msg = ''
        return item
