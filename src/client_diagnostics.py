"""只读检查 QQ 音乐接口；通过检查不代表所有文件均可解密。"""

import ctypes
import json
from ctypes import wintypes


REQUIRED_EXPORTS = [
    '??0EncAndDesMediaFile@@QAE@XZ', '??1EncAndDesMediaFile@@QAE@XZ',
    '?Open@EncAndDesMediaFile@@QAE_NPB_W_N1@Z',
    '?GetSize@EncAndDesMediaFile@@QAEKXZ', '?Read@EncAndDesMediaFile@@QAEKPAEK_J@Z',
]


def file_version(path):
    try:
        library = ctypes.WinDLL('version', use_last_error=True)
        library.GetFileVersionInfoSizeW.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(wintypes.DWORD)]
        library.GetFileVersionInfoW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID]
        library.VerQueryValueW.argtypes = [wintypes.LPCVOID, wintypes.LPCWSTR,
                                           ctypes.POINTER(wintypes.LPVOID), ctypes.POINTER(wintypes.UINT)]
        size = library.GetFileVersionInfoSizeW(path, None)
        if not size:
            return '未知'
        buffer = ctypes.create_string_buffer(size)
        if not library.GetFileVersionInfoW(path, 0, size, buffer):
            return '未知'
        pointer, length = wintypes.LPVOID(), wintypes.UINT()
        if not library.VerQueryValueW(buffer, '\\', ctypes.byref(pointer), ctypes.byref(length)) or length.value < 16:
            return '未知'
        words = ctypes.cast(pointer, ctypes.POINTER(wintypes.DWORD))
        return '.'.join(str(value) for value in (words[2] >> 16, words[2] & 65535,
                                                words[3] >> 16, words[3] & 65535))
    except (OSError, AttributeError, ValueError):
        return '未知'


def connection_error(exc):
    name = type(exc).__name__
    labels = {'ProcessNotFoundError': 'QQ 音乐未运行，请先启动客户端',
              'PermissionDeniedError': '无权连接 QQ 音乐，请检查两个程序的运行权限是否一致',
              'ProcessNotRespondingError': 'QQ 音乐进程无响应',
              'ModuleNotFoundError': '缺少 Frida 组件',
              'InvalidOperationError': 'QQ 音乐连接已断开，请重试'}
    return labels.get(name, '连接 QQ 音乐失败') + f'（{name}: {exc}）'


def assess_client(report):
    if report.get('arch') != 'ia32':
        return '当前解密接口仅支持 32 位 QQ 音乐，此客户端架构不兼容'
    if not report.get('module'):
        return '客户端尚未加载 QQMusicCommon.dll，请等待客户端启动完成后重试'
    if report.get('missing_exports'):
        return '客户端版本不兼容，缺少必要接口: ' + ', '.join(report['missing_exports'])
    return ''


def inspect_session(session):
    code = 'const required = ' + json.dumps(REQUIRED_EXPORTS) + ';' + '''
rpc.exports = {inspect() {
    const module = Process.findModuleByName('QQMusicCommon.dll');
    return {pid: Process.id, arch: Process.arch, executable: Process.enumerateModules()[0].path,
            module: module ? module.path : null,
            missing_exports: required.filter(name => !module || !module.findExportByName(name))};
}};
'''
    script = session.create_script(code)
    try:
        script.load()
        report = script.exports_sync.inspect()
        report['client_version'] = file_version(report.get('executable', ''))
        report['error'] = assess_client(report)
        return report
    finally:
        try:
            script.unload()
        except Exception:
            pass


def inspect_running_client():
    session = None
    try:
        import frida
        session = frida.attach('QQMusic.exe')
        report = inspect_session(session)
        report['frida_version'] = frida.__version__
        return report
    except Exception as exc:
        return {'error': connection_error(exc)}
    finally:
        if session:
            try:
                session.detach()
            except Exception:
                pass


def environment_report(converter):
    report = inspect_running_client()
    lines = ['QQ 音乐兼容性诊断', '', f"检查结果: {report.get('error') or '必要接口检查通过（仍需文件实测）'}"]
    for key, label in [('client_version', '客户端版本'), ('arch', '进程架构'),
                       ('pid', '进程 ID'), ('executable', '客户端路径'), ('frida_version', 'Frida 版本')]:
        if key in report:
            lines.append(f'{label}: {report[key]}')
    lines.extend(['', '本地音频工具'])
    for attr in ('flac_path', 'oggdec_path', 'lame_path', 'ffmpeg_path'):
        path = getattr(converter, attr)
        lines.append(f"{path.name}: {'可用' if path.is_file() else '缺失'} | {path}")
    return '\n'.join(lines)
