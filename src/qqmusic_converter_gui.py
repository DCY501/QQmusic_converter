#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QQ音乐解密/转码工具 - GUI版本
支持格式: .mflac .mgg .mmp4 .flac .ogg .wav .mp3 .m4a
"""

import sys
import os
import subprocess
import threading
import shutil
import tempfile
import json
import hashlib
from pathlib import Path
from datetime import datetime
from typing import List, Optional, Dict, Tuple
from dataclasses import dataclass, asdict
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import warnings

warnings.filterwarnings("ignore", "libpng warning")
warnings.filterwarnings("ignore", "iCCP")


# ============ 路径工具 ============

def get_app_dir() -> Path:
    """获取程序所在目录（兼容开发环境和打包环境）"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


APP_DIR = get_app_dir()


# Windows UTF-8 编码支持
if sys.platform == 'win32' and sys.stdout is not None:
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


# ============ 数据结构 ============

@dataclass
class FileItem:
    """文件项数据类"""
    id: str
    path: Path
    name: str
    size: int
    status: str  # 'waiting', 'converting', 'done', 'error'
    file_type: str  # 检测到的文件类型
    error_msg: str = ""

    @property
    def size_mb(self) -> float:
        return self.size / (1024 * 1024)


# ============ 配置管理 ============

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


# ============ 文件类型检测 ============

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


# ============ QQ音乐 FRIDA 解密器 ============

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


# ============ 音频转换器 ============

class AudioConverter:
    """音频转换器 - 封装小工具链"""

    def __init__(self):
        self.flac_path = APP_DIR / "flac.exe"
        self.oggdec_path = APP_DIR / "oggdec.exe"
        self.lame_path = APP_DIR / "lame.exe"
        # ffmpeg 作为备选解码器（处理 flac.exe 无法解码的损坏/特殊 FLAC）
        self.ffmpeg_path = APP_DIR.parent / "ffmpeg.exe"
        self._check_tools()

    def _check_tools(self):
        """检查工具是否存在"""
        self.flac_ok = self.flac_path.exists()
        self.oggdec_ok = self.oggdec_path.exists()
        self.lame_ok = self.lame_path.exists()
        self.ffmpeg_ok = self.ffmpeg_path.exists()

    def _run(self, cmd: List[str], timeout: int = 300) -> Tuple[bool, str]:
        """运行命令"""
        try:
            result = subprocess.run(
                cmd, capture_output=True, text=True,
                encoding='utf-8', errors='ignore',
                timeout=timeout, creationflags=subprocess.CREATE_NO_WINDOW
            )
            return result.returncode == 0, result.stderr or result.stdout
        except Exception as e:
            return False, str(e)

    def flac_to_wav(self, input_path: Path, output_path: Path) -> bool:
        """FLAC → WAV（flac.exe 失败时尝试 ffmpeg 备选）"""
        if not self.flac_ok:
            return False
        ok, _ = self._run([str(self.flac_path), "-d", "-f", "-o", str(output_path), str(input_path)])
        if ok and output_path.exists():
            return True
        # flac.exe 失败时，尝试用 ffmpeg 作为备选（更容错，可跳过损坏帧）
        if self.ffmpeg_ok:
            ok, _ = self._run([
                str(self.ffmpeg_path), "-hide_banner", "-y",
                "-i", str(input_path),
                "-ar", "44100", "-ac", "2",
                str(output_path)
            ])
            return ok and output_path.exists()
        return False

    def ogg_to_wav(self, input_path: Path, output_path: Path) -> bool:
        """OGG → WAV (oggdec 不支持空格路径，需临时复制)"""
        if not self.oggdec_ok:
            return False

        # oggdec 不支持空格路径，临时复制到纯 ASCII 路径
        tmp_dir = Path(tempfile.gettempdir()) / "qqmusic_ogg"
        tmp_dir.mkdir(exist_ok=True)
        tmp_ogg = tmp_dir / "tmp.ogg"
        tmp_wav = tmp_dir / "tmp.wav"

        try:
            shutil.copy2(str(input_path), str(tmp_ogg))
            ok, _ = self._run([str(self.oggdec_path), str(tmp_ogg)])
            if ok and tmp_wav.exists():
                if output_path.exists():
                    output_path.unlink()
                shutil.move(str(tmp_wav), str(output_path))
                return True
            return False
        finally:
            # 清理临时文件
            for f in [tmp_ogg, tmp_wav]:
                if f.exists():
                    try:
                        f.unlink()
                    except:
                        pass

    def wav_to_mp3(self, input_path: Path, output_path: Path, quality: str) -> bool:
        """WAV → MP3"""
        if not self.lame_ok:
            return False

        if quality == 'vbr':
            cmd = [str(self.lame_path), '-V', '0', str(input_path), str(output_path)]
        elif quality == '320k':
            cmd = [str(self.lame_path), '-b', '320', str(input_path), str(output_path)]
        else:  # 192k
            cmd = [str(self.lame_path), '-b', '192', str(input_path), str(output_path)]

        ok, _ = self._run(cmd)
        return ok and output_path.exists()

    def convert_to_mp3(self, input_path: Path, output_path: Path, file_type: str, quality: str) -> bool:
        """通用转 MP3 入口"""
        tmp_dir = Path(tempfile.gettempdir()) / "qqmusic_conv"
        tmp_dir.mkdir(exist_ok=True)
        tmp_wav = tmp_dir / f"{input_path.stem}.wav"

        try:
            # Step 1: 解码为 WAV
            if file_type in ['standard_flac', 'qmc_mflac_decrypted']:
                if not self.flac_to_wav(input_path, tmp_wav):
                    return False
            elif file_type in ['standard_ogg', 'qmc_mgg_decrypted']:
                if not self.ogg_to_wav(input_path, tmp_wav):
                    return False
            elif file_type == 'standard_wav':
                tmp_wav = input_path  # 直接就是 WAV
            elif file_type == 'standard_mp3':
                # MP3 → MP3，直接复制
                shutil.copy2(str(input_path), str(output_path))
                return True
            else:
                return False

            # Step 2: WAV → MP3
            if file_type == 'standard_wav':
                return self.wav_to_mp3(input_path, output_path, quality)
            else:
                ok = self.wav_to_mp3(tmp_wav, output_path, quality)
                return ok
        finally:
            # 清理临时 WAV
            if tmp_wav != input_path and tmp_wav.exists():
                try:
                    tmp_wav.unlink()
                except:
                    pass


# ============ 主 GUI 类 ============

class ConverterGUI:
    """主GUI类"""

    def __init__(self, root: Tk):
        self.root = root
        self.root.title("QQ音乐解密/转码工具")
        self.root.geometry("900x600")
        self.root.minsize(700, 450)
        self.center_window()

        # 主题
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TFrame', background='#f5f6f7')
        self.style.configure('TLabel', background='#f5f6f7', font=('Microsoft YaHei', 10))
        self.style.configure('TButton', font=('Microsoft YaHei', 9))
        self.style.configure('TLabelframe', background='#f5f6f7')
        self.style.configure('TLabelframe.Label', background='#f5f6f7', font=('Microsoft YaHei', 9, 'bold'))
        self.style.configure('Treeview', font=('Microsoft YaHei', 9))
        self.style.configure('Treeview.Heading', font=('Microsoft YaHei', 9, 'bold'))
        self.root.configure(bg='#f5f6f7')

        self.default_font = ('Microsoft YaHei', 10)
        self.root.option_add('*Font', self.default_font)

        # 数据
        self.files: Dict[str, FileItem] = {}
        self.file_counter = 0
        self.is_converting = False
        self.stop_flag = False

        # 工具
        self.decryptor = QQMusicDecryptor()
        self.converter = AudioConverter()

        # 加载配置
        self.config = Config.load()

        # 创建UI
        self.create_ui()

        # 检查工具
        self._check_tools()

        # 关闭处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def center_window(self):
        """窗口居中"""
        self.root.update_idletasks()
        width = 900
        height = 600
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def _check_tools(self):
        """检查工具状态"""
        missing = []
        if not self.converter.flac_ok:
            missing.append("flac.exe")
        if not self.converter.oggdec_ok:
            missing.append("oggdec.exe")
        if not self.converter.lame_ok:
            missing.append("lame.exe")

        if missing:
            status = f"缺少工具: {', '.join(missing)}"
        else:
            status = "工具就绪 | flac + oggdec + lame"

        # 检查 frida
        try:
            import frida
            status += " | frida 已安装"
        except:
            status += " | frida 未安装 (pip install frida)"

        self.status_bar.config(text=status)

    def create_ui(self):
        """创建界面"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(N, S, E, W))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # ===== 工具栏 =====
        toolbar = ttk.Frame(main_frame)
        toolbar.grid(row=0, column=0, sticky=(W, E), pady=(0, 10))

        ttk.Button(toolbar, text="📁 添加文件", command=self.add_files).pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="📂 添加文件夹", command=self.add_folder).pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="🗑️ 删除", command=self.remove_selected).pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="🧹 清空", command=self.clear_all).pack(side=LEFT)

        # ===== 文件列表 =====
        list_frame = ttk.LabelFrame(main_frame, text="文件列表", padding="5")
        list_frame.grid(row=1, column=0, rowspan=2, sticky=(N, S, E, W), pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        columns = ('name', 'type', 'size', 'status', 'action')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', selectmode='extended')

        self.tree.heading('name', text='文件名')
        self.tree.heading('type', text='类型')
        self.tree.heading('size', text='大小')
        self.tree.heading('status', text='状态')
        self.tree.heading('action', text='操作')

        self.tree.column('name', width=300)
        self.tree.column('type', width=80, anchor=CENTER)
        self.tree.column('size', width=80, anchor=CENTER)
        self.tree.column('status', width=100, anchor=CENTER)
        self.tree.column('action', width=80, anchor=CENTER)

        vsb = ttk.Scrollbar(list_frame, orient=VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(list_frame, orient=HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky=(N, S, E, W))
        vsb.grid(row=0, column=1, sticky=(N, S))
        hsb.grid(row=1, column=0, sticky=(E, W))

        self.tree.bind('<Delete>', lambda e: self.remove_selected())
        self.tree.bind('<Double-1>', self.on_item_double_click)

        self.file_stats_label = ttk.Label(list_frame, text="共0个文件，总大小0MB")
        self.file_stats_label.grid(row=2, column=0, columnspan=2, sticky=W, pady=(5, 0))

        # ===== 输出设置 =====
        settings_frame = ttk.LabelFrame(main_frame, text="输出设置", padding="5")
        settings_frame.grid(row=3, column=0, sticky=(W, E), pady=(0, 10))
        settings_frame.columnconfigure(1, weight=1)

        # 输出路径
        ttk.Label(settings_frame, text="保存到:").grid(row=0, column=0, sticky=W, padx=(0, 5))
        default_output = self.config.get('output_path', str((APP_DIR.parent / "converted").resolve()))
        self.output_path_var = StringVar(value=default_output)
        ttk.Entry(settings_frame, textvariable=self.output_path_var).grid(row=0, column=1, sticky=(W, E), padx=(0, 5))
        ttk.Button(settings_frame, text="浏览...", command=self.browse_output).grid(row=0, column=2)

        # 输出格式
        ttk.Label(settings_frame, text="输出格式:").grid(row=1, column=0, sticky=W, padx=(0, 5), pady=(10, 0))
        self.output_format_var = StringVar(value=self.config.get('output_format', 'keep'))
        format_frame = ttk.Frame(settings_frame)
        format_frame.grid(row=1, column=1, sticky=W, pady=(10, 0))
        ttk.Radiobutton(format_frame, text="保持原格式", variable=self.output_format_var, value='keep').pack(side=LEFT, padx=(0, 15))
        ttk.Radiobutton(format_frame, text="统一转 MP3", variable=self.output_format_var, value='mp3').pack(side=LEFT, padx=(0, 15))

        # 音质选项 (仅 MP3 时有效)
        ttk.Label(settings_frame, text="MP3音质:").grid(row=2, column=0, sticky=W, padx=(0, 5), pady=(10, 0))
        self.quality_var = StringVar(value=self.config.get('quality', 'vbr'))
        quality_frame = ttk.Frame(settings_frame)
        quality_frame.grid(row=2, column=1, sticky=W, pady=(10, 0))
        ttk.Radiobutton(quality_frame, text="VBR高质量", variable=self.quality_var, value='vbr').pack(side=LEFT, padx=(0, 15))
        ttk.Radiobutton(quality_frame, text="CBR 320k", variable=self.quality_var, value='320k').pack(side=LEFT, padx=(0, 15))
        ttk.Radiobutton(quality_frame, text="CBR 192k", variable=self.quality_var, value='192k').pack(side=LEFT)

        # ===== 转换区域 =====
        convert_frame = ttk.Frame(main_frame)
        convert_frame.grid(row=4, column=0, sticky=(W, E))
        convert_frame.columnconfigure(0, weight=1)

        self.convert_btn = ttk.Button(convert_frame, text="开始处理 (0个文件)", command=self.start_conversion)
        self.convert_btn.grid(row=0, column=0, sticky=(W, E), pady=(0, 5))

        self.progress_var = DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(convert_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=1, column=0, sticky=(W, E), pady=(0, 5))

        self.status_label = ttk.Label(convert_frame, text="就绪")
        self.status_label.grid(row=2, column=0, sticky=W)

        # 状态栏
        self.status_bar = ttk.Label(main_frame, text="检查工具中...", relief=SUNKEN, anchor=W)
        self.status_bar.grid(row=5, column=0, sticky=(W, E), pady=(10, 0))

    def add_files(self):
        """添加文件"""
        files = filedialog.askopenfilenames(
            title="选择音频文件",
            filetypes=[
                ("QQ音乐加密", "*.mflac *.mgg *.mmp4 *.mflac0 *.mgg0 *.mgg1 *.mggl"),
                ("音频文件", "*.flac *.ogg *.wav *.mp3 *.m4a"),
                ("所有文件", "*.*")
            ]
        )
        added = 0
        for f in files:
            if self.add_file_item(Path(f)):
                added += 1
        if added > 0:
            self.update_file_stats()
            self.update_convert_button()

    def add_folder(self):
        """添加文件夹"""
        folder = filedialog.askdirectory(title="选择包含音频文件的文件夹")
        if not folder:
            return
        folder_path = Path(folder)
        patterns = ['*.mflac', '*.mgg', '*.mmp4', '*.mflac0', '*.mgg0', '*.mgg1', '*.mggl',
                    '*.flac', '*.ogg', '*.wav', '*.mp3', '*.m4a']
        all_files = []
        for p in patterns:
            all_files.extend(folder_path.rglob(p))
        added = 0
        for path in all_files:
            if self.add_file_item(path):
                added += 1
        if added > 0:
            self.update_file_stats()
            self.update_convert_button()
            messagebox.showinfo("提示", f"已添加 {added} 个文件")
        elif all_files:
            messagebox.showinfo("提示", "所有文件已在列表中")
        else:
            messagebox.showinfo("提示", "未找到支持的音频文件")

    def add_file_item(self, path: Path) -> bool:
        """添加单个文件到列表"""
        # 检查是否已存在
        for item in self.files.values():
            if item.path == path:
                return False

        # 检测文件类型
        file_type = detect_file_type(path)
        if file_type == 'unsupported':
            return False
        if file_type == 'error':
            return False

        self.file_counter += 1
        item_id = f"file_{self.file_counter}"

        try:
            size = path.stat().st_size
        except:
            size = 0

        type_labels = {
            'qmc_mflac': '加密FLAC',
            'qmc_mgg': '加密OGG',
            'qmc_mmp4': '加密MP4',
            'standard_flac': 'FLAC',
            'standard_ogg': 'OGG',
            'standard_wav': 'WAV',
            'standard_mp3': 'MP3',
            'standard_m4a': 'M4A',
        }

        item = FileItem(
            id=item_id,
            path=path,
            name=path.name,
            size=size,
            status='waiting',
            file_type=file_type
        )

        self.files[item_id] = item
        self.tree.insert('', END, iid=item_id, values=(
            item.name,
            type_labels.get(file_type, file_type),
            f"{item.size_mb:.2f}MB",
            "等待中",
            ""
        ))
        return True

    def remove_selected(self):
        """删除选中项"""
        selected = self.tree.selection()
        for item_id in selected:
            if item_id in self.files:
                del self.files[item_id]
                self.tree.delete(item_id)
        self.update_file_stats()
        self.update_convert_button()

    def clear_all(self):
        """清空所有"""
        if not self.files:
            return
        if messagebox.askyesno("确认", "确定要清空所有文件吗？"):
            self.files.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.update_file_stats()
            self.update_convert_button()

    def update_file_stats(self):
        """更新文件统计"""
        total = len(self.files)
        total_size = sum(f.size for f in self.files.values())
        self.file_stats_label.config(text=f"共{total}个文件，总大小{total_size/(1024*1024):.1f}MB")

    def update_convert_button(self):
        """更新转换按钮"""
        waiting_count = sum(1 for f in self.files.values() if f.status == 'waiting')
        self.convert_btn.config(text=f"开始处理 ({waiting_count}个文件)")

    def browse_output(self):
        """浏览输出目录"""
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            self.output_path_var.set(folder)

    def on_item_double_click(self, event):
        """双击打开文件位置"""
        selected = self.tree.selection()
        if not selected:
            return
        item_id = selected[0]
        if item_id in self.files:
            folder = str(self.files[item_id].path.parent)
            if os.path.exists(folder):
                os.startfile(folder)

    def update_item_status(self, item_id: str, status_text: str, color: str):
        """更新列表项状态"""
        if item_id in self.files:
            item = self.files[item_id]
            type_labels = {
                'qmc_mflac': '加密FLAC', 'qmc_mgg': '加密OGG', 'qmc_mmp4': '加密MP4',
                'standard_flac': 'FLAC', 'standard_ogg': 'OGG', 'standard_wav': 'WAV',
                'standard_mp3': 'MP3', 'standard_m4a': 'M4A',
            }
            self.tree.item(item_id, values=(
                item.name,
                type_labels.get(item.file_type, item.file_type),
                f"{item.size_mb:.2f}MB",
                status_text,
                ""
            ), tags=(color,))
            self.tree.tag_configure('blue', foreground='blue')
            self.tree.tag_configure('green', foreground='green')
            self.tree.tag_configure('red', foreground='red')
            self.tree.tag_configure('gray', foreground='gray')

    def start_conversion(self):
        """开始转换"""
        if self.is_converting:
            self.stop_flag = True
            self.convert_btn.config(text="正在停止...")
            return

        waiting_files = [f for f in self.files.values() if f.status == 'waiting']
        if not waiting_files:
            messagebox.showinfo("提示", "没有等待处理的文件")
            return

        # 检查是否有加密文件需要 FRIDA
        has_encrypted = any(f.file_type.startswith('qmc_') for f in waiting_files)
        if has_encrypted:
            try:
                import frida
            except ImportError:
                messagebox.showerror("错误", "处理加密文件需要安装 frida\n请运行: pip install frida")
                return

            if not self.decryptor.is_qqmusic_running():
                messagebox.showerror("错误", "请先启动 QQ 音乐客户端！\n\n处理加密文件 (.mflac/.mgg/.mmp4) 需要 QQ 音乐保持运行。")
                return

        # 检查输出目录
        output_dir = Path(self.output_path_var.get())
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出目录: {e}")
            return

        # 检查工具
        need_mp3 = self.output_format_var.get() == 'mp3'
        if need_mp3:
            missing = []
            # 根据文件类型判断需要哪些解码工具
            need_flac = any(f.file_type in ['qmc_mflac', 'standard_flac'] for f in waiting_files)
            need_ogg = any(f.file_type in ['qmc_mgg', 'standard_ogg'] for f in waiting_files)
            
            if need_flac and not self.converter.flac_ok:
                missing.append("flac.exe")
            if need_ogg and not self.converter.oggdec_ok:
                missing.append("oggdec.exe")
            if not self.converter.lame_ok:
                missing.append("lame.exe")
            if missing:
                messagebox.showerror("错误", f"缺少必要工具: {', '.join(missing)}")
                return

        # 保存配置
        self.config['output_path'] = str(output_dir)
        self.config['output_format'] = self.output_format_var.get()
        self.config['quality'] = self.quality_var.get()
        Config.save(self.config)

        # 启动转换线程
        self.is_converting = True
        self.stop_flag = False
        self.convert_btn.config(text="停止处理")
        self.progress_var.set(0)

        thread = threading.Thread(target=self.convert_worker, args=(waiting_files, output_dir))
        thread.daemon = True
        thread.start()

    def convert_worker(self, files: List[FileItem], output_dir: Path):
        """后台处理线程"""
        total = len(files)
        output_format = self.output_format_var.get()
        quality = self.quality_var.get()

        # 如果有加密文件，先连接 FRIDA
        has_encrypted = any(f.file_type.startswith('qmc_') for f in files)
        if has_encrypted:
            self.root.after(0, lambda: self.status_label.config(text="正在连接 QQ 音乐..."))
            if not self.decryptor.connect():
                self.root.after(0, messagebox.showerror, "错误", "无法 attach 到 QQMusic.exe")
                self.is_converting = False
                self.root.after(0, self.conversion_finished, total)
                return

        for i, file_item in enumerate(files):
            if self.stop_flag:
                break

            file_item.status = 'converting'
            self.root.after(0, self.update_item_status, file_item.id, '处理中', 'blue')
            self.root.after(0, lambda n=file_item.name: self.status_label.config(text=f"正在处理: {n}"))
            self.root.after(0, lambda i=i, t=total, n=file_item.name: self.status_bar.config(text=f"处理中 ({i+1}/{t}): {n}"))

            try:
                # 确定输出路径
                if output_format == 'keep':
                    # 保持原格式
                    if file_item.file_type == 'qmc_mflac':
                        ext = '.flac'
                    elif file_item.file_type == 'qmc_mgg':
                        ext = '.ogg'
                    elif file_item.file_type == 'qmc_mmp4':
                        ext = '.mp4'
                    else:
                        ext = file_item.path.suffix
                    output_path = output_dir / (file_item.path.stem + ext)
                else:
                    # 统一转 MP3
                    if file_item.file_type == 'qmc_mmp4':
                        # .mmp4 不支持转 MP3
                        file_item.status = 'skipped'
                        file_item.error_msg = ".mmp4 (E-AC3格式) 暂不支持转 MP3，请选择保持原格式"
                        self.root.after(0, self.update_item_status, file_item.id, '不支持转MP3', 'gray')
                        self.log_error(file_item)
                        continue
                    output_path = output_dir / (file_item.path.stem + '.mp3')

                # 检查是否已存在
                if output_path.exists():
                    pass

                # Step 1: 解密（如果是加密文件）
                current_path = file_item.path
                decrypted_type = file_item.file_type

                if file_item.file_type.startswith('qmc_'):
                    self.root.after(0, lambda n=file_item.name: self.status_label.config(text=f"解密中: {n}"))
                    # 使用唯一临时文件名，避免同名不同扩展名的文件冲突
                    tmp_suffix = file_item.path.suffix.replace('.', '_')
                    decrypted_path = output_dir / (file_item.path.stem + f"{tmp_suffix}.decrypted.tmp")
                    if self.decryptor.decrypt(file_item.path, decrypted_path):
                        current_path = decrypted_path
                        if file_item.file_type == 'qmc_mflac':
                            decrypted_type = 'qmc_mflac_decrypted'
                        elif file_item.file_type == 'qmc_mgg':
                            decrypted_type = 'qmc_mgg_decrypted'
                        elif file_item.file_type == 'qmc_mmp4':
                            decrypted_type = 'qmc_mmp4_decrypted'
                    else:
                        file_item.status = 'error'
                        file_item.error_msg = "FRIDA 解密失败"
                        self.root.after(0, self.update_item_status, file_item.id, '失败', 'red')
                        self.log_error(file_item)
                        continue

                # Step 2: 转码（如果需要）
                if output_format == 'keep':
                    # 保持原格式，直接移动/复制解密后的文件
                    if current_path != file_item.path:
                        if output_path.exists():
                            output_path.unlink()
                        shutil.move(str(current_path), str(output_path))
                    else:
                        shutil.copy2(str(current_path), str(output_path))
                    file_item.status = 'done'
                    self.root.after(0, self.update_item_status, file_item.id, '已完成', 'green')
                else:
                    # 转 MP3
                    self.root.after(0, lambda n=file_item.name: self.status_label.config(text=f"转码中: {n}"))
                    if self.converter.convert_to_mp3(current_path, output_path, decrypted_type, quality):
                        file_item.status = 'done'
                        self.root.after(0, self.update_item_status, file_item.id, '已完成', 'green')
                        # 清理临时解密文件
                        if current_path != file_item.path and current_path.exists():
                            try:
                                current_path.unlink()
                            except:
                                pass
                    else:
                        file_item.status = 'error'
                        file_item.error_msg = "转码失败"
                        self.root.after(0, self.update_item_status, file_item.id, '失败', 'red')
                        self.log_error(file_item)

            except Exception as e:
                file_item.status = 'error'
                file_item.error_msg = str(e)
                self.root.after(0, self.update_item_status, file_item.id, '失败', 'red')
                self.log_error(file_item)

            # 更新进度
            progress = ((i + 1) / total) * 100
            self.root.after(0, self.progress_var.set, progress)

        # 断开 FRIDA
        if has_encrypted:
            self.decryptor.disconnect()

        # 完成
        self.is_converting = False
        self.root.after(0, self.conversion_finished, total)

    def conversion_finished(self, total: int):
        """处理完成回调"""
        waiting = len([f for f in self.files.values() if f.status == 'waiting'])
        self.convert_btn.config(text=f"开始处理 ({waiting}个文件)")
        self.status_label.config(text="处理完成")
        self.status_bar.config(text="就绪")

        done_count = sum(1 for f in self.files.values() if f.status == 'done')
        error_count = sum(1 for f in self.files.values() if f.status == 'error')
        skip_count = sum(1 for f in self.files.values() if f.status == 'skipped')

        msg_lines = ["处理完成！"]
        if done_count > 0:
            msg_lines.append(f"成功: {done_count} 个")
        if skip_count > 0:
            msg_lines.append(f"跳过: {skip_count} 个")
        if error_count > 0:
            msg_lines.append(f"失败: {error_count} 个")
            msg_lines.append("\n失败记录已保存到 error.log")

        if error_count > 0:
            messagebox.showwarning("完成", "\n".join(msg_lines))
        elif skip_count > 0:
            messagebox.showinfo("完成", "\n".join(msg_lines))
        else:
            messagebox.showinfo("完成", f"处理完成！\n成功处理 {done_count} 个文件")

        self.update_convert_button()

    def log_error(self, item: FileItem):
        """记录错误日志"""
        try:
            log_path = APP_DIR / "error.log"
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {item.path}\n")
                f.write(f"  类型: {item.file_type}\n")
                f.write(f"  错误: {item.error_msg}\n\n")
        except:
            pass

    def on_close(self):
        """关闭窗口"""
        if self.is_converting:
            if not messagebox.askyesno("确认", "处理正在进行中，确定要退出吗？"):
                return
            self.stop_flag = True

        # 断开 FRIDA
        self.decryptor.disconnect()

        # 保存配置
        self.config['output_path'] = self.output_path_var.get()
        self.config['output_format'] = self.output_format_var.get()
        self.config['quality'] = self.quality_var.get()
        Config.save(self.config)

        self.root.destroy()


def main():
    """主函数"""
    root = Tk()
    app = ConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
