#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主 GUI 界面
"""

import os
import threading
import time
from dataclasses import replace
from queue import Empty, Queue
from pathlib import Path
from typing import Dict
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import warnings

warnings.filterwarnings("ignore", "libpng warning")
warnings.filterwarnings("ignore", "iCCP")

from models import FileItem, Config, detect_file_type, TYPE_LABELS, EXTENSION_TYPES, APP_DIR
from scanner import scan_paths
from decryptor import QQMusicDecryptor
from converter import AudioConverter
from worker import BatchConverter
from session_store import SessionStore, file_signature
from file_details import describe_item
from client_diagnostics import environment_report


class ConverterGUI:
    """主GUI类"""

    def __init__(self, root: Tk, session_store=None, instance_guard=None):
        self.root = root
        self.instance_guard = instance_guard
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
        self.path_keys = set()
        self.file_counter = 0
        self.is_converting = False
        self.stop_event = threading.Event()
        self.events = Queue()
        self.worker_thread = None
        self.closing = False
        self._pending_stats = None
        self.is_scanning = False
        self.scan_thread = None
        self.scan_stop = threading.Event()
        self._scan_result = None
        self.scan_added = 0
        self.scan_errors = []
        self.session_store = session_store if session_store is not None else SessionStore()
        self._session_dirty = False
        self._last_session_save = time.monotonic()
        self._active_stage = None
        self._file_started = None
        self._detail_requests = {}
        self._detail_counter = 0

        # 工具
        self.decryptor = QQMusicDecryptor()
        self.converter = AudioConverter()
        self.batch_converter = BatchConverter(self.decryptor, self.converter)

        # 加载配置
        self.config = Config.load()

        # 创建UI
        self.create_ui()
        menu = Menu(self.root)
        tools_menu = Menu(menu, tearoff=False)
        tools_menu.add_command(label='环境与兼容性诊断', command=self.show_environment_diagnostics)
        tools_menu.add_command(label='打开配置目录', command=self.open_config_folder)
        menu.add_cascade(label='工具', menu=tools_menu)
        self.root.config(menu=menu)
        if getattr(self.root, 'dnd_available', False):
            from tkinterdnd2 import DND_FILES
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.on_drop)
        self._update_quality_state()
        self.output_format_var.trace_add('write', lambda *_: self._update_quality_state())

        # 检查工具
        self._check_tools()
        self._restore_session()
        self.update_convert_button()
        self.update_file_stats()
        for variable in (self.output_path_var, self.output_format_var, self.quality_var, self.verify_full_var):
            variable.trace_add('write', self._mark_session_dirty)

        # 关闭处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self._poll_id = self.root.after(75, self._poll_events)

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
        except ImportError:
            status += " | frida 未安装 (pip install frida)"

        self.status_bar.config(text=status)
        if Config.last_error:
            self.status_bar.config(text=Config.last_error)

    def on_drop(self, event):
        if self.is_converting or self.is_scanning or self.closing:
            return 'refuse_drop'
        self._start_scan(self.root.tk.splitlist(event.data))
        return 'copy'

    # ========== UI 创建 ==========

    def create_ui(self):
        """创建界面"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(N, S, E, W))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # 工具栏
        self._create_toolbar(main_frame)
        # 文件列表
        self._create_file_list(main_frame)
        # 输出设置
        self._create_settings(main_frame)
        # 转换区域
        self._create_convert_area(main_frame)

    def _create_toolbar(self, parent):
        """创建工具栏"""
        toolbar = ttk.Frame(parent)
        self.toolbar = toolbar
        toolbar.grid(row=0, column=0, sticky=(W, E), pady=(0, 10))

        ttk.Button(toolbar, text="📁 添加文件", command=self.add_files).pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="📂 添加文件夹", command=self.add_folder).pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="🗑️ 删除", command=self.remove_selected).pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="🧹 清空", command=self.clear_all).pack(side=LEFT)
        ttk.Button(toolbar, text="重试失败", command=self.retry_failed).pack(side=LEFT, padx=5)
        ttk.Button(toolbar, text="打开输出目录", command=self.open_output_folder).pack(side=RIGHT)

    def _create_file_list(self, parent):
        """创建文件列表"""
        title = '文件列表（可拖入文件或文件夹）' if getattr(self.root, 'dnd_available', False) else '文件列表'
        list_frame = ttk.LabelFrame(parent, text=title, padding="5")
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
        hint = '拖入音频文件或文件夹\n也可以点击上方“添加文件”' if getattr(self.root, 'dnd_available', False) else '点击上方“添加文件”或“添加文件夹”'
        self.empty_hint = ttk.Label(self.tree, text=hint, justify=CENTER,
                                    background='white', foreground='#687078')
        self.empty_hint.place(relx=0.5, rely=0.55, anchor=CENTER)
        vsb.grid(row=0, column=1, sticky=(N, S))
        hsb.grid(row=1, column=0, sticky=(E, W))

        self.tree.bind('<Delete>', lambda e: self.remove_selected())
        self.tree.bind('<Double-1>', self.on_item_double_click)
        self.tree.bind('<Button-1>', self.on_action_click)
        self.tree.bind('<Button-3>', self.on_context_menu)
        self.context_menu = Menu(self.root, tearoff=False)
        self.context_menu.add_command(label='重新处理选中项', command=self.reprocess_selected)
        self.context_menu.add_command(label='文件详情', command=self.show_selected_details)

        self.file_stats_label = ttk.Label(list_frame, text="共0个文件，总大小0MB")
        self.file_stats_label.grid(row=2, column=0, columnspan=2, sticky=W, pady=(5, 0))

    def _create_settings(self, parent):
        """创建输出设置面板"""
        settings_frame = ttk.LabelFrame(parent, text="输出设置", padding="5")
        self.settings_frame = settings_frame
        settings_frame.grid(row=3, column=0, sticky=(W, E), pady=(0, 10))
        settings_frame.columnconfigure(1, weight=1)

        # 输出路径
        ttk.Label(settings_frame, text="保存到:").grid(row=0, column=0, sticky=W, padx=(0, 5))
        default_output = self.config.get('output_path', str((APP_DIR / "converted").resolve()))
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
        self.quality_frame = quality_frame
        quality_frame.grid(row=2, column=1, sticky=W, pady=(10, 0))
        ttk.Radiobutton(quality_frame, text="VBR高质量", variable=self.quality_var, value='vbr').pack(side=LEFT, padx=(0, 15))
        ttk.Radiobutton(quality_frame, text="CBR 320k", variable=self.quality_var, value='320k').pack(side=LEFT, padx=(0, 15))
        ttk.Radiobutton(quality_frame, text="CBR 192k", variable=self.quality_var, value='192k').pack(side=LEFT)
        self.verify_full_var = BooleanVar(value=self.config.get('verify_full', False))
        ttk.Checkbutton(settings_frame, text='完整解码校验（需要 FFmpeg，耗时更长）',
                        variable=self.verify_full_var).grid(row=3, column=1, sticky=W, pady=(8, 0))

    def _create_convert_area(self, parent):
        """创建转换区域"""
        convert_frame = ttk.Frame(parent)
        convert_frame.grid(row=4, column=0, sticky=(W, E))
        convert_frame.columnconfigure(0, weight=1)

        self.convert_btn = ttk.Button(convert_frame, text="开始处理 (0个文件)", command=self.start_conversion)
        self.convert_btn.grid(row=0, column=0, sticky=(W, E), pady=(0, 5))

        self.progress_var = DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(convert_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=1, column=0, sticky=(W, E), pady=(0, 5))

        self.stage_bar = ttk.Progressbar(convert_frame, maximum=100)
        self.stage_bar.grid(row=2, column=0, sticky=(W, E), pady=(0, 5))
        self.progress_bar.grid_remove()
        self.stage_bar.grid_remove()
        self.status_label = ttk.Label(convert_frame, text="就绪")
        self.status_label.grid(row=3, column=0, sticky=W)

        # 状态栏
        self.status_bar = ttk.Label(parent, text="检查工具中...", relief=SUNKEN, anchor=W)
        self.status_bar.grid(row=5, column=0, sticky=(W, E), pady=(10, 0))

    # ========== 文件管理 ==========

    def add_files(self):
        """添加文件"""
        if self.is_converting or self.is_scanning or self.closing:
            return
        files = filedialog.askopenfilenames(
            title="选择音频文件",
            filetypes=[
                ("支持的音频文件", " ".join('*' + ext for ext in EXTENSION_TYPES)),
                ("所有文件", "*.*")
            ]
        )
        if files:
            self._start_scan(files)

    def add_folder(self):
        """添加文件夹"""
        if self.is_converting or self.is_scanning or self.closing:
            return
        folder = filedialog.askdirectory(title="选择包含音频文件的文件夹")
        if not folder:
            return
        self._start_scan([folder])

    def _start_scan(self, paths):
        if self.is_converting or self.is_scanning or self.closing:
            return
        self.is_scanning = True
        self.scan_stop.clear()
        self.scan_added = 0
        self.scan_errors = []
        self._set_controls_enabled(False)
        self.convert_btn.config(state="disabled")
        self.status_label.config(text="正在扫描文件…")
        output = self.output_path_var.get().strip()
        def run():
            try:
                scan_paths(paths,
                           lambda *args: self.events.put(("scan_file", *args)),
                           lambda error: self.events.put(("scan_error", error)),
                           self.scan_stop.is_set, output or None)
            except Exception as exc:
                self.events.put(("scan_error", str(exc)))
            finally:
                self.events.put(("scan_finished",))
        self.scan_thread = threading.Thread(target=run, daemon=False)
        self.scan_thread.start()

    def add_file_item(self, path: Path, file_type=None, size=None) -> bool:
        """添加单个文件到列表"""
        path = path.resolve()
        key = os.path.normcase(str(path))
        if key in self.path_keys:
            return False
        file_type = file_type or detect_file_type(path)
        if file_type in ('unsupported', 'error'):
            return False

        self.file_counter += 1
        item_id = f"file_{self.file_counter}"

        try:
            size = path.stat().st_size if size is None else size
        except OSError:
            size = 0

        item = FileItem(
            id=item_id,
            path=path,
            name=path.name,
            size=size,
            status='waiting',
            file_type=file_type,
            source_signature=file_signature(path)
        )

        self.files[item_id] = item
        self.path_keys.add(key)
        self.tree.insert('', END, iid=item_id, values=(
            item.name,
            TYPE_LABELS.get(file_type, file_type),
            f"{item.size_mb:.2f}MB",
            "等待中",
            ""
        ))
        self._mark_session_dirty()
        self.empty_hint.place_forget()
        return True

    def remove_selected(self):
        """删除选中项"""
        if self.is_converting or self.is_scanning or self.closing:
            return
        selected = self.tree.selection()
        for item_id in selected:
            if item_id in self.files:
                self.path_keys.discard(os.path.normcase(str(self.files[item_id].path)))
                del self.files[item_id]
                self.tree.delete(item_id)
        self.update_file_stats()
        self.update_convert_button()
        self._save_session()

    def clear_all(self):
        """清空所有"""
        if self.is_converting or self.is_scanning or self.closing:
            return
        if not self.files:
            return
        if messagebox.askyesno("确认", "确定要清空所有文件吗？"):
            self.files.clear()
            self.path_keys.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.update_file_stats()
            self.update_convert_button()
            self._save_session()

    def update_file_stats(self):
        """更新文件统计"""
        total = len(self.files)
        total_size = sum(f.size for f in self.files.values())
        self.file_stats_label.config(text=f"共{total}个文件，总大小{total_size/(1024*1024):.1f}MB")
        if total:
            self.empty_hint.place_forget()
        else:
            self.empty_hint.place(relx=0.5, rely=0.55, anchor=CENTER)

    def update_convert_button(self):
        """更新转换按钮"""
        if self.is_converting:
            return
        waiting_count = sum(1 for f in self.files.values() if f.status == 'waiting')
        enabled = waiting_count > 0 and not (self.is_scanning or self.closing)
        self.convert_btn.config(text=f"开始处理 ({waiting_count}个文件)",
                                state='normal' if enabled else 'disabled')

    def browse_output(self):
        """浏览输出目录"""
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            self.output_path_var.set(folder)

    def _update_quality_state(self):
        enabled = not (self.is_converting or self.is_scanning) and self.output_format_var.get() == 'mp3'
        for widget in self.quality_frame.winfo_children():
            widget.state(['!disabled'] if enabled else ['disabled'])

    def retry_failed(self):
        if self.is_converting or self.is_scanning or self.closing:
            return
        failed = [item for item in self.files.values() if item.status == 'error']
        if not failed:
            messagebox.showinfo('提示', '没有失败的文件')
            return
        for item in failed:
            item.status, item.error_msg, item.output_path = 'waiting', '', None
            item.warning_msg = ''
            item.validation_msg = ''
            self.update_item_status(item.id, '等待中', 'gray')
        self.update_convert_button()
        self.start_conversion(only_ids={item.id for item in failed})

    def on_context_menu(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        if item_id not in self.tree.selection():
            self.tree.selection_set(item_id)
        busy = self.is_converting or self.is_scanning or self.closing
        self.context_menu.entryconfigure(0, state='disabled' if busy else 'normal')
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def reprocess_selected(self):
        if self.is_converting or self.is_scanning or self.closing:
            return
        selected = [self.files[key] for key in self.tree.selection() if key in self.files]
        if not selected:
            return
        for item in selected:
            item.status, item.error_msg, item.warning_msg = 'waiting', '', ''
            item.validation_msg = ''
            self.update_item_status(item.id, '等待中', 'gray')
        self.update_convert_button()
        self.start_conversion(only_ids={item.id for item in selected})

    def show_selected_details(self):
        selected = self.tree.selection()
        if selected:
            self.show_file_details(selected[0])

    def show_file_details(self, item_id):
        if self.closing or item_id not in self.files:
            return
        item = replace(self.files[item_id])
        window = Toplevel(self.root)
        window.title(f'文件详情 - {item.name}')
        window.geometry('740x520')
        window.columnconfigure(0, weight=1)
        window.rowconfigure(0, weight=1)
        text = Text(window, wrap='word', padx=12, pady=12)
        scrollbar = ttk.Scrollbar(window, orient=VERTICAL, command=text.yview)
        text.config(yscrollcommand=scrollbar.set)
        text.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')
        text.insert('1.0', '正在读取文件信息…')
        text.config(state='disabled')
        buttons = ttk.Frame(window, padding=8)
        buttons.grid(row=1, column=0, columnspan=2, sticky='ew')
        ttk.Button(buttons, text='打开源目录', command=lambda: self._open_folder(item.path.parent)).pack(side=LEFT)
        output_button = ttk.Button(buttons, text='打开输出目录',
                                   command=lambda: self._open_folder(item.output_path.parent))
        output_button.pack(side=LEFT, padx=6)
        if item.output_path is None:
            output_button.state(['disabled'])
        def copy_paths():
            window.clipboard_clear()
            window.clipboard_append(str(item.path) + ('\n' + str(item.output_path) if item.output_path else ''))
        ttk.Button(buttons, text='复制路径', command=copy_paths).pack(side=LEFT)
        self._detail_counter += 1
        token = self._detail_counter
        self._detail_requests[token] = (window, text)
        def close():
            self._detail_requests.pop(token, None)
            window.destroy()
        window.protocol('WM_DELETE_WINDOW', close)
        def read_details():
            try:
                report = describe_item(item)
            except Exception as exc:
                report = f'无法读取文件详情: {exc}'
            self.events.put(('details', token, report))
        threading.Thread(target=read_details, daemon=True).start()

    def _mark_session_dirty(self, *_):
        self._session_dirty = True

    def show_environment_diagnostics(self):
        if self.is_converting or self.closing:
            messagebox.showinfo('提示', '请在本轮处理结束后运行诊断')
            return
        window = Toplevel(self.root)
        window.title('环境与兼容性诊断')
        window.geometry('760x480')
        text = Text(window, wrap='word', padx=12, pady=12)
        text.pack(fill=BOTH, expand=True)
        text.insert('1.0', '正在检查客户端和工具…')
        text.config(state='disabled')
        def copy_report():
            window.clipboard_clear()
            window.clipboard_append(text.get('1.0', 'end-1c'))
        ttk.Button(window, text='复制诊断信息', command=copy_report).pack(pady=8)
        self._detail_counter += 1
        token = self._detail_counter
        self._detail_requests[token] = (window, text)
        def close():
            self._detail_requests.pop(token, None)
            window.destroy()
        window.protocol('WM_DELETE_WINDOW', close)
        def inspect():
            try:
                report = environment_report(AudioConverter())
            except Exception as exc:
                report = f'诊断失败: {exc}'
            self.events.put(('details', token, report))
        threading.Thread(target=inspect, daemon=True).start()

    def _session_settings(self):
        return dict(output_path=self.output_path_var.get(),
                    output_format=self.output_format_var.get(), quality=self.quality_var.get(),
                    verify_full=self.verify_full_var.get())

    def _save_session(self):
        previous_error = self.session_store.last_error
        saved = self.session_store.save(list(self.files.values()), self._session_settings())
        if not saved:
            self.status_bar.config(text=self.session_store.last_error)
        elif previous_error and self.status_bar.cget('text') == previous_error:
            self.status_bar.config(text='任务记录已保存')
        self._session_dirty = not saved
        self._last_session_save = time.monotonic()
        return saved

    def _restore_session(self):
        items, settings = self.session_store.load()
        for key, variable in [('output_path', self.output_path_var),
                              ('output_format', self.output_format_var), ('quality', self.quality_var),
                              ('verify_full', self.verify_full_var)]:
            if key in settings:
                variable.set(settings[key])
        for item in items:
            self.file_counter += 1
            item.id = f'file_{self.file_counter}'
            self.files[item.id] = item
            self.path_keys.add(os.path.normcase(str(item.path)))
            self.tree.insert('', END, iid=item.id)
            label, color = {'waiting': ('等待中', 'gray'), 'done': ('已完成', 'green'),
                            'error': ('失败', 'red'), 'skipped': ('已跳过', 'gray')}[item.status]
            self.update_item_status(item.id, label, color)
        if items:
            self.update_file_stats()
            self.update_convert_button()
            self.status_label.config(text=f'已恢复 {len(items)} 个文件，可继续处理或右键重新处理')
        if self.session_store.last_error:
            self.status_bar.config(text=self.session_store.last_error)

    def open_output_folder(self):
        try:
            raw_path = self.output_path_var.get().strip()
            if not raw_path:
                raise ValueError('请先选择输出目录')
            path = Path(raw_path).expanduser().resolve()
        except (OSError, ValueError, RuntimeError) as exc:
            messagebox.showerror('无法打开目录', str(exc))
            return
        self._open_folder(path)

    def open_config_folder(self):
        try:
            directory = Config.CONFIG_FILE.parent
            directory.mkdir(parents=True, exist_ok=True)
            self._open_folder(directory.resolve())
        except (OSError, ValueError, RuntimeError) as exc:
            messagebox.showerror('无法打开配置目录', str(exc))

    def _open_folder(self, path):
        try:
            if not path.is_dir():
                raise OSError(f'目录不存在: {path}')
            os.startfile(str(path))
        except (OSError, ValueError, RuntimeError) as exc:
            messagebox.showerror('无法打开目录', str(exc))

    def on_item_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if self.tree.identify_column(event.x) != '#5':
            self._show_item(item_id)

    def on_action_click(self, event):
        if self.tree.identify_column(event.x) == '#5':
            self._show_item(self.tree.identify_row(event.y))

    def _show_item(self, item_id):
        item = self.files.get(item_id)
        if not item:
            return
        if item.error_msg or item.warning_msg:
            messagebox.showinfo(item.name, item.error_msg or item.warning_msg)
        else:
            self._open_folder((item.output_path or item.path).parent)

    def update_item_status(self, item_id: str, status_text: str, color: str):
        """更新列表项状态"""
        if item_id in self.files:
            item = self.files[item_id]
            self.tree.item(item_id, values=(
                item.name,
                TYPE_LABELS.get(item.file_type, item.file_type),
                f"{item.size_mb:.2f}MB",
                status_text,
                "查看原因" if item.error_msg else ("查看提示" if item.warning_msg else ("打开输出" if item.output_path else "源文件"))
            ), tags=(color,))
            self.tree.tag_configure('blue', foreground='blue')
            self.tree.tag_configure('green', foreground='green')
            self.tree.tag_configure('red', foreground='red')
            self.tree.tag_configure('gray', foreground='gray')
            self.tree.tag_configure('orange', foreground='#9a6700')

    # ========== 转换流程 ==========

    def _set_controls_enabled(self, enabled):
        def visit(widget):
            for child in widget.winfo_children():
                if isinstance(child, (ttk.Button, ttk.Entry, ttk.Radiobutton, ttk.Checkbutton)):
                    child.state(["!disabled"] if enabled else ["disabled"])
                visit(child)
        visit(self.toolbar)
        visit(self.settings_frame)
        self._update_quality_state()

    def _save_config(self):
        self.config.update(output_path=self.output_path_var.get(),
                           output_format=self.output_format_var.get(),
                           quality=self.quality_var.get(), verify_full=self.verify_full_var.get())
        if not Config.save(self.config):
            messagebox.showwarning('设置未保存', Config.last_error)

    def start_conversion(self, only_ids=None):
        if self.closing or self.is_scanning:
            return
        if self.is_converting:
            self.stop_event.set()
            self.convert_btn.config(text="当前文件完成后停止", state="disabled")
            self.status_label.config(text="正在停止：等待当前文件处理和清理完成")
            return

        waiting_files = [f for f in self.files.values() if f.status == "waiting"
                         and (only_ids is None or f.id in only_ids)]
        if not waiting_files:
            messagebox.showinfo("提示", "没有等待处理的文件")
            return
        raw_output = self.output_path_var.get().strip()
        if not raw_output:
            messagebox.showerror("错误", "请选择输出目录")
            return
        try:
            output_dir = Path(raw_output).expanduser()
            output_dir.mkdir(parents=True, exist_ok=True)
            output_dir = output_dir.resolve()
        except (OSError, ValueError, RuntimeError) as exc:
            messagebox.showerror("错误", f"无法使用输出目录，请检查路径: {exc}")
            return

        # 所有 Tk 变量均在主线程读取，运行中参数固定。
        output_format, quality = self.output_format_var.get(), self.quality_var.get()
        self.output_path_var.set(str(output_dir))
        self._save_config()
        self._save_session()
        self.converter._check_tools()
        verify_full = self.verify_full_var.get()
        if verify_full and not self.converter.ffmpeg_ok:
            messagebox.showerror('无法完整校验', '请将 ffmpeg.exe 放到程序同级目录后重试，或取消完整解码校验以使用基础校验。')
            return
        self.is_converting = True
        self.stop_event.clear()
        self._set_controls_enabled(False)
        self.convert_btn.config(text="停止处理", state="normal")
        self.progress_var.set(0)
        self.progress_bar.grid()
        self.stage_bar.config(mode='determinate', value=0)
        self.stage_bar.grid()
        self.worker_thread = threading.Thread(
            target=self.convert_worker,
            args=(waiting_files, output_dir, output_format, quality, verify_full),
            daemon=False)
        self.worker_thread.start()

    def convert_worker(self, files, output_dir, output_format, quality, verify_full=False):
        """只发送队列事件，不读取或调用 Tk。"""
        def on_status(item, text, color):
            self.events.put(("status", item.id if item else None, text, color))
        def on_progress(current, total):
            self.events.put(("progress", current, total))
        def on_stage(item, label, percent):
            self.events.put(('stage', item.id, label, percent))
        try:
            stats = self.batch_converter.process_batch(
                files, output_dir, output_format, quality,
                on_status=on_status, on_progress=on_progress,
                stop_check=self.stop_event.is_set, on_stage=on_stage, verify_full=verify_full)
        except Exception as exc:
            for item in files:
                if item.status == "converting":
                    item.status = "error"
                    item.error_msg = str(exc)
                    self.events.put(("status", item.id, "失败", "red"))
            stats = dict(done=sum(f.status == "done" for f in files),
                         error=sum(f.status == "error" for f in files),
                         skipped=sum(f.status == "skipped" for f in files),
                         fatal_error=str(exc))
        self.events.put(("finished", stats))

    def _poll_events(self):
        self._poll_id = None
        if self.instance_guard and self.instance_guard.activation_requested():
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        for _ in range(200):
            try:
                event = self.events.get_nowait()
            except Empty:
                break
            kind, *args = event
            if kind == "status":
                item_id, text, color = args
                if item_id is None:
                    self.status_label.config(text=text)
                else:
                    self.update_item_status(item_id, text, color)
                    if text == "处理中":
                        self._file_started = time.monotonic()
                        self._set_stage(item_id, '准备处理', None)
                    else:
                        self.stage_bar.stop()
                        self._active_stage = None
                    self._mark_session_dirty()
            elif kind == 'stage':
                self._set_stage(*args)
            elif kind == 'details':
                token, report = args
                request = self._detail_requests.pop(token, None)
                if request is not None:
                    window, text = request
                    if window.winfo_exists():
                        text.config(state='normal')
                        text.delete('1.0', END)
                        text.insert('1.0', report)
                        text.config(state='disabled')
            elif kind == "progress":
                current, total = args
                self.progress_var.set(current / total * 100 if total else 0)
                self.status_bar.config(text=f"已处理 {current}/{total}")
            elif kind == "scan_file":
                if not self.closing and self.add_file_item(*args):
                    self.scan_added += 1
            elif kind == "scan_error":
                if len(self.scan_errors) < 20:
                    self.scan_errors.append(args[0])
            elif kind == "scan_finished":
                self._scan_result = True
            elif kind == "finished":
                self._pending_stats = args[0]

        if self.is_scanning:
            self.update_file_stats()
            self.status_label.config(text=f"正在扫描，已添加 {self.scan_added} 个文件…")
        if self._scan_result and not self.scan_thread.is_alive():
            self._scan_result = None
            self.is_scanning = False
            if self.closing:
                self._destroy()
                return
            self._set_controls_enabled(True)
            self.convert_btn.config(state="normal")
            self.update_convert_button()
            self.status_label.config(text=f"已添加 {self.scan_added} 个文件")
            self._save_session()
            if self.scan_errors:
                messagebox.showwarning("部分文件无法添加（最多显示 20 条）", "\n".join(self.scan_errors))
        if self._pending_stats is not None and not self.worker_thread.is_alive():
            stats, self._pending_stats = self._pending_stats, None
            self.is_converting = False
            if self.closing:
                self._destroy()
                return
            self.conversion_finished(stats)
        self._refresh_stage_label()
        if self._session_dirty and time.monotonic() - self._last_session_save >= 1:
            self._save_session()
        self._poll_id = self.root.after(75, self._poll_events)

    def _set_stage(self, item_id, label, percent):
        changed = self._active_stage is None or self._active_stage[:2] != (item_id, label)
        was_unknown = self._active_stage is None or self._active_stage[2] is None
        self._active_stage = (item_id, label, percent)
        if percent is None:
            if changed:
                self.stage_bar.stop()
                self.stage_bar.config(mode='indeterminate', value=0)
                self.stage_bar.start(25)
        else:
            if changed or was_unknown:
                self.stage_bar.stop()
                self.stage_bar.config(mode='determinate')
            self.stage_bar['value'] = percent

    def _refresh_stage_label(self):
        if self._active_stage is None or self._file_started is None:
            return
        item_id, stage, percent = self._active_stage
        item = self.files.get(item_id)
        if item is None:
            return
        elapsed = int(time.monotonic() - self._file_started)
        text = f'{stage}: {item.name} · 已用 {elapsed // 60:02d}:{elapsed % 60:02d}'
        if percent is not None:
            text += f' · 当前阶段 {percent}%'
        if self.stop_event.is_set():
            text += ' · 完成当前文件后退出' if self.closing else ' · 完成当前文件后停止'
        self.status_label.config(text=text)

    def conversion_finished(self, stats: Dict):
        self.stage_bar.stop()
        self.progress_bar.grid_remove()
        self.stage_bar.grid_remove()
        self._active_stage = None
        self._save_session()
        self._set_controls_enabled(True)
        self.convert_btn.config(state="normal")
        self.update_convert_button()
        title = "已停止" if stats.get("stopped") else "处理完成"
        if stats.get("fatal_error"):
            title = "处理异常"
        self.status_label.config(text=title)
        self.status_bar.config(text=self.session_store.last_error or "就绪")
        lines = [title, f"成功: {stats.get('done', 0)} 个",
                 f"跳过: {stats.get('skipped', 0)} 个",
                 f"失败: {stats.get('error', 0)} 个"]
        if stats.get("stopped"):
            lines.append(f"未处理: {stats.get('remaining', 0)} 个，可继续开始处理")
        if stats.get('warnings'):
            lines.append(f"{stats['warnings']} 个文件的标签/封面未能保留，点击列表中的“查看提示”查看原因")
        if stats.get("error"):
            if stats.get("log_error"):
                lines.append(stats["log_error"])
            elif stats.get('log_path'):
                lines.append(f"错误日志: {stats['log_path']}")
        if stats.get("fatal_error"):
            lines.append(stats["fatal_error"])
        show = messagebox.showwarning if stats.get("error") or stats.get("fatal_error") else messagebox.showinfo
        show(title, "\n".join(lines))

    def on_close(self):
        if self.closing:
            return
        if self.is_scanning:
            self.closing = True
            self.scan_stop.set()
            self.status_label.config(text="正在结束扫描…")
            return
        if self.is_converting:
            if not messagebox.askyesno("退出", "将在当前文件完成并清理后退出，确定吗？"):
                return
            self.closing = True
            self.stop_event.set()
            self.convert_btn.config(text="清理完成后退出", state="disabled")
            self.status_label.config(text="等待当前文件结束，请稍候…")
            return
        self._destroy()

    def _destroy(self):
        if self._poll_id is not None:
            self.root.after_cancel(self._poll_id)
            self._poll_id = None
        self._save_config()
        if not self._save_session():
            messagebox.showwarning('任务记录未保存', self.session_store.last_error)
        self.stage_bar.stop()
        self._detail_requests.clear()
        self.root.destroy()


def create_root():
    try:
        from tkinterdnd2 import TkinterDnD
    except ImportError:
        root = Tk()
        root.dnd_available = False
        return root
    root = TkinterDnD.Tk()
    root.dnd_available = True
    return root


def main():
    from single_instance import SingleInstance
    guard = SingleInstance()
    try:
        if not guard.acquire():
            if guard.notify_existing():
                return
            # 首个实例可能在通知期间退出；重新竞争一次，避免启动请求丢失。
            if not guard.acquire():
                root = Tk()
                root.withdraw()
                messagebox.showerror('程序已在运行', '已有实例暂时无法唤起，请等待其启动完成后重试。')
                root.destroy()
                return
        root = create_root()
        ConverterGUI(root, instance_guard=guard)
        root.mainloop()
    finally:
        guard.close()


if __name__ == "__main__":
    main()
