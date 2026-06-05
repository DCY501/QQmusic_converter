#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主 GUI 界面
"""

import sys
import os
import threading
from pathlib import Path
from datetime import datetime
from typing import Dict
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import warnings

warnings.filterwarnings("ignore", "libpng warning")
warnings.filterwarnings("ignore", "iCCP")

# Windows UTF-8 编码支持
if sys.platform == 'win32' and sys.stdout is not None:
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

from models import FileItem, Config, detect_file_type, TYPE_LABELS, APP_DIR
from decryptor import QQMusicDecryptor
from converter import AudioConverter
from worker import BatchConverter


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
        self.batch_converter = BatchConverter(self.decryptor, self.converter)

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
        toolbar.grid(row=0, column=0, sticky=(W, E), pady=(0, 10))

        ttk.Button(toolbar, text="📁 添加文件", command=self.add_files).pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="📂 添加文件夹", command=self.add_folder).pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="🗑️ 删除", command=self.remove_selected).pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="🧹 清空", command=self.clear_all).pack(side=LEFT)

    def _create_file_list(self, parent):
        """创建文件列表"""
        list_frame = ttk.LabelFrame(parent, text="文件列表", padding="5")
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

    def _create_settings(self, parent):
        """创建输出设置面板"""
        settings_frame = ttk.LabelFrame(parent, text="输出设置", padding="5")
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
        quality_frame.grid(row=2, column=1, sticky=W, pady=(10, 0))
        ttk.Radiobutton(quality_frame, text="VBR高质量", variable=self.quality_var, value='vbr').pack(side=LEFT, padx=(0, 15))
        ttk.Radiobutton(quality_frame, text="CBR 320k", variable=self.quality_var, value='320k').pack(side=LEFT, padx=(0, 15))
        ttk.Radiobutton(quality_frame, text="CBR 192k", variable=self.quality_var, value='192k').pack(side=LEFT)

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

        self.status_label = ttk.Label(convert_frame, text="就绪")
        self.status_label.grid(row=2, column=0, sticky=W)

        # 状态栏
        self.status_bar = ttk.Label(parent, text="检查工具中...", relief=SUNKEN, anchor=W)
        self.status_bar.grid(row=5, column=0, sticky=(W, E), pady=(10, 0))

    # ========== 文件管理 ==========

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
            TYPE_LABELS.get(file_type, file_type),
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
            self.tree.item(item_id, values=(
                item.name,
                TYPE_LABELS.get(item.file_type, item.file_type),
                f"{item.size_mb:.2f}MB",
                status_text,
                ""
            ), tags=(color,))
            self.tree.tag_configure('blue', foreground='blue')
            self.tree.tag_configure('green', foreground='green')
            self.tree.tag_configure('red', foreground='red')
            self.tree.tag_configure('gray', foreground='gray')

    # ========== 转换流程 ==========

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

    def convert_worker(self, files, output_dir: Path):
        """后台处理线程 - 委托给 BatchConverter"""
        output_format = self.output_format_var.get()
        quality = self.quality_var.get()

        def on_status(file_item, status_text, color):
            if file_item is None:
                self.root.after(0, self.status_label.config, {'text': status_text})
            else:
                self.root.after(0, self.update_item_status, file_item.id, status_text, color)
                if status_text in ('处理中', '解密中', '转码中'):
                    self.root.after(0, self.status_label.config, {'text': f"{status_text}: {file_item.name}"})

        def on_progress(current, total):
            progress = (current / total) * 100
            self.root.after(0, self.progress_var.set, progress)
            if current < total:
                name = files[current - 1].name if current > 0 else ""
                self.root.after(0, self.status_bar.config, {'text': f"处理中 ({current}/{total}): {name}"})

        def on_error(file_item, error_msg):
            self.root.after(0, self.log_error, file_item)

        def stop_check():
            return self.stop_flag

        stats = self.batch_converter.process_batch(
            files, output_dir, output_format, quality,
            on_status=on_status,
            on_progress=on_progress,
            on_error=on_error,
            stop_check=stop_check
        )

        self.is_converting = False
        self.root.after(0, self.conversion_finished, stats)

    def conversion_finished(self, stats: Dict):
        """处理完成回调"""
        waiting = len([f for f in self.files.values() if f.status == 'waiting'])
        self.convert_btn.config(text=f"开始处理 ({waiting}个文件)")
        self.status_label.config(text="处理完成")
        self.status_bar.config(text="就绪")

        done_count = stats.get('done', 0)
        error_count = stats.get('error', 0)
        skip_count = stats.get('skipped', 0)

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
