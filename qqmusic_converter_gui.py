#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QQ音乐OGG转MP3工具 - GUI版本

使用方法:
    python qqmusic_converter_gui.py
"""

import sys
import os
import subprocess
import threading
import json
from pathlib import Path
from datetime import datetime
from typing import List, Optional, Dict
from dataclasses import dataclass, asdict
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import warnings

# 忽略PNG色彩配置警告（双重保险）
warnings.filterwarnings("ignore", "libpng warning")
warnings.filterwarnings("ignore", "iCCP")


def get_app_dir() -> Path:
    """获取程序所在目录（兼容开发环境和打包环境）"""
    if getattr(sys, 'frozen', False):
        # PyInstaller打包后：使用EXE所在目录
        return Path(sys.executable).parent
    else:
        # 开发环境：使用脚本所在目录
        return Path(__file__).parent

# Windows UTF-8 编码支持（仅在控制台模式下）
if sys.platform == 'win32' and sys.stdout is not None:
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


@dataclass
class FileItem:
    """文件项数据类"""
    id: str
    path: Path
    name: str
    size: int
    status: str  # 'waiting', 'converting', 'done', 'error'
    error_msg: str = ""
    
    @property
    def size_mb(self) -> float:
        return self.size / (1024 * 1024)


class Config:
    """配置管理"""
    CONFIG_FILE = get_app_dir() / "config.json"
    
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


def find_tools() -> tuple[Optional[str], Optional[str]]:
    """查找oggdec和lame可执行文件"""
    app_dir = get_app_dir().resolve()
    
    oggdec_path = app_dir / 'oggdec.exe'
    lame_path = app_dir / 'lame.exe'
    
    oggdec_found = False
    lame_found = False
    
    # 检查oggdec
    try:
        result = subprocess.run([str(oggdec_path)], capture_output=True, timeout=5, creationflags=subprocess.CREATE_NO_WINDOW)
        # oggdec无参数时会返回错误码，但说明文件存在
        oggdec_found = oggdec_path.exists()
    except:
        pass
    
    # 检查lame
    try:
        result = subprocess.run([str(lame_path), '--version'], capture_output=True, timeout=5, creationflags=subprocess.CREATE_NO_WINDOW)
        if result.returncode == 0:
            lame_found = True
    except:
        pass
    
    return (str(oggdec_path) if oggdec_found else None, 
            str(lame_path) if lame_found else None)


def detect_file_type(file_path: Path) -> str:
    """检测文件类型"""
    try:
        with open(file_path, 'rb') as f:
            header = f.read(8)
        if len(header) >= 4 and header[:4] == b'OggS':
            return 'standard_ogg'
        return 'unsupported'
    except:
        return 'error'


class ConverterGUI:
    """主GUI类"""
    
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("QQ音乐OGG转MP3工具")
        self.root.geometry("800x550")
        self.root.minsize(600, 400)
        
        # 居中窗口
        self.center_window()
        
        # ===== 方案二：使用clam主题（不加载系统PNG） =====
        self.style = ttk.Style()
        # 使用clam主题避免加载系统主题的PNG资源
        self.style.theme_use('clam')
        
        # 配置clam主题颜色（美化）
        self.style.configure('TFrame', background='#f5f6f7')
        self.style.configure('TLabel', background='#f5f6f7', font=('Microsoft YaHei', 10))
        self.style.configure('TButton', font=('Microsoft YaHei', 9))
        self.style.configure('TLabelframe', background='#f5f6f7')
        self.style.configure('TLabelframe.Label', background='#f5f6f7', font=('Microsoft YaHei', 9, 'bold'))
        self.style.configure('Treeview', font=('Microsoft YaHei', 9))
        self.style.configure('Treeview.Heading', font=('Microsoft YaHei', 9, 'bold'))
        
        # 设置根窗口背景
        self.root.configure(bg='#f5f6f7')
        # ====================================================
        
        self.default_font = ('Microsoft YaHei', 10)
        self.root.option_add('*Font', self.default_font)
        
        # 数据
        self.files: Dict[str, FileItem] = {}
        self.file_counter = 0
        self.is_converting = False
        self.stop_flag = False
        self.oggdec_path, self.lame_path = find_tools()
        
        # 加载配置
        self.config = Config.load()
        
        # 创建UI
        self.create_ui()
        
        # 检查工具
        if not self.oggdec_path or not self.lame_path:
            messagebox.showerror("错误", "需要oggdec.exe和lame.exe放在程序目录")
        
        # 设置关闭处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def center_window(self):
        """窗口居中"""
        self.root.update_idletasks()
        width = 800
        height = 550
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_ui(self):
        """创建界面"""
        # 主框架 - 使用clam主题背景色
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(N, S, E, W))
        
        # 配置grid权重
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
        
        # Treeview
        columns = ('name', 'size', 'status', 'action')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', selectmode='extended')
        
        self.tree.heading('name', text='文件名')
        self.tree.heading('size', text='大小')
        self.tree.heading('status', text='状态')
        self.tree.heading('action', text='操作')
        
        self.tree.column('name', width=300)
        self.tree.column('size', width=80, anchor=CENTER)
        self.tree.column('status', width=100, anchor=CENTER)
        self.tree.column('action', width=80, anchor=CENTER)
        
        # 滚动条
        vsb = ttk.Scrollbar(list_frame, orient=VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(list_frame, orient=HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky=(N, S, E, W))
        vsb.grid(row=0, column=1, sticky=(N, S))
        hsb.grid(row=1, column=0, sticky=(E, W))
        
        # 绑定事件
        self.tree.bind('<Delete>', lambda e: self.remove_selected())
        self.tree.bind('<Double-1>', self.on_item_double_click)
        self.tree.bind('<Button-3>', self.show_context_menu)
        
        # 文件统计
        self.file_stats_label = ttk.Label(list_frame, text="共0个文件，总大小0MB")
        self.file_stats_label.grid(row=2, column=0, columnspan=2, sticky=W, pady=(5, 0))
        
        # ===== 输出设置 =====
        settings_frame = ttk.LabelFrame(main_frame, text="输出设置", padding="5")
        settings_frame.grid(row=3, column=0, sticky=(W, E), pady=(0, 10))
        settings_frame.columnconfigure(1, weight=1)
        
        # 输出路径
        ttk.Label(settings_frame, text="保存到:").grid(row=0, column=0, sticky=W, padx=(0, 5))
        
        default_output = self.config.get('output_path', str((get_app_dir() / "转化").resolve()))
        self.output_path_var = StringVar(value=default_output)
        ttk.Entry(settings_frame, textvariable=self.output_path_var).grid(row=0, column=1, sticky=(W, E), padx=(0, 5))
        ttk.Button(settings_frame, text="浏览...", command=self.browse_output).grid(row=0, column=2)
        
        # 音质选项
        ttk.Label(settings_frame, text="音质:").grid(row=1, column=0, sticky=W, padx=(0, 5), pady=(10, 0))
        
        self.quality_var = StringVar(value=self.config.get('quality', 'vbr'))
        quality_frame = ttk.Frame(settings_frame)
        quality_frame.grid(row=1, column=1, sticky=W, pady=(10, 0))
        
        ttk.Radiobutton(quality_frame, text="VBR高质量", variable=self.quality_var, value='vbr').pack(side=LEFT, padx=(0, 15))
        ttk.Radiobutton(quality_frame, text="CBR 320k", variable=self.quality_var, value='320k').pack(side=LEFT, padx=(0, 15))
        ttk.Radiobutton(quality_frame, text="CBR 192k", variable=self.quality_var, value='192k').pack(side=LEFT)
        
        # ===== 转换区域 =====
        convert_frame = ttk.Frame(main_frame)
        convert_frame.grid(row=4, column=0, sticky=(W, E))
        convert_frame.columnconfigure(0, weight=1)
        
        self.convert_btn = ttk.Button(convert_frame, text="开始转换 (0个文件)", command=self.start_conversion)
        self.convert_btn.grid(row=0, column=0, sticky=(W, E), pady=(0, 5))
        
        # 进度条
        self.progress_var = DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(convert_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=1, column=0, sticky=(W, E), pady=(0, 5))
        
        self.status_label = ttk.Label(convert_frame, text="就绪")
        self.status_label.grid(row=2, column=0, sticky=W)
        
        # 状态栏
        if self.oggdec_path and self.lame_path:
            status_text = "就绪 | oggdec+lame已就绪"
        else:
            status_text = "未找到oggdec.exe或lame.exe"
        self.status_bar = ttk.Label(main_frame, text=status_text, relief=SUNKEN, anchor=W)
        self.status_bar.grid(row=5, column=0, sticky=(W, E), pady=(10, 0))
        
        # 右键菜单
        self.context_menu = Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="打开所在位置", command=self.open_file_location)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="从列表移除", command=self.remove_selected)
    
    def add_files(self):
        """添加文件"""
        files = filedialog.askopenfilenames(
            title="选择OGG文件",
            filetypes=[("OGG音频", "*.ogg"), ("所有文件", "*.*")]
        )
        added = 0
        for f in files:
            path = Path(f)
            if path.suffix.lower() != '.ogg':
                continue
            if self.add_file_item(path):
                added += 1
        
        if added > 0:
            self.update_file_stats()
            self.update_convert_button()
    
    def add_folder(self):
        """添加文件夹"""
        folder = filedialog.askdirectory(title="选择包含OGG文件的文件夹")
        if not folder:
            return
        
        folder_path = Path(folder)
        ogg_files = list(folder_path.rglob("*.ogg"))
        
        added = 0
        for path in ogg_files:
            if self.add_file_item(path):
                added += 1
        
        if added > 0:
            self.update_file_stats()
            self.update_convert_button()
            messagebox.showinfo("提示", f"已添加 {added} 个文件")
        elif ogg_files:
            messagebox.showinfo("提示", "所有文件已在列表中")
        else:
            messagebox.showinfo("提示", "未找到OGG文件")
    
    def add_file_item(self, path: Path) -> bool:
        """添加单个文件到列表"""
        # 检查是否已存在
        for item in self.files.values():
            if item.path == path:
                return False
        
        # 检测文件类型
        file_type = detect_file_type(path)
        if file_type != 'standard_ogg':
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
            status='waiting'
        )
        
        self.files[item_id] = item
        
        # 插入到Treeview
        self.tree.insert('', END, iid=item_id, values=(
            item.name,
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
        self.convert_btn.config(text=f"开始转换 ({waiting_count}个文件)")
    
    def browse_output(self):
        """浏览输出目录"""
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            self.output_path_var.set(folder)
    
    def on_item_double_click(self, event):
        """双击打开文件位置"""
        self.open_file_location()
    
    def show_context_menu(self, event):
        """显示右键菜单"""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def open_file_location(self):
        """打开文件所在位置"""
        selected = self.tree.selection()
        if not selected:
            return
        
        item_id = selected[0]
        if item_id in self.files:
            path = self.files[item_id].path
            folder = str(path.parent)
            if os.path.exists(folder):
                os.startfile(folder)
    
    def start_conversion(self):
        """开始转换"""
        if self.is_converting:
            self.stop_flag = True
            self.convert_btn.config(text="正在停止...")
            return
        
        if not self.oggdec_path or not self.lame_path:
            messagebox.showerror("错误", "未找到oggdec.exe或lame.exe！")
            return
        
        waiting_files = [f for f in self.files.values() if f.status == 'waiting']
        if not waiting_files:
            messagebox.showinfo("提示", "没有等待转换的文件")
            return
        
        # 检查输出目录
        output_dir = Path(self.output_path_var.get())
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出目录: {e}")
            return
        
        # 保存配置
        self.config['output_path'] = str(output_dir)
        self.config['quality'] = self.quality_var.get()
        Config.save(self.config)
        
        # 启动转换线程
        self.is_converting = True
        self.stop_flag = False
        self.convert_btn.config(text="停止转换")
        self.progress_var.set(0)
        
        thread = threading.Thread(target=self.convert_worker, args=(waiting_files, output_dir))
        thread.daemon = True
        thread.start()
    
    def convert_worker(self, files: List[FileItem], output_dir: Path):
        """后台转换线程"""
        total = len(files)
        quality = self.quality_var.get()
        
        for i, file_item in enumerate(files):
            if self.stop_flag:
                break
            
            # 更新状态
            file_item.status = 'converting'
            self.root.after(0, self.update_item_status, file_item.id, '转换中', 'blue')
            self.root.after(0, self.status_label.config, {'text': f"正在转换: {file_item.name}..."})
            self.root.after(0, self.status_bar.config, {'text': f"转换中 ({i+1}/{total}): {file_item.name}"})
            
            # 执行转换
            temp_wav = None
            try:
                output_path = output_dir / file_item.path.with_suffix('.mp3').name
                
                # 检查是否已存在
                if output_path.exists():
                    pass
                
                # 步骤1: OGG解码为WAV
                temp_wav = output_dir / (file_item.path.stem + '_temp.wav')
                decode_cmd = [self.oggdec_path, str(file_item.path), '-o', str(temp_wav)]
                
                decode_result = subprocess.run(decode_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW)
                
                if decode_result.returncode != 0 or not temp_wav.exists():
                    file_item.status = 'error'
                    file_item.error_msg = f"OGG解码失败: {decode_result.stderr[:100] if decode_result.stderr else '未知错误'}"
                    self.root.after(0, self.update_item_status, file_item.id, '失败', 'red')
                    self.log_error(file_item)
                    continue
                
                # 步骤2: WAV编码为MP3
                if quality == 'vbr':
                    encode_cmd = [self.lame_path, '-V', '0', str(temp_wav), str(output_path)]
                elif quality == '320k':
                    encode_cmd = [self.lame_path, '-b', '320', str(temp_wav), str(output_path)]
                else:  # 192k
                    encode_cmd = [self.lame_path, '-b', '192', str(temp_wav), str(output_path)]
                
                encode_result = subprocess.run(encode_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW)
                
                # 步骤3: 删除临时WAV文件
                try:
                    if temp_wav.exists():
                        temp_wav.unlink()
                except:
                    pass
                
                if encode_result.returncode == 0 and output_path.exists():
                    file_item.status = 'done'
                    self.root.after(0, self.update_item_status, file_item.id, '已完成', 'green')
                else:
                    file_item.status = 'error'
                    file_item.error_msg = f"MP3编码失败: {encode_result.stderr[:100] if encode_result.stderr else '未知错误'}"
                    self.root.after(0, self.update_item_status, file_item.id, '失败', 'red')
                    self.log_error(file_item)
                    
            except Exception as e:
                # 清理临时文件
                if temp_wav and temp_wav.exists():
                    try:
                        temp_wav.unlink()
                    except:
                        pass
                file_item.status = 'error'
                file_item.error_msg = str(e)
                self.root.after(0, self.update_item_status, file_item.id, '失败', 'red')
                self.log_error(file_item)
            
            # 更新进度
            progress = ((i + 1) / total) * 100
            self.root.after(0, self.progress_var.set, progress)
        
        # 完成
        self.is_converting = False
        self.root.after(0, self.conversion_finished, total)
    
    def update_item_status(self, item_id: str, status_text: str, color: str):
        """更新列表项状态"""
        if item_id in self.files:
            item = self.files[item_id]
            self.tree.item(item_id, values=(
                item.name,
                f"{item.size_mb:.2f}MB",
                status_text,
                ""
            ), tags=(color,))
            
            # 设置标签颜色
            self.tree.tag_configure('blue', foreground='blue')
            self.tree.tag_configure('green', foreground='green')
            self.tree.tag_configure('red', foreground='red')
            self.tree.tag_configure('gray', foreground='gray')
    
    def conversion_finished(self, total: int):
        """转换完成回调"""
        self.convert_btn.config(text=f"开始转换 ({len([f for f in self.files.values() if f.status == 'waiting'])}个文件)")
        self.status_label.config(text="转换完成")
        self.status_bar.config(text="就绪")
        
        done_count = sum(1 for f in self.files.values() if f.status == 'done')
        error_count = sum(1 for f in self.files.values() if f.status == 'error')
        
        if error_count > 0:
            messagebox.showwarning("完成", f"转换完成！\n成功: {done_count} 个\n失败: {error_count} 个\n\n失败记录已保存到 error.log")
        else:
            messagebox.showinfo("完成", f"转换完成！\n成功转换 {done_count} 个文件")
        
        self.update_convert_button()
    
    def log_error(self, item: FileItem):
        """记录错误日志"""
        try:
            log_path = get_app_dir() / "error.log"
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {item.path}\n")
                f.write(f"  错误: {item.error_msg}\n\n")
        except:
            pass
    
    def on_close(self):
        """关闭窗口"""
        if self.is_converting:
            if not messagebox.askyesno("确认", "转换正在进行中，确定要退出吗？"):
                return
            self.stop_flag = True
        
        # 保存配置
        self.config['output_path'] = self.output_path_var.get()
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