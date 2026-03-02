### 源代码目录 (Source Code)

#### 实现原理

使用独立的音频编解码工具链完成格式转换：

1. **解码**：`oggdec` 将 OGG 格式解码为标准的 WAV 音频文件
2. **编码**：`lame` 将 WAV 音频压缩编码为 MP3 格式
3. **清理**：转换完成后自动删除中间临时 WAV 文件

#### 文件说明

- `main.py` - 主程序，提供 GUI 界面和转换流程控制
- 依赖  `oggdec.exe` 和 `lame.exe`

#### 运行要求

- Python 3.8+
- tkinter（GUI 库）
- oggdec & lame（音频工具）
