### 源代码目录 (Source Code)

#### 模块结构

| 文件 | 职责 |
|------|------|
| `models.py` | 数据模型 (`FileItem`, `Config`)、文件类型检测 (`detect_file_type`)、常量 |
| `decryptor.py` | `QQMusicDecryptor` — FRIDA Hook 解密器 |
| `converter.py` | `AudioConverter` — 音频工具链封装 (flac/oggdec/lame/ffmpeg) |
| `worker.py` | `BatchConverter` — 纯逻辑的批量转换工作器，不依赖 GUI |
| `gui.py` | `ConverterGUI` + `main()` — tkinter 界面和事件处理 |
| `qqmusic_converter_gui.py` | 入口文件，兼容旧运行方式 |

#### 实现原理

**加密文件处理**：通过 FRIDA attach 到 `QQMusic.exe`，Hook `QQMusicCommon.dll` 的 `EncAndDesMediaFile` 类，让 QQ 音乐自身完成解密。

**音频转换工具链**：

| 步骤 | 工具 | 说明 |
|------|------|------|
| 解密 | FRIDA + `hook_qqmusic_v17.js` | `.mflac` → `.flac` / `.mgg` → `.ogg` / `.mmp4` → `.mp4` |
| 解码 | `flac.exe` / `oggdec.exe` | FLAC/OGG → WAV（`oggdec` 需临时复制 workaround 空格路径） |
| 备选 | `ffmpeg.exe` | `flac.exe` 遇到损坏帧失败时自动 fallback |
| 编码 | `lame.exe` | WAV → MP3（VBR / 320k / 192k） |
| 清理 | - | 转换完成后自动删除中间临时文件 |

#### 支持格式

| 输入 | 解码 | 转 MP3 |
|------|-----------|--------|
| `.mflac` | → `.flac` | ✅ |
| `.mgg` | → `.ogg` | ✅ |
| `.mmp4` | → `.mp4` | ❌ 不支持（E-AC3） |
| `.flac` | 直接复制 | ✅ |
| `.ogg` | 直接复制 | ✅ |
| `.wav` | 直接复制 | ✅ |
| `.mp3` | 直接复制 | 直接复制 |

#### 运行要求

- Python 3.8+
- `tkinter`（GUI 库）
- `frida`（`pip install frida`）— 仅处理加密文件时需要
- 处理加密文件前**必须先启动 QQ 音乐客户端**

#### 路径兼容性处理

- **FRIDA 解密**：必须使用绝对路径（`Path.resolve()`），因 Hook 在 QQ 音乐进程内执行
- **FRIDA 写文件**：使用纯 ASCII 临时路径（`TEMP/qqmusic_decrypt/`），避免中文路径问题
- **oggdec**：不支持空格路径，临时复制到无空格路径后再解码
- **flac**：支持中文路径；遇到损坏帧自动 fallback `ffmpeg.exe`（需手动放置到同级目录）
