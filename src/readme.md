### 源代码目录 (Source Code)

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

#### 文件说明

- `qqmusic_converter_gui.py` - 主程序，提供 GUI 界面、FRIDA 解密管理和转换流程控制
- `hook_qqmusic_v17.js` - FRIDA 17.x 兼容的 Hook 脚本
- `flac.exe` + `libFLAC.dll` + `libFLAC++.dll` - FLAC 解码器
- `oggdec.exe` - OGG Vorbis 解码器
- `lame.exe` + `lame_enc.dll` - MP3 编码器
- `config.json` - 用户配置（输出路径、格式、音质）

#### 支持格式

| 输入 | 保持原格式 | 转 MP3 |
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
