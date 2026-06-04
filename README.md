### 🎵 QQ音乐解密/转码工具 —— 个人学习，请勿商用

支持 QQ 音乐加密格式（`.mflac` / `.mgg` / `.mmp4`）解密，以及标准音频格式（`.flac` / `.ogg` / `.wav` / `.mp3`）转码。

---

#### ✨ 已实现功能

| 输入格式 | 保持原格式 | 转 MP3 | 说明 |
|---------|-----------|--------|------|
| `.mflac` | → `.flac` | ✅ | FRIDA Hook 解密后转码 |
| `.mgg` | → `.ogg` | ✅ | FRIDA Hook 解密后转码 |
| `.mmp4` | → `.mp4` | ❌ 不支持 | E-AC3 音频，无法转 MP3 |
| `.flac` | 直接复制 | ✅ | `flac.exe` 解码，失败时 fallback `ffmpeg` |
| `.ogg` | 直接复制 | ✅ | `oggdec.exe` 解码 |
| `.wav` | 直接复制 | ✅ | 直接编码 |
| `.mp3` | 直接复制 | 直接复制 | 无需处理 |

- 🚀 **批量处理**：支持文件/文件夹拖拽添加，批量转换
- 🎚️ **音质选择**：VBR 高质量 / CBR 320k / CBR 192k
- 💻 **图形界面**：基于 tkinter，无需命令行
- 🔓 **FRIDA 解密**：Hook QQ 音乐进程实现 `.mflac` / `.mgg` / `.mmp4` 解密
- 💾 **配置记忆**：自动保存输出目录、格式和音质设置
- 🔒 **本地处理**：不上传任何服务器

---

#### ⚠️ 重要使用提示

**使用 FRIDA 解密功能前，必须先启动 QQ 音乐客户端！**

程序通过 FRIDA attach 到 `QQMusic.exe` 进程，调用其内部的 `EncAndDesMediaFile` 类完成解密。如果 QQ 音乐未运行，加密文件将无法处理。

> 注意：FRIDA Hook 仅对当前运行的 QQ 音乐客户端所下载的文件有效。如果文件是其他版本/其他账号下载的，可能无法解密。

---

#### 📦 运行方式

##### 方式一：Python 源码运行

```bash
pip install frida
python src/qqmusic_converter_gui.py
```

##### 方式二：打包为 EXE（推荐分发）

```bash
pip install pyinstaller
pyinstaller -F -w src/qqmusic_converter_gui.py --add-data "src/flac.exe;src" --add-data "src/oggdec.exe;src" --add-data "src/lame.exe;src" --add-data "src/lame_enc.dll;src" --add-data "src/libFLAC.dll;src" --add-data "src/libFLAC++.dll;src" --add-data "src/hook_qqmusic_v17.js;src"
```

---

#### 🛠️ 关于 `ffmpeg.exe` 是否打包

| 方案 | 体积 | 效果 |
|------|------|------|
| **不打包 ffmpeg** | `src/` 目录约 **3.5MB** | 能处理 95%+ 的正常 FLAC，足够日常使用 |
| **打包 ffmpeg** | 加上 ffmpeg 约 **90MB** | 能处理损坏/特殊的 FLAC（极少数情况） |

**建议：不打包 ffmpeg**。QQ 音乐解密出来的 `.mflac` 基本都是正常 FLAC，`flac.exe` 足够用。ffmpeg 只是"保险"，遇到极少数损坏文件时才 fallback。

> 如果你之后遇到某首歌 `flac.exe` 解码失败，手动把 `ffmpeg.exe` 放到程序同级目录即可自动生效。

---

#### 🙏 致谢与参考

- **FRIDA Hook 思路**：借鉴自 [decrypt-mflac-frida](https://github.com/Noisrev/decrypt-mflac-frida) 项目，通过 Hook `QQMusicCommon.dll` 的 `EncAndDesMediaFile` 类实现解密
- **FRIDA 17.x API 适配**：原项目使用旧版 FRIDA API（`Module.findExportByName`），本项目已迁移到 frida 17.x 的 `Process.findModuleByName(...).findExportByName(...)`
- **音频工具链**：
  - `flac.exe` / `libFLAC*.dll` — [Xiph.Org FLAC](https://xiph.org/flac/)
  - `oggdec.exe` — [Vorbis Tools](https://xiph.org/downloads/)
  - `lame.exe` / `lame_enc.dll` — [LAME MP3 Encoder](https://lame.sourceforge.io/)
  - `ffmpeg.exe` — [FFmpeg](https://ffmpeg.org/)（作为 flac 解码失败的备选方案）

---

#### 📝 本次更新解决的问题

1. **FRIDA 路径问题**：Hook 在 QQ 音乐进程内执行，必须使用绝对路径，已修复
2. **oggdec 空格路径**：`oggdec.exe` 无法处理含空格的路径，已添加临时复制 workaround
3. **flac 容错解码**：某些 FLAC 文件帧损坏导致 `flac.exe` 失败，已添加 `ffmpeg` fallback
4. **临时文件冲突**：同名不同扩展名文件解密时临时文件互相覆盖，已使用唯一命名
5. **`.mmp4` 状态显示**：转 MP3 时不应显示"失败"，已改为"不支持转MP3"
