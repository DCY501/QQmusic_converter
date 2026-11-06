# 第三方组件

本项目的 MIT License 仅适用于本项目自身代码。运行环境、第三方库及音频工具保留各自版权和许可。

| 组件 | 用途 | 上游与许可信息 |
| --- | --- | --- |
| Python / Tcl / Tk | 桌面运行环境 | [Python](https://docs.python.org/3/license.html)、[Tcl/Tk](https://www.tcl.tk/software/tcltk/license.html) |
| PyInstaller | EXE 打包 | [许可及运行时例外](https://pyinstaller.org/en/stable/license.html) |
| Frida 17.10.1 | 客户端接口调用 | [源码与许可](https://github.com/frida/frida) |
| TkinterDnD2 0.4.3 / TkDND | 文件拖放 | [源码与许可](https://github.com/Eliav2/tkinterdnd2) |
| Mutagen 1.47.0 | 标签与音频信息 | [源码与许可](https://github.com/quodlibet/mutagen/tree/release-1.47.0) |
| FLAC | FLAC 解码 | [源码与许可](https://github.com/xiph/flac) |
| Vorbis Tools | OGG 解码 | [源码与许可](https://github.com/xiph/vorbis-tools) |
| LAME | MP3 编码 | [项目与源码](https://lame.sourceforge.io/) |
| FFmpeg（完整包） | MP4/M4A 解码及音频校验 | [FFmpeg 源码](https://github.com/FFmpeg/FFmpeg/tree/n7.1)、[Windows 构建提供方](https://www.gyan.dev/ffmpeg/builds/) |

## 完整包中的 FFmpeg

使用未修改的 `7.1-essentials_build-www.gyan.dev` Windows 构建。其版本输出声明启用 `--enable-gpl --enable-version3`，以 GNU GPL v3 或更新版本提供；完整包附带 GPL v3 许可文本及 `ffmpeg -version` 的构建信息。

FFmpeg 及该构建所包含的第三方库不适用本项目的 MIT License。构建说明、依赖及源码获取信息见 [Gyan 构建页面](https://www.gyan.dev/ffmpeg/builds/)；FFmpeg 7.1 源码见 [n7.1 标签](https://github.com/FFmpeg/FFmpeg/tree/n7.1)。
