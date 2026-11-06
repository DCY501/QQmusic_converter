# QQ 音乐解密 / 转码工具

Windows 桌面工具，保留简洁的 Tkinter 界面，支持批量处理本地音乐文件。个人学习娱乐使用。

## 使用

1. 添加文件、添加文件夹，或将文件/文件夹拖入窗口。
2. 选择输出目录，以及“保持原格式”或“统一转 MP3”。
3. 点击“开始处理”。需要解密时，请先运行兼容的 QQ 音乐客户端。
4. 点击条目的“打开输出”“查看原因”或“查看提示”；失败后可点“重试失败”。
5. 要换格式或音质，可多选文件，右键“重新处理选中项”；右键“文件详情”可查看完整源/输出路径、时长、采样率和标签。

同名输出自动编号，例如 `歌曲 (2).mp3`。源文件已在目标位置时跳过。转换先在独立临时目录完成，成功后再提交结果，不覆盖原文件和已有输出。

新版在同一 Windows 登录会话中只运行一个窗口，重复启动会通知并唤起已有窗口，不再加载和改写另一份任务列表。程序异常退出后锁自动释放。旧版不具备此机制，使用时请统一启动新版。

“停止处理”表示**当前文件完成后停止**。关闭窗口也会等当前文件处理和资源清理完成。未处理的条目可继续；“重试失败”只处理失败条目。

处理时，上方进度条显示整个批次，下方显示当前阶段；空闲时隐藏进度条，无待处理文件时禁用开始按钮。阶段包括连接、解密、解码、编码、保留标签和保存结果；解码/编码支持读取工具实际进度，无法取得百分比时显示活动指示和当前文件耗时。

所有新输出在正式保存前默认进行基础校验：确认真实音频格式、参数有效，并比较源音频和输出时长（允许少量编码延迟差异）。勾选“完整解码校验”后，还会用 FFmpeg 从头到尾解码输出，发现错误即判失败，临时输出不提交。基础校验不能保证每个音频帧都无损坏；完整校验需要 FFmpeg 并增加处理时间。结果可在“文件详情”中查看，随任务记录保存；旧任务未校验时明确显示暂无记录。

“工具 → 环境与兼容性诊断”可检查 QQ 音乐进程架构、文件版本、必要导出函数和本地工具，并复制诊断信息。解密前会执行同样的接口检查；当前仅支持具有所需接口的 32 位 QQ 音乐。接口检查通过仍不保证每个文件都能解密，文件权限、文件状态及客户端实现仍可能影响结果。

自动保存最近一次窗口的任务列表与设置，重新打开后恢复但不会自动开始转换。恢复时校验源文件和输出的大小、修改时间；输出丢失/变化或源文件变化会回到等待状态，源文件无法访问会显示失败原因。异常退出时正在处理的文件会从头重做，已完整保存并记录的文件不会重复处理。清空列表也会清空保存的任务记录。

## 格式支持

| 输入 | 保持原格式 | 转 MP3 |
|---|---|---|
| MFLAC / MFLAC0 | 解密为 FLAC | 支持 |
| MGG / MGG0 / MGG1 / MGGL | 解密为 OGG | 支持 |
| MMP4 | 解密为 MP4 | 有 FFmpeg 时支持，否则提示并跳过 |
| FLAC / OGG / WAV | 原样复制 | 支持 |
| MP3 | 原样复制 | 原样复制，不重复编码 |
| M4A / MP4 | 原样复制 | 有 FFmpeg 时支持，否则提示并跳过 |

转码迁移常用标签：歌名、歌手、专辑、专辑歌手、作曲、流派、日期、曲目/碟号、歌词、备注，以及 FLAC/OGG、MP4/M4A 的内嵌封面和 WAV 的 ID3 封面。并非所有自定义标签都能迁移。标签处理失败时保留正常音频，在列表中单独提示。

MP4/M4A 只转换第一条音轨，视频不输出，多声道音频混合为立体声 MP3；选择“保持原格式”则保留解密/复制后的原文件。

普通音频不需要 QQ 音乐或 Frida 连接成功。混合批次中，连接失败只影响加密文件。

## 运行与打包

需要 Windows、Python 3.8+（本机验证为 Python 3.12.1）。在项目根目录使用 PowerShell：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe src\qqmusic_converter_gui.py
```

新版程序：`dist\v5\QQMusicConverter.exe`。

重新打包：运行 `./build.ps1`，脚本先执行测试，再生成 EXE。

固定依赖见 `requirements.txt`，构建依赖见 `requirements-dev.txt`，资源清单见 `qqmusic_converter.spec`。EXE 包含解码器、编码器、Frida、拖拽组件和标签处理组件。FFmpeg 不内嵌在 EXE 中；若项目根目录有 `ffmpeg.exe`，打包脚本会将它复制到输出目录。

可选的 `ffmpeg.exe` 放在 **QQMusicConverter.exe 同级目录**，用于 MP4/M4A 转码及 FLAC 备用解码；源码运行也会检查项目根目录。移动程序时一起携带 FFmpeg 即可保留这些能力。

## 配置与日志

- 配置：`%LOCALAPPDATA%\QQMusicConverter\config.json`。
- “工具 → 打开配置目录”可打开本次运行实际使用的位置。Windows 应用虚拟化可能将此目录重定向到另一磁盘；保存时会解析实际目录，确保临时文件与目标位于同一位置。
- 失败日志：`%LOCALAPPDATA%\QQMusicConverter\error.log`。
- 最近一次任务：`%LOCALAPPDATA%\QQMusicConverter\session.json`（原子替换保存）。
- 首次启动尝试读取应用目录中的旧 `config.json`；之后保存到用户配置目录。
- 旧配置里的相对输出路径以应用目录为基准；保存时转换为绝对路径。
- 配置损坏时使用默认设置；保存采用临时文件替换，避免破坏旧配置。
- 扫描排除所选目录内部的输出目录和本工具工作目录，避免重复导入产物。

## 开发验证

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

测试覆盖防覆盖与并发提交、混合批次、停止续跑、失败清理、标签降级、真实 WAV/FLAC/OGG/M4A 转换、中文路径、配置校验、目录扫描、任务恢复、文件详情、阶段进度和 Tk 界面流程。Hook 模拟测试需要 Node.js，OGG/M4A 集成测试需要可选 FFmpeg；缺少时对应测试跳过。

打包自检可调用 `QQMusicConverter.exe --self-check 报告路径.json`，等待进程退出后检查 JSON 的 `ok` 字段。自检验证资源、拖拽加载、真实转码、中文标签、任务记录恢复和文件详情；检测到 FFmpeg 时额外验证 M4A 转码与阶段进度。自检不连接 QQ 音乐或写入用户配置。

自检还覆盖单实例通知、基础校验和（有 FFmpeg 时的）完整解码校验。测试套件包含损坏音频帧、错误时长/格式、校验失败不提交输出、跨进程锁释放与不兼容客户端的清理行为。

自检在用户数据目录下创建独立临时目录，验证配置和任务记录的首次保存、覆盖与读取，结束后清理，不改动现有配置。GUI 回归测试还覆盖非法/空输出路径、无待处理任务时的按钮状态，以及保存恢复后的错误提示清除。

2026-11-07 已通过源码批处理连接本机 QQ 音乐 22.05（32 位）实测：MGG、MFLAC、MMP4 解密成功；MGG、MFLAC 及普通 FLAC、OGG 转 MP3 成功，7 个输出均通过 FFmpeg 全文件解码检查，源文件 SHA-256 保持不变。首轮结果位于 `converted/live-test-20261107-b48ae4fb/`。

本轮新增 MMP4 → MP3 也已实测通过：`converted/live-test-v3-20261107-6ac9ff90/`，包含《九万字》的输出和报告；全文件解码无报错，原文件哈希未变。

v4 实测：客户端文件版本 22.5.0.0、架构 ia32，必要接口检查通过；项目下载目录的 5 首歌曲均成功转换并通过程序内的完整解码校验，源文件哈希未变。报告位于 `converted/live-test-v4-20261107-fb48844b/report.json`。

当前 Hook 使用客户端导出函数及固定对象布局，以上验证不代表所有客户端版本和文件均兼容。

## 项目结构

```text
src/
  models.py                  文件模型、结果、格式映射、配置与路径
  scanner.py                 后台扫描与文件检测
  decryptor.py               Frida 会话和解密临时文件
  hook_qqmusic_v17.js         客户端解密调用与读取完整性检查
  converter.py               flac / oggdec / lame / ffmpeg 工具链
  metadata.py                常用标签和封面迁移
  session_store.py           任务持久化与恢复校验
  file_details.py            本地文件与音频详情
  audio_validation.py        输出格式、时长及完整解码校验
  client_diagnostics.py      客户端接口检查与环境报告
  single_instance.py         单实例运行与窗口唤起通知
  worker.py                  批处理、失败隔离、输出提交、错误日志
  gui.py                     界面与主线程事件队列
  diagnostics.py             打包自检
  qqmusic_converter_gui.py   入口
tests/                       回归与集成测试
build.ps1                    测试后打包
qqmusic_converter.spec       PyInstaller 资源清单
```

## 参考与依赖

- 解密思路参考 [decrypt-mflac-frida](https://github.com/Noisrev/decrypt-mflac-frida)。
- 音频工具：[FLAC](https://xiph.org/flac/)、[Vorbis Tools](https://xiph.org/downloads/)、[LAME](https://lame.sourceforge.io/)、可选 [FFmpeg](https://ffmpeg.org/)。
- 拖拽：[TkinterDnD2](https://github.com/Eliav2/tkinterdnd2)。
- 标签：[Mutagen](https://mutagen.readthedocs.io/en/latest/user/id3.html)。
