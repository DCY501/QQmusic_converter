# 源码说明

运行、格式支持和打包步骤统一维护在上级 README.md。

数据流：

```text
gui → scanner → 队列 → gui（添加文件）
gui → worker → decryptor / converter / metadata
             → 独立临时目录 → 不覆盖地提交最终文件
             → 队列 → gui（进度和结果）
```

- GUI 只在主线程读写 Tk；工作线程不调用 Tk，也不读取 Tk 变量。
- 转换参数在启动时固定；转换期间锁定文件列表与输出设置。
- 工具失败抛出带诊断的异常，worker 将其转换为 ProcessingResult。
- 单文件错误不终止批次，错误日志由 worker 记录。
- 停止通过 threading.Event 在文件边界检查；退出等待工作线程释放资源。
- 文件扩展名与解密输出映射集中在 models.py。
- 资源目录和用户配置目录分开；打包资源位于 _MEIPASS/src。
- 标签写入独立副本，失败时保留正常编码的音频并提示。
