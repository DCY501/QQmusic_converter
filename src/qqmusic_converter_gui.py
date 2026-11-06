#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QQ音乐解密/转码工具 - 入口文件
兼容旧运行方式: python src/qqmusic_converter_gui.py
"""

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='QQ 音乐解密/转码工具')
    parser.add_argument('--self-check', metavar='REPORT', help='运行打包自检并写入 JSON 报告')
    args = parser.parse_args()
    if args.self_check:
        from diagnostics import run_self_check
        raise SystemExit(run_self_check(args.self_check))
    from gui import main
    main()
