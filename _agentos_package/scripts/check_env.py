#!/usr/bin/env python3
"""
邮局报表助手 · 环境预检脚本
检查 Python 版本和必要依赖是否可用。

用法: python3 check_env.py
退出码 0 = 通过, 非 0 = 失败
"""
import sys
import subprocess


def main():
    print(f"Python: {sys.version.split()[0]} ({sys.executable})")

    if sys.version_info < (3, 10):
        print(f"❌ Python 版本过低，需要 3.10+，当前 {sys.version_info.major}.{sys.version_info.minor}")
        sys.exit(1)

    try:
        import openpyxl
        print(f"openpyxl: {openpyxl.__version__} ✅")
    except ImportError:
        print("❌ openpyxl 未安装")
        print("   安装命令: pip install openpyxl")
        print("   或:       pip3 install openpyxl")
        sys.exit(2)

    # 检查脚本是否都在
    from pathlib import Path
    script_dir = Path(__file__).parent
    for name in ['build_draft.py', 'build_split.py', 'build_pnl.py']:
        p = script_dir / name
        if p.exists():
            print(f"{name}: 存在 ✅")
        else:
            print(f"❌ 缺失脚本: {p}")
            sys.exit(3)

    print("\n✅ 环境就绪。可以执行报表生成。")
    sys.exit(0)


if __name__ == '__main__':
    main()
