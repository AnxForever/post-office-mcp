# -*- coding: utf-8 -*-
"""路径容错工具。

AgentOS / AionUi 环境里，用户上传的文件通常挂载在 /workspace，
输出写到 /workspace/output。模型有时会臆造 /inputs、/outputs 这类路径，
导致 PermissionError 或 FileNotFoundError。本模块负责：

  1. 输出目录不可写时，自动回退到可写目录，而不是直接抛 PermissionError；
  2. 输入文件不存在时，打印同级目录的文件清单和相近文件名，便于立刻纠正；
  3. 目标文件被 Excel 占用时，改存副本而不是直接失败。
"""
from __future__ import annotations

import difflib
import os
from pathlib import Path

# 输出目录回退顺序：优先当前工作目录下，其次 AgentOS 约定的 /workspace/output
_OUTPUT_CANDIDATES = ("output", "./output", "/workspace/output")

# 找不到输入文件时，在这些根目录下按文件名搜索
_SEARCH_ROOTS = ("/workspace", ".", "/mnt/data", "/data", "/inputs")


def _writable(path: Path) -> bool:
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / (".write_probe_%d" % os.getpid())
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return True
    except Exception:
        return False


def resolve_out(raw, label="输出目录") -> Path:
    """返回一个确认可写的输出目录。"""
    given = Path(raw).expanduser()
    if _writable(given):
        return given
    print("[WARN] %s 不可写: %s" % (label, given))
    for cand in _OUTPUT_CANDIDATES:
        path = Path(cand).expanduser()
        if _writable(path):
            print("[WARN] 已改用可写目录: %s" % path.resolve())
            return path
    raise SystemExit("[STOP] 找不到可写的输出目录，请显式传入一个可写路径")


def require_dir(raw, label="目录") -> Path:
    """检查目录存在，不存在时打印候选目录。"""
    given = Path(raw).expanduser()
    if given.is_dir():
        return given
    print("[STOP] %s 不存在: %s" % (label, given))
    parent = given.parent
    if parent.is_dir():
        dirs = sorted(x.name for x in parent.iterdir() if x.is_dir())
        print("  目录 %s 下的子目录 (%d 个):" % (parent, len(dirs)))
        for name in dirs[:30]:
            print("    -", name)
    else:
        print("  上级目录也不存在:", parent)
    raise SystemExit(2)


def require_file(raw, label="输入文件") -> Path:
    """检查文件存在，不存在时打印目录清单和相近文件名后退出。"""
    given = Path(raw).expanduser()
    if given.is_file():
        return given

    print("[STOP] %s 不存在: %s" % (label, given))
    name = given.name

    parent = given.parent
    if parent.is_dir():
        files = sorted(x.name for x in parent.iterdir() if x.is_file())
        print("  目录 %s 现有文件 (%d 个):" % (parent, len(files)))
        for filename in files[:30]:
            print("    -", filename)
        close = difflib.get_close_matches(name, files, n=5, cutoff=0.3)
        if close:
            print("  可能是其中之一:", ", ".join(close))
    else:
        print("  目录也不存在:", parent)
        cwd = Path.cwd()
        files = sorted(x.name for x in cwd.iterdir() if x.is_file())
        if files:
            print("  当前目录 %s 现有文件 (%d 个):" % (cwd, len(files)))
            for filename in files[:30]:
                print("    -", filename)
            close = difflib.get_close_matches(name, files, n=5, cutoff=0.3)
            if close:
                print("  可能是其中之一:", ", ".join(close))

    hits = []
    for root in _SEARCH_ROOTS:
        root_path = Path(root)
        if not root_path.is_dir():
            continue
        try:
            for found in root_path.rglob(name):
                if found.is_file():
                    hits.append(found)
                    if len(hits) >= 5:
                        break
        except Exception:
            continue
        if len(hits) >= 5:
            break
    if hits:
        print("  在其他位置找到同名文件:")
        for found in hits:
            print("    -", found.resolve())

    raise SystemExit(2)


def save_workbook(workbook, target, label=""):
    """保存工作簿；被占用时改存副本。返回实际保存路径。"""
    target = Path(target)
    try:
        workbook.save(target)
        workbook.close()
        print("[OK] 已保存: %s" % target.resolve())
        return target
    except PermissionError:
        alt = target.with_name("%s_副本%s" % (target.stem, target.suffix))
        print("[WARN] %s 被占用或无写权限（可能正在 Excel 中打开），改存为 %s" % (target, alt))
        workbook.save(alt)
        workbook.close()
        print("[OK] 已保存: %s" % alt.resolve())
        return alt
