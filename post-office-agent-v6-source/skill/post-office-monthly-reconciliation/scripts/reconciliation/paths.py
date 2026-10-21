# -*- coding: utf-8 -*-
"""上报目录解析。

AgentOS 不同版本把用户上传的附件挂在不同位置：旧版是 `/workspace`，
新版是只读的 `/inputs`。业务脚本不能假设挂载点，统一用这里解析。

约定：
- 工作区优先。如果 workspace 里已经有 Excel，就继续用它（本地运行、旧版平台）。
- 工作区没有 Excel 时，依次尝试 `/inputs`、`/workspace/uploads`、`/mnt/data`。
- 都没有时返回 workspace 本身，让调用方按原有逻辑报「缺少输入」。
"""
from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

SKIP_DIRS = {"skills", "output", "__pycache__", ".git", "node_modules", "proc", "sys", "dev"}
PLATFORM_INPUT_DIRS = (Path("/inputs"), Path("/workspace/uploads"), Path("/mnt/data"))
# AgentOS 的成果目录：写入这里的文件由平台自动登记为可下载的「成果文件」。
PLATFORM_OUTPUT_DIR = Path("/outputs")


def has_excel(root: Path | str) -> bool:
    root = Path(root)
    if not root.is_dir():
        return False
    for dirpath, dirnames, filenames in os.walk(root):
        relative = Path(dirpath).relative_to(root)
        if relative.parts and relative.parts[0].lower() in {"skills", "output"}:
            dirnames[:] = []
            continue
        dirnames[:] = [
            name for name in dirnames if name not in SKIP_DIRS and not name.startswith(".")
        ]
        if any(
            name.lower().endswith((".xlsx", ".xls")) and not name.startswith("~$")
            for name in filenames
        ):
            return True
    return False


def resolve_input_dir(workspace: Path | str) -> Path:
    workspace = Path(workspace).resolve()
    for candidate in (workspace, *PLATFORM_INPUT_DIRS):
        try:
            if has_excel(candidate):
                return candidate.resolve()
        except OSError:
            continue
    return workspace


def resolve_publish_dir() -> Path | None:
    """返回平台成果目录；没有挂载或不可写时返回 None。

    只在沙箱容器里生效：Windows 上 Path('/outputs') 会落到当前盘符根目录，
    不能当平台目录用，否则本地跑一次就会往 C:\\outputs 写文件。
    """
    if sys.platform == "win32":
        return None
    try:
        if PLATFORM_OUTPUT_DIR.is_dir() and os.access(PLATFORM_OUTPUT_DIR, os.W_OK):
            return PLATFORM_OUTPUT_DIR
    except OSError:
        pass
    return None


def publish_files(
    paths,
    target_dir: Path | str | None = None,
) -> list[dict[str, object]]:
    """把交付文件复制到平台成果目录，供 AgentOS 登记下载。

    平台只收集 /outputs 下的文件；业务脚本仍把中间状态和 run-state
    留在 workspace/output。本地运行（没有 /outputs）时返回空列表。
    """
    target = Path(target_dir) if target_dir is not None else resolve_publish_dir()
    if target is None:
        return []
    target.mkdir(parents=True, exist_ok=True)
    published: list[dict[str, object]] = []
    for raw in paths:
        source = Path(raw)
        if not source.is_file():
            continue
        destination = target / source.name
        if destination.resolve() != source.resolve():
            shutil.copy2(source, destination)
        published.append({
            "name": destination.name,
            "path": str(destination),
            "size": destination.stat().st_size,
        })
    return published
