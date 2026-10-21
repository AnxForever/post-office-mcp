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
from pathlib import Path

SKIP_DIRS = {"skills", "output", "__pycache__", ".git", "node_modules", "proc", "sys", "dev"}
PLATFORM_INPUT_DIRS = (Path("/inputs"), Path("/workspace/uploads"), Path("/mnt/data"))


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
