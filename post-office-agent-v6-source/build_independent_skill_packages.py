# -*- coding: utf-8 -*-
"""把四个独立技能分别打成 AgentOS skill 包。

    python build_independent_skill_packages.py

产物落在仓库上一级目录（和原来的 *-skill.zip 放一起）：
    post-office-model-summary-skill.zip
    post-office-draft-report-skill.zip
    post-office-split-report-skill.zip
    post-office-profit-report-skill.zip

四个包互相不依赖 run-state，按顺序单独上传即可：
    01 车型汇总 -> 02 初版 -> 人工核对 -> 03 拆分 -> 04/05 盈亏
"""
from __future__ import annotations

import json
import zipfile
from pathlib import Path, PurePosixPath

ROOT = Path(__file__).resolve().parent
SKILL_ROOT = ROOT / "skill"
DIST = ROOT.parent
FIXED_ZIP_TIME = (2026, 7, 28, 0, 0, 0)

SKILLS = (
    "post-office-model-summary",
    "post-office-draft-report",
    "post-office-split-report",
    "post-office-profit-report",
)
REQUIRED = {"SKILL.md", "manifest.json"}


def source_files(root: Path) -> list[Path]:
    return sorted(
        path
        for path in root.rglob("*")
        if path.is_file() and "__pycache__" not in path.parts and path.suffix != ".pyc"
    )


def add_bytes(archive: zipfile.ZipFile, name: str, content: bytes) -> None:
    info = zipfile.ZipInfo(str(PurePosixPath(name)), FIXED_ZIP_TIME)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o644 << 16
    archive.writestr(info, content)


def main() -> None:
    results = []
    for name in SKILLS:
        root = SKILL_ROOT / name
        if not root.is_dir():
            raise FileNotFoundError(f"技能目录不存在: {root}")
        files = source_files(root)
        present = {path.relative_to(root).as_posix() for path in files}
        missing = sorted(REQUIRED - present)
        if missing:
            raise ValueError(f"{name} 缺少必需文件: {', '.join(missing)}")
        manifest = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
        output = DIST / f"{name}-skill.zip"
        with zipfile.ZipFile(output, "w") as archive:
            for path in files:
                add_bytes(archive, path.relative_to(root).as_posix(), path.read_bytes())
        results.append({
            "skill": name,
            "version": manifest.get("version"),
            "zip": str(output),
            "files": len(files),
        })
    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
