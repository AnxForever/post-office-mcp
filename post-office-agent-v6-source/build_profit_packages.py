# -*- coding: utf-8 -*-
"""打包「盈亏表」技能与专家包。

    python build_profit_packages.py

产物（项目根目录）：
    post-office-profit-report-skill.zip      技能包，POST /api/skills/upload
    邮局一干二干盈亏表V1.0-agent.zip          专家包，控制台「发布专家」
"""
from __future__ import annotations

import json
import zipfile
from pathlib import Path, PurePosixPath

ROOT = Path(__file__).resolve().parent
SKILL_NAME = "post-office-profit-report"
SKILL_ROOT = ROOT / "skill" / SKILL_NAME
AGENT_ROOT = ROOT / "agent-profit"
DIST = ROOT.parent
FIXED_ZIP_TIME = (2026, 10, 28, 0, 0, 0)

MANIFEST = json.loads((SKILL_ROOT / "manifest.json").read_text(encoding="utf-8"))
VERSION = MANIFEST["version"]
SKILL_ZIP = DIST / f"{SKILL_NAME}-skill.zip"
AGENT_ZIP = DIST / f"邮局一干二干盈亏表V{'.'.join(VERSION.split('.')[:2])}-agent.zip"


def skill_files() -> list[Path]:
    return sorted(
        path for path in SKILL_ROOT.rglob("*")
        if path.is_file() and "__pycache__" not in path.parts and path.suffix != ".pyc"
    )


def add(archive: zipfile.ZipFile, name: str, data: bytes) -> None:
    info = zipfile.ZipInfo(str(PurePosixPath(name)), FIXED_ZIP_TIME)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o644 << 16
    archive.writestr(info, data)


def main() -> None:
    files = skill_files()

    with zipfile.ZipFile(SKILL_ZIP, "w") as archive:
        for path in files:
            add(archive, path.relative_to(SKILL_ROOT).as_posix(), path.read_bytes())

    with zipfile.ZipFile(AGENT_ZIP, "w") as archive:
        for name in ("agent.yaml", "prompt.md"):
            add(archive, name, (AGENT_ROOT / name).read_bytes())
        for path in files:
            add(archive, f"skills/{SKILL_NAME}/{path.relative_to(SKILL_ROOT).as_posix()}",
                path.read_bytes())

    for path in (SKILL_ZIP, AGENT_ZIP):
        print(f"[OK] {path.name}  {path.stat().st_size:,} B  ({len(files)} 个技能文件)")


if __name__ == "__main__":
    main()
