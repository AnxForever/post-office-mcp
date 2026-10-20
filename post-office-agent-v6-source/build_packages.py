from __future__ import annotations

import base64
import json
import time
import zipfile
from pathlib import Path, PurePosixPath


ROOT = Path(__file__).resolve().parent
SKILL_NAME = "post-office-monthly-reconciliation"
SKILL_ROOT = ROOT / "skill" / SKILL_NAME
AGENT_ROOT = ROOT / "agent"
DIST = ROOT.parent
SKILL_ZIP = DIST / "post-office-monthly-reconciliation-skill.zip"
AGENT_ZIP = DIST / "邮局委办运输报表核对V6.0-agent.zip"
FIXED_ZIP_TIME = (2026, 7, 28, 0, 0, 0)


def source_files() -> list[Path]:
    return sorted(
        path
        for path in SKILL_ROOT.rglob("*")
        if path.is_file()
        and "__pycache__" not in path.parts
        and path.suffix != ".pyc"
    )


def add_bytes(archive: zipfile.ZipFile, name: str, content: bytes) -> None:
    info = zipfile.ZipInfo(str(PurePosixPath(name)), FIXED_ZIP_TIME)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o644 << 16
    archive.writestr(info, content)


def build_skill_zip(files: list[Path]) -> None:
    with zipfile.ZipFile(SKILL_ZIP, "w") as archive:
        for path in files:
            add_bytes(archive, path.relative_to(SKILL_ROOT).as_posix(), path.read_bytes())


def build_agent_json(files: list[Path]) -> dict[str, object]:
    prompt = (AGENT_ROOT / "prompt.md").read_text(encoding="utf-8")
    embedded = [
        {
            "path": f"{SKILL_NAME}/{path.relative_to(SKILL_ROOT).as_posix()}",
            "content": base64.b64encode(path.read_bytes()).decode("ascii"),
        }
        for path in files
    ]
    return {
        "format": "agentos-v1",
        "exportedAt": int(time.time() * 1000),
        "agent": {
            "name": "邮局月度委办运输报表核对 V6.0",
            "description": "通过受控 Skill 执行月度预检、人工核对和报表交付：车型汇总表 → 基础报表初版 → 人工闸门 → 拆分版 → 盈亏表",
            "icon": "[PO]",
            "version": "6.1.0",
            "category": "财务分析",
            "type": "general",
            "skills": [SKILL_NAME],
            "toolWhitelist": ["code.execute"],
            "systemPrompt": prompt,
        },
        "skillFiles": {SKILL_NAME: embedded},
    }


def build_agent_zip(payload: dict[str, object]) -> None:
    agent_json = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    prompt = (AGENT_ROOT / "prompt.md").read_bytes()
    with zipfile.ZipFile(AGENT_ZIP, "w") as archive:
        add_bytes(archive, "agent.json", agent_json)
        add_bytes(archive, "prompt.md", prompt)


def main() -> None:
    files = source_files()
    required = {"SKILL.md", "manifest.json", "assets/split-report-template.xlsx"}
    present = {path.relative_to(SKILL_ROOT).as_posix() for path in files}
    missing = sorted(required - present)
    if missing:
        raise ValueError(f"Skill 缺少必需文件: {', '.join(missing)}")
    build_skill_zip(files)
    build_agent_zip(build_agent_json(files))
    print(json.dumps({
        "skill_zip": str(SKILL_ZIP),
        "agent_zip": str(AGENT_ZIP),
        "embedded_files": len(files),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
