from __future__ import annotations

import runpy
import sys
from pathlib import Path
from typing import Mapping, Sequence


ALLOWED_SCRIPTS = {
    "build_model_summary": "build_model_summary.py",
    "build_profit_tables": "build_profit_tables.py",
    "inspect_inputs": "inspect_inputs.py",
    "build_review_workbook": "build_review_workbook.py",
    "build_final_reports": "build_final_reports.py",
    "validate_outputs": "validate_outputs.py",
    "check_runtime": "check_runtime.py",
}


def run_script(
    script_name: str,
    arguments: Sequence[str],
    *,
    scripts_dir: Path | str | None = None,
    allowed_scripts: Mapping[str, str] | None = None,
) -> int:
    """Run one approved Skill entry with normal Python script semantics."""
    allowed = dict(allowed_scripts or ALLOWED_SCRIPTS)
    if script_name not in allowed:
        raise ValueError(f"未批准的 Skill 脚本: {script_name}")

    root = Path(scripts_dir or Path(__file__).resolve().parent).resolve()
    target = (root / allowed[script_name]).resolve()
    try:
        target.relative_to(root)
    except ValueError as exc:
        raise ValueError(f"脚本路径越界: {script_name}") from exc
    if not target.is_file():
        raise FileNotFoundError(f"Skill 脚本不存在: {target}")

    original_argv = list(sys.argv)
    original_path = list(sys.path)
    sys.argv[:] = [str(target), *arguments]
    sys.path.insert(0, str(root))
    try:
        runpy.run_path(str(target), run_name="__main__")
    finally:
        sys.argv[:] = original_argv
        sys.path[:] = original_path
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    values = list(sys.argv[1:] if argv is None else argv)
    if not values:
        approved = ", ".join(sorted(ALLOWED_SCRIPTS))
        raise ValueError(f"缺少脚本名；允许值: {approved}")
    return run_script(values[0], values[1:])


if __name__ == "__main__":
    raise SystemExit(main())
