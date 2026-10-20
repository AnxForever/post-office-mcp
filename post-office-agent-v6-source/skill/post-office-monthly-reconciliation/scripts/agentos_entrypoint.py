from __future__ import annotations

import argparse
import json
import runpy
import sys
from pathlib import Path


def _resolve_skill_root(explicit: Path | str | None) -> Path:
    if explicit:
        return Path(explicit).resolve()
    # 尝试一：从脚本自身位置推断（适用于从 skill 包内直接运行）
    candidate = Path(__file__).resolve().parent.parent
    if (candidate / "manifest.json").exists():
        return candidate
    # 尝试二：搜索 /workspace/skills/ 下包含本脚本的 skill 包
    # （适用于 AgentOS 把入口脚本映射到 /workspace/scripts/ 的情况）
    for base in ("/workspace/skills", "/workspace"):
        base_path = Path(base)
        if not base_path.is_dir():
            continue
        for child in base_path.iterdir():
            target = child / "scripts" / "agentos_entrypoint.py"
            if target.is_file():
                return child.resolve()
    raise FileNotFoundError(
        "无法定位 skill 包根目录。请通过 --bundle-root 显式指定。"
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="AgentOS Skill 统一入口")
    parser.add_argument(
        "--phase", required=True,
        choices=["summary", "review", "final", "profit", "inspect", "runtime", "validate"],
    )
    parser.add_argument("--workspace", required=True, type=Path)
    parser.add_argument("--month")
    parser.add_argument("--caotan", type=Path)
    parser.add_argument("--gangwu", type=Path)
    parser.add_argument(
        "--bundle-root",
        type=Path,
        help="Skill 包根目录（默认从脚本位置自动推断）",
    )
    parser.add_argument("--review-file", type=Path)
    parser.add_argument("--base-file", type=Path, help="profit 阶段：基础报表原版")
    parser.add_argument("--finance-file", type=Path, help="profit 阶段：财务收入表")
    parser.add_argument("--profit-tpl-dir", type=Path, help="profit 阶段（可选）：盈亏模板目录，仅作路线占位")
    parser.add_argument("--no-template-routes", action="store_true", help="profit 阶段：不追加模板占位路线")
    parser.add_argument("--previous-first-profit", type=Path)
    parser.add_argument("--previous-second-profit", type=Path)
    parser.add_argument("--output-dir", type=Path)

    args = parser.parse_args(argv)
    skill_root = _resolve_skill_root(args.bundle_root)
    scripts_dir = skill_root / "scripts"
    workspace = Path(args.workspace).resolve()

    # AgentOS 沙箱中 skill 包放在 /workspace/skills/<name>/，
    # 需要在 sys.path 中注册才能让脚本内的 from reconciliation.xxx import 生效
    sys.path.insert(0, str(skill_root))
    sys.path.insert(0, str(scripts_dir))

    if args.phase == "runtime":
        child_argv = [
            "check_runtime.py",
            "--workspace", str(workspace),
        ]
        if args.month:
            child_argv += ["--month", args.month]
        sys.argv[:] = child_argv
        runpy.run_path(str(scripts_dir / "check_runtime.py"), run_name="__main__")
        return 0

    if args.phase == "inspect":
        child_argv = [
            "inspect_inputs.py",
            "--workspace", str(workspace),
        ]
        if args.month:
            child_argv += ["--month", args.month]
        sys.argv[:] = child_argv
        runpy.run_path(str(scripts_dir / "inspect_inputs.py"), run_name="__main__")
        return 0

    if args.phase == "validate":
        child_argv = [
            "validate_outputs.py",
            "--workspace", str(workspace),
        ]
        if args.month:
            child_argv += ["--month", args.month]
        sys.argv[:] = child_argv
        runpy.run_path(str(scripts_dir / "validate_outputs.py"), run_name="__main__")
        return 0

    if args.phase == "summary":
        if not args.caotan or not args.gangwu:
            print(json.dumps({
                "ok": False,
                "phase": "failed",
                "errors": ["--caotan 和 --gangwu 必须同时指定"],
                "next_action": "请上传草滩和港务在线测量表",
            }, ensure_ascii=False))
            return 2
        output_dir = args.output_dir or (workspace / "output")
        child_argv = [
            "build_model_summary.py",
            "--caotan", str(args.caotan),
            "--gangwu", str(args.gangwu),
            "--output-dir", str(output_dir),
            "--month", args.month or "2026-04",
        ]
        sys.argv[:] = child_argv
        runpy.run_path(str(scripts_dir / "build_model_summary.py"), run_name="__main__")
        return 0

    if args.phase == "review":
        child_argv = [
            "run_workflow.py",
            "--phase", "review",
            "--workspace", str(workspace),
            "--bundle-root", str(skill_root),
        ]
        if args.month:
            child_argv += ["--month", args.month]
        sys.argv[:] = child_argv
        runpy.run_path(str(scripts_dir / "run_workflow.py"), run_name="__main__")
        return 0

    if args.phase == "final":
        if not args.review_file:
            print(json.dumps({
                "ok": False,
                "phase": "failed",
                "errors": ["--review-file 必须指定（核对后的初版路径）"],
                "next_action": "请上传人工核对修正后的基础报表初版",
            }, ensure_ascii=False))
            return 2
        child_argv = [
            "run_workflow.py",
            "--phase", "final",
            "--workspace", str(workspace),
            "--review-file", str(args.review_file),
            "--bundle-root", str(skill_root),
        ]
        if args.previous_first_profit:
            child_argv += ["--previous-first-profit", str(args.previous_first_profit)]
        if args.previous_second_profit:
            child_argv += ["--previous-second-profit", str(args.previous_second_profit)]
        sys.argv[:] = child_argv
        runpy.run_path(str(scripts_dir / "run_workflow.py"), run_name="__main__")
        return 0

    if args.phase == "profit":
        # 仅生成盈亏表：基础报表原版 + 财务收入表 是第一基准；盈亏模板可选（只作版式占位）
        missing = [
            flag for flag, value in (
                ("--base-file", args.base_file),
                ("--finance-file", args.finance_file),
                ("--month", args.month),
            ) if not value
        ]
        if missing:
            print(json.dumps({
                "ok": False,
                "phase": "failed",
                "errors": [f"缺少参数: {', '.join(missing)}"],
                "next_action": "请提供基础报表原版、财务收入表和目标月份",
            }, ensure_ascii=False))
            return 2
        output_dir = args.output_dir or (workspace / "output")
        child_argv = [
            "build_profit_tables.py",
            "--base", str(args.base_file),
            "--finance", str(args.finance_file),
            "--month", args.month,
            "--out", str(output_dir),
        ]
        if args.profit_tpl_dir:
            child_argv += ["--tpl-dir", str(args.profit_tpl_dir)]
        if args.no_template_routes:
            child_argv += ["--no-template-routes"]
        sys.argv[:] = child_argv
        runpy.run_path(str(scripts_dir / "build_profit_tables.py"), run_name="__main__")
        return 0

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
