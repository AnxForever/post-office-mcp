from __future__ import annotations

import argparse
import json
import runpy
import sys
from datetime import datetime
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
        choices=["list", "detect", "summary", "review", "final", "profit", "inspect", "runtime", "validate"],
    )
    parser.add_argument("--workspace", required=True, type=Path)
    parser.add_argument("--month")
    parser.add_argument(
        "--task", choices=["profit", "review", "summary"], default="review",
        help="detect 阶段：按哪条流程检查输入",
    )
    parser.add_argument(
        "--stage", choices=["review", "final"], default="review",
        help="validate 阶段：校验哪一阶段的输出",
    )
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

    if args.phase == "list":
        # 列出工作区文件，供调用方确认真实文件名。
        # 用途：避免模型照抄示例文件名或臆造 /inputs、/outputs 这类不存在的路径。
        files = []
        if workspace.is_dir():
            for path in sorted(workspace.iterdir()):
                if not path.is_file():
                    continue
                stat = path.stat()
                files.append({
                    "name": path.name,
                    "path": str(path),
                    "size": stat.st_size,
                    "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M"),
                })
        print(json.dumps({
            "phase": "listed",
            "workspace": str(workspace),
            "exists": workspace.is_dir(),
            "file_count": len(files),
            "files": files,
            "hint": "下一步执行 --phase detect 按表结构识别文件角色；脚本参数只使用 resolved 中的真实路径。",
        }, ensure_ascii=False, indent=2))
        return 0 if workspace.is_dir() else 2

    if args.phase == "detect":
        # 按表结构识别工作区里每个文件是什么，不依赖文件名。
        from reconciliation.detect import detect_for_phase

        if not args.month:
            print(json.dumps({
                "ok": False,
                "phase": "failed",
                "errors": ["detect 阶段需要 --month YYYY-MM"],
            }, ensure_ascii=False))
            return 2
        result = detect_for_phase(workspace, args.month, args.task)
        ok = not result["missing"] and not result["ambiguous"]
        payload = {"ok": ok, **result}
        if result["ambiguous"]:
            payload["next_action"] = "存在多个同角色文件，请用参数指明使用哪一个"
        elif result["missing"]:
            payload["next_action"] = "缺少必需输入：" + "、".join(result["missing"])
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0 if ok else 3

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
        # 只做输入预检，不生成报表。workspace 既是输入目录，也是 run-state 的落盘位置。
        output_dir = args.output_dir or (workspace / "output")
        output_dir.mkdir(parents=True, exist_ok=True)
        child_argv = [
            "inspect_inputs.py",
            "--input-dir", str(workspace),
            "--workspace-dir", str(output_dir),
        ]
        if args.month:
            child_argv += ["--month", args.month]
        sys.argv[:] = child_argv
        runpy.run_path(str(scripts_dir / "inspect_inputs.py"), run_name="__main__")
        return 0

    if args.phase == "validate":
        # 校验已有输出，读取 <workspace>/output/run-state.json
        output_dir = args.output_dir or (workspace / "output")
        child_argv = [
            "validate_outputs.py",
            "--workspace-dir", str(output_dir),
            "--phase", args.stage,
        ]
        sys.argv[:] = child_argv
        runpy.run_path(str(scripts_dir / "validate_outputs.py"), run_name="__main__")
        return 0

    if args.phase == "summary":
        if not args.month:
            print(json.dumps({
                "ok": False,
                "phase": "failed",
                "errors": ["summary 阶段需要 --month YYYY-MM"],
            }, ensure_ascii=False))
            return 2
        caotan, gangwu = args.caotan, args.gangwu
        if not caotan or not gangwu:
            # 未指定文件时，按表结构自动识别草滩/港务在线表
            from reconciliation.detect import detect_for_phase

            result = detect_for_phase(workspace, args.month, "summary")
            resolved = result["resolved"]
            if not caotan and len(resolved.get("online_caotan", [])) == 1:
                caotan = Path(resolved["online_caotan"][0])
            if not gangwu and len(resolved.get("online_gangwu", [])) == 1:
                gangwu = Path(resolved["online_gangwu"][0])
            if not caotan or not gangwu:
                print(json.dumps({
                    "ok": False,
                    "phase": "failed",
                    "errors": ["无法唯一确定草滩/港务在线测量表"],
                    "detected": result["files"],
                    "missing": result["missing"],
                    "ambiguous": result["ambiguous"],
                    "next_action": "请显式传入 --caotan 和 --gangwu",
                }, ensure_ascii=False, indent=2))
                return 3
        output_dir = args.output_dir or (workspace / "output")
        child_argv = [
            "build_model_summary.py",
            "--caotan", str(caotan),
            "--gangwu", str(gangwu),
            "--output-dir", str(output_dir),
            "--month", args.month,
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
        if not args.month:
            print(json.dumps({
                "ok": False,
                "phase": "failed",
                "errors": ["profit 阶段需要 --month YYYY-MM"],
            }, ensure_ascii=False))
            return 2
        base_file, finance_file = args.base_file, args.finance_file
        if not base_file or not finance_file:
            # 未指定文件时，按表结构自动识别：含「一干汇总/二干基础」的是基础报表，
            # 表头含 结算收入 + 普服收入 的是财务收入表（客户可能叫「运输局」）。
            from reconciliation.detect import detect_for_phase

            result = detect_for_phase(workspace, args.month, "profit")
            resolved = result["resolved"]
            if not base_file and len(resolved.get("base_report", [])) == 1:
                base_file = Path(resolved["base_report"][0])
            if not finance_file and len(resolved.get("finance", [])) == 1:
                finance_file = Path(resolved["finance"][0])
            if not base_file or not finance_file:
                print(json.dumps({
                    "ok": False,
                    "phase": "failed",
                    "errors": ["无法唯一确定基础报表原版 / 财务收入表"],
                    "detected": result["files"],
                    "missing": result["missing"],
                    "ambiguous": result["ambiguous"],
                    "next_action": "请显式传入 --base-file 和 --finance-file",
                }, ensure_ascii=False, indent=2))
                return 3
        output_dir = args.output_dir or (workspace / "output")
        child_argv = [
            "build_profit_tables.py",
            "--base", str(base_file),
            "--finance", str(finance_file),
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
