from __future__ import annotations
from reconciliation import xml_compat  # noqa: F401  必须在 openpyxl 之前导入
from reconciliation.paths import SKIP_DIRS, resolve_input_dir

import argparse
import json
import os
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


def infer_month_from_inputs(input_dir: Path | str) -> dict:
    """从上传文件推断目标月份，让用户不必记 YYYY-MM。

    关键点：财务收入表常常是「上个月」的（业务上允许），不能拿它当目标月。
    所以按角色加权——车管/容积测量/基础报表这些「业务月」证据权重大，
    财务表权重小；只有在没有其他证据时才用财务表。

    依据优先级：
      1. 文件名里的月份（9月.xlsx、2026.08运输局.xlsx）
      2. 工作表名里的月份（草滩的 9月1日、港务的 9.1）

    返回结构（不再「猜一个就完事」）：
        {
          "month": "2026-09" | None,
          "month_only": 9,              # 只推断了月，没有年份
          "year_confirmed": bool,       # 年份是否来自文件而非默认
          "confidence": "high"|"medium"|"low",
          "evidence": [{"file":..., "month":..., "year":...}],
          "conflicts": [...],           # 不同月份的证据，供反问用户
          "questions": [...]            # 需要 agent 主动反问用户的点
        }
    只要 confidence != "high"，调用方就必须先反问用户，不能直接往下跑。
    """
    from collections import Counter

    from reconciliation.excel_io import discover_workbooks
    from reconciliation.normalize import parse_month_token

    root = Path(input_dir)
    result = {
        "month": None, "month_only": None, "year_confirmed": False,
        "confidence": "low", "evidence": [], "conflicts": [], "questions": [], "notes": [],
    }
    if not root.is_dir():
        result["questions"].append("没有找到上传文件，请先上传当月报表")
        return result

    # 财务表月份可能滞后，权重低；其余业务表权重高
    LOW_WEIGHT_HINTS = ("运输局", "财务", "收入")

    counter: Counter[tuple[int | None, int]] = Counter()
    for path in discover_workbooks(root):
        token = parse_month_token(path.name)
        if not token:
            continue
        weight = 1 if any(hint in path.name for hint in LOW_WEIGHT_HINTS) else 3
        counter[token] += weight
        result["evidence"].append({
            "file": path.name, "month": token[1], "year": token[0], "weight": weight,
        })

    if not counter:
        # 文件名没有月份，退回读工作表名（只读 sheet 名，不加载数据）
        try:
            import openpyxl

            for path in discover_workbooks(root):
                try:
                    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                except Exception:
                    continue
                try:
                    for name in wb.sheetnames:
                        token = parse_month_token(name)
                        if token:
                            counter[token] += 1
                            result["evidence"].append({
                                "file": f"{path.name}#{name}", "month": token[1],
                                "year": token[0], "weight": 1,
                            })
                finally:
                    wb.close()
        except Exception:
            pass

    if not counter:
        result["questions"].append(
            "文件名和工作表名里都没有月份信息，请告诉我这是哪个月的报表"
        )
        return result

    # 按月聚合权重，取最高分
    month_score: Counter[int] = Counter()
    for (year, mon), weight in counter.items():
        month_score[mon] += weight
    best = max(month_score.values())
    winners = [mon for mon, score in month_score.items() if score == best]

    if len(winners) != 1:
        result["conflicts"] = [
            {"month": mon, "score": score} for mon, score in month_score.most_common()
        ]
        result["questions"].append(
            "上传的文件里出现了多个不同的月份（"
            + "、".join(f"{mon}月" for mon in sorted(month_score))
            + "），请确认要核对的是哪个月"
        )
        return result

    month = winners[0]
    # 次高分只差一点点时也要问，避免两个月份证据接近却硬选一个
    second = sorted(month_score.values(), reverse=True)
    if len(second) > 1 and second[1] > 0 and (best - second[1]) <= 1:
        result["conflicts"] = [
            {"month": mon, "score": score} for mon, score in month_score.most_common()
        ]
        result["questions"].append(
            f"{month}月和其他月份的线索接近，请确认要核对的是不是 {month} 月"
        )

    # 年份判定：优先取「同一月份」的证据；该月没有年份时，用其他文件的年份兜底
    # （财务表常写上月的完整日期，年份同样有效）。全都没有才需要反问。
    years = {year for (year, mon) in counter if year is not None and mon == month}
    result["month_only"] = month
    if len(years) == 1:
        year = next(iter(years))
        result["year_confirmed"] = True
    elif not years:
        other_years = {year for (year, _mon) in counter if year is not None}
        if len(other_years) == 1:
            # 从别的文件借到唯一确定的年份：能定下来，只需在汇报里说明依据
            year = next(iter(other_years))
            result["year_confirmed"] = True
            result.setdefault("notes", []).append(
                f"{month} 月的文件本身没写年份，参照同批文件的 {year} 年"
            )
        else:
            year = datetime.now().year
            result["year_confirmed"] = False
            result["questions"].append(
                f"文件里只有「{month}月」没有年份，我暂按 {year} 年处理，"
                f"如果不对请告诉我正确年份"
            )
    else:
        year = datetime.now().year
        result["year_confirmed"] = False
        result["questions"].append(
            f"{month} 月的文件出现了多个年份（{'、'.join(str(y) for y in sorted(years))}），"
            "请确认用哪一年"
        )

    result["month"] = f"{year:04d}-{month:02d}"
    if result["conflicts"]:
        result["confidence"] = "low"
    elif result["year_confirmed"]:
        result["confidence"] = "high"
    else:
        result["confidence"] = "medium"
    return result


def _scan_input_roots(workspace: Path, skill_root: Path) -> dict:
    """诊断上传文件的实际落盘位置。

    不同版本的平台把附件放在不同目录，只认 `/workspace` 会在部署变更后
    误报「缺少全部输入」。这里列出几个候选根目录下的 Excel 文件，供排查。
    """
    candidates = [
        workspace,
        Path("/workspace"),
        Path("/workspace/uploads"),
        Path("/mnt/data"),
        Path("/tmp"),
        Path("/mnt"),
        Path("/home"),
        Path("/root"),
        Path("/data"),
        Path.cwd(),
        skill_root.parent,
        Path("/"),
    ]
    seen: set[str] = set()
    roots = []
    for root in candidates:
        try:
            key = str(root.resolve())
        except OSError:
            continue
        if key in seen or not root.is_dir():
            continue
        seen.add(key)
        matches: list[str] = []
        for dirpath, dirnames, filenames in os.walk(root):
            depth = len(Path(dirpath).relative_to(root).parts)
            max_depth = 2 if root == Path("/") else 3
            if depth >= max_depth:
                dirnames[:] = []
            dirnames[:] = [name for name in dirnames if name not in SKIP_DIRS and not name.startswith(".")]
            for name in sorted(filenames):
                if name.startswith("~$") or not name.lower().endswith((".xlsx", ".xls")):
                    continue
                matches.append(str(Path(dirpath) / name))
                if len(matches) >= 50:
                    break
            if len(matches) >= 50:
                break
        roots.append({"root": str(root), "excel_count": len(matches), "samples": matches[:10]})
    return {"cwd": os.getcwd(), "roots": roots}


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
    # 旧版平台把附件挂在 /workspace，新版挂在 /inputs。两者都接受。
    input_dir = resolve_input_dir(workspace)

    # AgentOS 沙箱中 skill 包放在 /workspace/skills/<name>/，
    # 需要在 sys.path 中注册才能让脚本内的 from reconciliation.xxx import 生效
    sys.path.insert(0, str(skill_root))
    sys.path.insert(0, str(scripts_dir))

    if args.phase == "list":
        # 列出工作区文件，供调用方确认真实文件名。
        # 用途：避免模型照抄示例文件名或臆造 /inputs、/outputs 这类不存在的路径。
        files = []
        if input_dir.is_dir():
            for path in sorted(input_dir.rglob("*")):
                if not path.is_file() or path.name.startswith("~$"):
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
            "input_dir": str(input_dir),
            "exists": workspace.is_dir(),
            "file_count": len(files),
            "files": files,
            "diagnostics": _scan_input_roots(workspace, skill_root),
            "hint": "下一步执行 --phase detect 按表结构识别文件角色；脚本参数只使用 resolved 中的真实路径。",
        }, ensure_ascii=False, indent=2))
        return 0 if workspace.is_dir() else 2

    if args.phase == "detect":
        # 按表结构识别工作区里每个文件是什么，不依赖文件名。
        from reconciliation.detect import describe_input_gaps, detect_for_phase, needs_input_message

        # 没传月份时先从文件内容推断，不要让用户记 YYYY-MM：
        # 文件名里的 9月/2026.08、以及工作表名里的 9月1日 都能给出月份。
        inference = None
        if not args.month:
            inference = infer_month_from_inputs(input_dir)
            if inference["month"] is None:
                # 推不出来就别往下跑，让 agent 用大白话反问用户
                print(json.dumps({
                    "ok": True,
                    "status": "needs_month",
                    "phase": "needs_month",
                    "target_month": None,
                    "questions": inference["questions"],
                    "month_inference": inference,
                    "hint": "请把这些疑问原样转告用户，拿到答复后再重跑 detect",
                }, ensure_ascii=False))
                return 0
            args.month = inference["month"]
        result = detect_for_phase(input_dir, args.month, args.task)
        if inference:
            result["inferred_month"] = inference["month"]
            result["month_inference"] = inference
            if inference["confidence"] != "high":
                # 不确定就要问：把待确认点放在 questions 里，agent 必须转述
                result["status"] = "needs_confirm"
                result["questions"] = inference["questions"]
                result.setdefault("warnings", []).append(
                    f"月份是按文件推断的（置信度 {inference['confidence']}），"
                    "请与用户确认后再继续"
                )
            else:
                result.setdefault("warnings", []).append(
                    f"未指定月份，已根据上传文件推断目标月为 {inference['month']}"
                )
        gaps = describe_input_gaps(result)
        if gaps:
            # 输入不全不是执行失败：返回 needs_input，让 Agent 请用户补文件。
            payload = {
                "ok": True,
                "status": "needs_input",
                "phase": "awaiting_inputs",
                "task_phase": result["phase"],
                "target_month": args.month,
                "missing_inputs": gaps,
                "missing": result.get("missing", []),
                "ambiguous": result.get("ambiguous", []),
                "message": needs_input_message(gaps),
                "next_action": "请补充或确认上述文件后重新执行；已识别到的文件会保留。",
                **{k: v for k, v in result.items()
                   if k not in {"phase", "files", "missing_inputs", "message", "next_action",
                                "resolved", "missing", "ambiguous"}},
                "files": result.get("files", []),
                "resolved": result.get("resolved", {}),
            }
            print(f"[STOP] 还缺输入：{'、'.join(item['label'] for item in gaps)}。"
                  "按 missing_inputs 请用户补齐，不要自己猜文件名。")
        else:
            payload = {
                "ok": True,
                "status": "ready",
                "phase": "detected",
                "task_phase": result["phase"],
                "target_month": args.month,
                "resolved": result.get("resolved", {}),
                "missing_inputs": [],
                "missing": result.get("missing", []),
                "ambiguous": result.get("ambiguous", []),
                "next_action": "输入齐全，可继续执行该阶段。",
                **{k: v for k, v in result.items()
                   if k not in {"phase", "files", "missing", "ambiguous", "resolved"}},
                "files": result.get("files", []),
            }
            names = [Path(p).name for paths in result.get("resolved", {}).values() for p in paths]
            print(f"[OK] 识别完成（目标月 {args.month}）：{'、'.join(names)}。"
                  "输入齐全，直接执行下一步，不要重跑 detect。")
        # 紧凑输出：平台会按长度截断工具结果，缩进版体积翻倍，
        # 一旦被截断模型看不到结论，就会反复重跑同一阶段（线上真实踩过）。
        print(json.dumps(payload, ensure_ascii=False))
        return 0

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
        # 只做输入预检，不生成报表。上传目录可能不是 workspace（新版在 /inputs），
        # run-state 始终落在 workspace/output。
        output_dir = args.output_dir or (workspace / "output")
        output_dir.mkdir(parents=True, exist_ok=True)
        child_argv = [
            "inspect_inputs.py",
            "--input-dir", str(input_dir),
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
            inference = infer_month_from_inputs(input_dir)
            if inference["month"] is None:
                print(json.dumps({
                    "ok": True, "phase": "needs_month", "status": "needs_month",
                    "questions": inference["questions"],
                    "month_evidence": inference["evidence"],
                    "conflicts": inference["conflicts"],
                }, ensure_ascii=False))
                return 0
            args.month = inference["month"]
        caotan, gangwu = args.caotan, args.gangwu
        if not caotan or not gangwu:
            # 未指定文件时，按表结构自动识别草滩/港务在线表
            from reconciliation.detect import (
                describe_input_gaps,
                detect_for_phase,
                needs_input_message,
            )

            result = detect_for_phase(input_dir, args.month, "summary")
            resolved = result["resolved"]
            if not caotan and len(resolved.get("online_caotan", [])) == 1:
                caotan = Path(resolved["online_caotan"][0])
            if not gangwu and len(resolved.get("online_gangwu", [])) == 1:
                gangwu = Path(resolved["online_gangwu"][0])
            if not caotan or not gangwu:
                gaps = describe_input_gaps(result)
                if not gaps:
                    gaps = [{
                        "kind": "ambiguous",
                        "role": "online",
                        "label": "草滩/港务在线测量表",
                        "action": "请显式传入 --caotan 和 --gangwu，或分别上传两张在线测量表",
                    }]
                print(json.dumps({
                    "ok": True,
                    "status": "needs_input",
                    "phase": "awaiting_inputs",
                    "task_phase": "summary",
                    "detected": result["files"],
                    "missing": result["missing"],
                    "ambiguous": result["ambiguous"],
                    "missing_inputs": gaps,
                    "message": needs_input_message(gaps),
                    "next_action": "请补充或确认草滩/港务在线测量表后重试。",
                }, ensure_ascii=False, indent=2))
                return 0
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
            "--input-dir", str(input_dir),
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
                "ok": True,
                "status": "needs_input",
                "phase": "awaiting_inputs",
                "task_phase": "final",
                "missing_inputs": [{
                    "kind": "missing",
                    "role": "review_file",
                    "label": "人工核对后的基础报表初版",
                    "required": 1,
                    "found": 0,
                    "action": "请上传人工核对修正后的基础报表初版",
                }],
                "message": "当前上传的文件不完整，需要补充或确认：人工核对后的基础报表初版。",
                "next_action": "请上传人工核对修正后的基础报表初版后重试。",
            }, ensure_ascii=False, indent=2))
            return 0
        child_argv = [
            "run_workflow.py",
            "--phase", "final",
            "--workspace", str(workspace),
            "--input-dir", str(input_dir),
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
            inference = infer_month_from_inputs(input_dir)
            if inference["month"] is None:
                print(json.dumps({
                    "ok": True, "phase": "needs_month", "status": "needs_month",
                    "questions": inference["questions"],
                    "month_evidence": inference["evidence"],
                    "conflicts": inference["conflicts"],
                }, ensure_ascii=False))
                return 0
            args.month = inference["month"]
        base_file, finance_file = args.base_file, args.finance_file
        if not base_file or not finance_file:
            # 未指定文件时，按表结构自动识别：含「一干汇总/二干基础」的是基础报表，
            # 表头含 结算收入 + 普服收入 的是财务收入表（客户可能叫「运输局」）。
            from reconciliation.detect import (
                describe_input_gaps,
                detect_for_phase,
                needs_input_message,
            )

            result = detect_for_phase(input_dir, args.month, "profit")
            resolved = result["resolved"]
            if not base_file and len(resolved.get("base_report", [])) == 1:
                base_file = Path(resolved["base_report"][0])
            if not finance_file and len(resolved.get("finance", [])) == 1:
                finance_file = Path(resolved["finance"][0])
            if not base_file or not finance_file:
                gaps = describe_input_gaps(result)
                if not gaps:
                    gaps = [{
                        "kind": "ambiguous",
                        "role": "base_report/finance",
                        "label": "基础报表原版 / 财务收入表",
                        "action": "请确认或显式指定基础报表原版和财务收入表",
                    }]
                print(json.dumps({
                    "ok": True,
                    "status": "needs_input",
                    "phase": "awaiting_inputs",
                    "task_phase": "profit",
                    "detected": result["files"],
                    "missing": result["missing"],
                    "ambiguous": result["ambiguous"],
                    "missing_inputs": gaps,
                    "message": needs_input_message(gaps),
                    "next_action": "请补充或确认基础报表原版和财务收入表后重试。",
                }, ensure_ascii=False, indent=2))
                return 0
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
    try:
        exit_code = main()
    except SystemExit as exc:
        if isinstance(exc.code, int):
            exit_code = exc.code
        else:
            exit_code = 0 if exc.code is None else 1
    if os.environ.get("POST_OFFICE_HARD_EXIT") == "1":
        # Windows + openpyxl 偶发在解释器收尾阶段崩溃（0xC0000005，输出已完整打印）。
        # 只有显式设置该环境变量时才跳过解释器终结流程（测试/CLI 场景）。
        # 默认不启用：AgentOS 是通过 runpy 在宿主进程内调用本入口的，
        # 此时 os._exit 会把宿主进程一起杀掉。
        sys.stdout.flush()
        sys.stderr.flush()
        os._exit(exit_code)
    raise SystemExit(exit_code)
