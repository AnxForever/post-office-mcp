from __future__ import annotations
from reconciliation import xml_compat  # noqa: F401  必须在 openpyxl 之前导入

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from build_final_reports import build_final_reports
from build_review_workbook import build_review_workbook
from check_runtime import REQUIRED_REFERENCES, run_checks
from inspect_inputs import inspect_inputs
from validate_outputs import validate_outputs
from reconciliation.integrity import file_sha256
from reconciliation.paths import publish_files, resolve_input_dir


class _TerminalArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        raise ValueError(message)


def _describe_exception(exc: BaseException) -> str:
    """错误信息必须可读：MemoryError 这类异常的 str() 是空串，直接上报等于没报。"""
    if isinstance(exc, MemoryError):
        return "内存不足，工作簿读取被中断；请关闭占用内存的程序后重试（未生成任何交付件）"
    return str(exc) or type(exc).__name__


def _terminal(
    ok: bool,
    phase: str,
    *,
    status: str | None = None,
    target_month: str | None = None,
    outputs=(),
    metrics: dict[str, Any] | None = None,
    issues: dict[str, Any] | None = None,
    warnings=(),
    errors=(),
    audit: dict[str, Any] | None = None,
    published=(),
    missing_inputs=(),
    next_action: str,
) -> dict[str, Any]:
    if status is None:
        status = "completed" if ok and phase in {"awaiting_review", "final_completed"} else (
            "failed" if not ok else "ok"
        )
    return {
        "ok": ok,
        "terminal": True,
        "status": status,
        "phase": phase,
        "target_month": target_month,
        "outputs": [str(Path(path)) for path in outputs],
        "published": list(published),
        "metrics": metrics or {},
        "issues": issues or {},
        "warnings": list(warnings),
        "errors": list(errors),
        "missing_inputs": list(missing_inputs),
        "audit": audit or {},
        "next_action": next_action,
    }


def _final_audit(state: dict[str, Any], validation: dict[str, Any]) -> dict[str, Any]:
    expected_hashes = state.get("final_output_hashes", {})
    output_verification = []
    for item in state.get("final_outputs", []):
        path = Path(item).resolve()
        expected = expected_hashes.get(path.name)
        output_verification.append({
            "path": str(path),
            "sha256": expected,
            "verified": bool(path.is_file() and expected and file_sha256(path) == expected),
        })
    checks = validation.get("checks", {})
    return {
        "finance_classification": state.get("finance_classification", {}),
        "special_sheet_status": state.get("special_sheet_status", {}),
        "profit_value_groups_checked": checks.get("profit_value_groups_checked", 0),
        "profit_spot_checks": checks.get("profit_spot_checks", []),
        "output_verification": output_verification,
        "validation_passed": bool(validation.get("ok")),
    }


def _run_workflow(
    phase: str,
    workspace: Path | str,
    month: str | None = None,
    review_file: Path | str | None = None,
    bundle_root: Path | str | None = None,
    previous_first_profit: Path | str | None = None,
    previous_second_profit: Path | str | None = None,
    input_dir: Path | str | None = None,
) -> dict[str, Any]:
    if phase not in {"review", "final"}:
        return _terminal(
            False,
            "failed",
            errors=["phase 必须是 review 或 final"],
            next_action="请使用受支持的工作流阶段",
        )

    workspace = Path(workspace).resolve()
    # 上传附件可能挂载在 /inputs（新版平台）而不是 workspace；输出始终写 workspace/output。
    input_dir = Path(input_dir).resolve() if input_dir is not None else resolve_input_dir(workspace)

    # review 没给月份时从文件内容推断，别让用户记 YYYY-MM。
    # 推断不确定时不能硬跑：返回 needs_confirm 并把疑问交回给 agent 反问用户。
    if phase == "review" and not month:
        from agentos_entrypoint import infer_month_from_inputs

        inference = infer_month_from_inputs(input_dir)
        if inference["month"] is None:
            return _terminal(
                True,
                "needs_month",
                status="needs_month",
                warnings=tuple(inference["questions"]),
                errors=(),
                next_action="请把这些疑问转告用户，拿到答复后重跑",
            )
        month = inference["month"]
        if inference["confidence"] != "high":
            return _terminal(
                True,
                "needs_confirm",
                status="needs_confirm",
                target_month=month,
                warnings=tuple(inference["questions"]),
                errors=(),
                next_action="月份是按文件推断的，请与用户确认后再执行 review",
            )
    bundle_root = Path(bundle_root).resolve() if bundle_root is not None else Path(__file__).resolve().parents[1]
    output_dir = workspace / "output"
    assets = bundle_root / "assets"
    runtime = run_checks(
        workspace,
        bundle_root,
        phase=phase,
        require_agentos_layout=workspace == Path("/workspace").resolve(),
    )
    if not runtime["ok"]:
        return _terminal(
            False,
            "failed",
            errors=runtime["errors"],
            next_action="请修复沙箱环境或重新部署 Skill",
        )
    for name in REQUIRED_REFERENCES:
        content = (bundle_root / "references" / name).read_text(encoding="utf-8")
        if not content.strip():
            return _terminal(
                False,
                "failed",
                errors=[f"Skill 规则为空: {name}"],
                next_action="请重新部署完整 Skill",
            )

    if phase == "review":
        inspection = inspect_inputs(input_dir, output_dir, month, assets_dir=assets)
        if not inspection["ok"]:
            gaps = inspection.get("missing_inputs", ())
            if gaps:
                # 上传不全属于正常等待用户补充，不是执行失败。
                return _terminal(
                    True,
                    "awaiting_inputs",
                    status="needs_input",
                    target_month=inspection.get("target_month"),
                    warnings=inspection.get("warnings", ()),
                    errors=(),
                    missing_inputs=gaps,
                    next_action="请补充或确认上传的文件后重试；已识别到的文件会保留。",
                )
            return _terminal(
                False,
                "failed",
                target_month=inspection.get("target_month"),
                warnings=inspection.get("warnings", ()),
                errors=inspection["fatal_errors"],
                next_action="请按阻断原因补充或修正当月文件",
            )
        try:
            model_summary, review = build_review_workbook(input_dir, output_dir, month, assets_dir=assets)
        except FileNotFoundError as exc:
            # 预检通过后仍缺文件，通常是上传目录在两次调用间变化或文件被移除。
            return _terminal(
                True,
                "awaiting_inputs",
                status="needs_input",
                target_month=inspection.get("target_month"),
                warnings=inspection.get("warnings", ()),
                errors=(),
                missing_inputs=[{
                    "kind": "missing",
                    "role": "input",
                    "label": "当月输入文件",
                    "action": f"输入文件已不可用，请重新上传后重试：{exc}",
                }],
                next_action="请重新上传当月输入文件后重试。",
            )
        validation = validate_outputs(output_dir, "review")
        state = json.loads((output_dir / "run-state.json").read_text(encoding="utf-8"))
        if not validation["ok"]:
            return _terminal(
                False,
                "failed",
                target_month=state.get("target_month"),
                warnings=validation["warnings"],
                errors=validation["errors"],
                next_action="初版验证失败，请修复 Skill 后重新执行",
            )
        published = publish_files([model_summary, review])
        return _terminal(
            True,
            "awaiting_review",
            target_month=state["target_month"],
            outputs=[model_summary, review],
            published=published,
            metrics=state.get("metrics"),
            issues=state.get("issue_counts"),
            warnings=validation["warnings"],
            next_action="请人工复核并重新上传基础报表初版",
        )

    if review_file is None:
        return _terminal(
            True,
            "awaiting_inputs",
            status="needs_input",
            missing_inputs=[{
                "kind": "missing",
                "role": "review_file",
                "label": "人工核对后的基础报表初版",
                "required": 1,
                "found": 0,
                "action": "请上传人工核对修正后的基础报表初版",
            }],
            next_action="请上传人工核对修正后的基础报表初版后重试。",
        )
    review_path = Path(review_file)
    if not review_path.is_absolute():
        review_path = input_dir / review_path
    review_path = review_path.resolve()
    if not review_path.is_file():
        return _terminal(
            True,
            "awaiting_inputs",
            status="needs_input",
            missing_inputs=[{
                "kind": "missing",
                "role": "review_file",
                "label": "人工核对后的基础报表初版",
                "required": 1,
                "found": 0,
                "action": f"没有找到初版文件：{review_path.name}，请重新上传",
            }],
            next_action="请上传人工核对修正后的基础报表初版后重试。",
        )
    # 上传目录可能是 /inputs，也可能是 workspace；两者都允许。
    allowed = False
    for root in (input_dir, workspace):
        try:
            review_path.relative_to(root)
            allowed = True
            break
        except ValueError:
            continue
    if not allowed:
        return _terminal(
            True,
            "awaiting_inputs",
            status="needs_input",
            missing_inputs=[{
                "kind": "wrong_location",
                "role": "review_file",
                "label": "人工核对后的基础报表初版",
                "action": "初版文件不在当前任务的上传目录或 workspace 内，请重新上传到当前任务",
            }],
            next_action="请重新上传当前任务可见的人工修订初版。",
        )
    try:
        outputs = build_final_reports(review_path, output_dir, assets_dir=assets, previous_first_profit=previous_first_profit, previous_second_profit=previous_second_profit)
    except FileNotFoundError as exc:
        return _terminal(
            True,
            "awaiting_inputs",
            status="needs_input",
            missing_inputs=[{
                "kind": "prerequisite",
                "role": "run_state",
                "label": "同一任务上一阶段的运行状态（run-state.json）",
                "action": "请先在同一任务内完成 review 阶段，再执行 final；或确认初版和当月输入文件都在当前任务里",
            }],
            next_action="请先在同一任务内完成 review 阶段后重试 final。",
        )
    validation = validate_outputs(output_dir, "final")
    state = json.loads((output_dir / "run-state.json").read_text(encoding="utf-8"))
    audit = _final_audit(state, validation)
    if not validation["ok"]:
        return _terminal(
            False,
            "failed",
            target_month=state.get("target_month"),
            warnings=validation["warnings"],
            errors=validation["errors"],
            audit=audit,
            next_action="最终文件验证失败，请修复 Skill 后重新执行",
        )
    published = publish_files(outputs)
    return _terminal(
        True,
        "final_completed",
        target_month=state["target_month"],
        outputs=outputs,
        published=published,
        metrics=state.get("final_metrics"),
        issues=state.get("issue_counts"),
        warnings=validation["warnings"],
        audit=audit,
        next_action="请从「成果文件」下载并交付三个最终月度报表",
    )


def run_workflow(
    phase: str,
    workspace: Path | str,
    month: str | None = None,
    review_file: Path | str | None = None,
    bundle_root: Path | str | None = None,
    previous_first_profit: Path | str | None = None,
    previous_second_profit: Path | str | None = None,
    input_dir: Path | str | None = None,
) -> dict[str, Any]:
    try:
        return _run_workflow(
            phase,
            workspace,
            month,
            review_file,
            bundle_root,
            previous_first_profit,
            previous_second_profit,
            input_dir,
        )
    except Exception as exc:
        return _terminal(
            False,
            "failed",
            errors=[_describe_exception(exc)],
            next_action="请根据错误修正输入或重新部署完整 Skill",
        )


def main(argv: list[str] | None = None) -> int:
    parser = _TerminalArgumentParser(description="执行月度委办运输报表单次工作流")
    parser.add_argument("--phase", required=True)
    parser.add_argument("--workspace", required=True, type=Path)
    parser.add_argument("--input-dir", type=Path, help="上传目录（默认与 workspace 相同）")
    parser.add_argument("--month")
    parser.add_argument("--review-file", type=Path)
    parser.add_argument("--bundle-root", type=Path)
    parser.add_argument("--previous-first-profit", type=Path, default=None)
    parser.add_argument("--previous-second-profit", type=Path, default=None)
    try:
        args = parser.parse_args(argv)
        result = run_workflow(
            args.phase,
            args.workspace,
            month=args.month,
            review_file=args.review_file,
            bundle_root=args.bundle_root,
            previous_first_profit=args.previous_first_profit,
            previous_second_profit=args.previous_second_profit,
            input_dir=args.input_dir,
        )
    except Exception as exc:
        result = _terminal(
            False,
            "failed",
            errors=[_describe_exception(exc)],
            next_action="请根据错误修正输入或重新部署完整 Skill",
        )
    print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
    return 0 if result["ok"] else 2


if __name__ == "__main__":
    sys.exit(main())
