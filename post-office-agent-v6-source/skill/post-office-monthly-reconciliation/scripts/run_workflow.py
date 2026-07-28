from __future__ import annotations

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


class _TerminalArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        raise ValueError(message)


def _terminal(
    ok: bool,
    phase: str,
    *,
    target_month: str | None = None,
    outputs=(),
    metrics: dict[str, Any] | None = None,
    issues: dict[str, Any] | None = None,
    warnings=(),
    errors=(),
    audit: dict[str, Any] | None = None,
    next_action: str,
) -> dict[str, Any]:
    return {
        "ok": ok,
        "terminal": True,
        "phase": phase,
        "target_month": target_month,
        "outputs": [str(Path(path)) for path in outputs],
        "metrics": metrics or {},
        "issues": issues or {},
        "warnings": list(warnings),
        "errors": list(errors),
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
) -> dict[str, Any]:
    if phase not in {"review", "final"}:
        return _terminal(
            False,
            "failed",
            errors=["phase 必须是 review 或 final"],
            next_action="请使用受支持的工作流阶段",
        )

    if phase == "review" and not month:
        return _terminal(
            False,
            "failed",
            errors=["review 阶段必须显式提供 month"],
            next_action="请明确提供目标月份 YYYY-MM",
        )

    workspace = Path(workspace).resolve()
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
        inspection = inspect_inputs(workspace, output_dir, month, assets_dir=assets)
        if not inspection["ok"]:
            return _terminal(
                False,
                "failed",
                target_month=inspection.get("target_month"),
                warnings=inspection.get("warnings", ()),
                errors=inspection["fatal_errors"],
                next_action="请按阻断原因补充或修正当月文件",
            )
        model_summary, review = build_review_workbook(workspace, output_dir, month, assets_dir=assets)
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
        return _terminal(
            True,
            "awaiting_review",
            target_month=state["target_month"],
            outputs=[model_summary, review],
            metrics=state.get("metrics"),
            issues=state.get("issue_counts"),
            warnings=validation["warnings"],
            next_action="请人工复核并重新上传基础报表初版",
        )

    if review_file is None:
        return _terminal(
            False,
            "failed",
            errors=["final 阶段必须提供 review_file"],
            next_action="请上传人工复核后的基础报表初版",
        )
    review_path = Path(review_file)
    if not review_path.is_absolute():
        review_path = workspace / review_path
    review_path = review_path.resolve()
    try:
        review_path.relative_to(workspace)
    except ValueError:
        return _terminal(
            False,
            "failed",
            errors=["review_file 必须位于当前 workspace 内"],
            next_action="请上传当前任务 workspace 内的人工修订初版",
        )
    outputs = build_final_reports(review_path, output_dir, assets_dir=assets, previous_first_profit=previous_first_profit, previous_second_profit=previous_second_profit)
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
    return _terminal(
        True,
        "final_completed",
        target_month=state["target_month"],
        outputs=outputs,
        metrics=state.get("final_metrics"),
        issues=state.get("issue_counts"),
        warnings=validation["warnings"],
        audit=audit,
        next_action="请下载并交付三个最终月度报表",
    )


def run_workflow(
    phase: str,
    workspace: Path | str,
    month: str | None = None,
    review_file: Path | str | None = None,
    bundle_root: Path | str | None = None,
    previous_first_profit: Path | str | None = None,
    previous_second_profit: Path | str | None = None,
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
        )
    except Exception as exc:
        return _terminal(
            False,
            "failed",
            errors=[str(exc)],
            next_action="请根据错误修正输入或重新部署完整 Skill",
        )


def main(argv: list[str] | None = None) -> int:
    parser = _TerminalArgumentParser(description="执行月度委办运输报表单次工作流")
    parser.add_argument("--phase", required=True)
    parser.add_argument("--workspace", required=True, type=Path)
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
        )
    except Exception as exc:
        result = _terminal(
            False,
            "failed",
            errors=[str(exc)],
            next_action="请根据错误修正输入或重新部署完整 Skill",
        )
    print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
    return 0 if result["ok"] else 2


if __name__ == "__main__":
    sys.exit(main())
