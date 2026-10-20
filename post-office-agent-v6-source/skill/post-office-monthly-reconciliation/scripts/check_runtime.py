from __future__ import annotations
from reconciliation import xml_compat  # noqa: F401  必须在 openpyxl 之前导入

import argparse
import importlib
import json
import tempfile
from io import BytesIO
from pathlib import Path
from typing import Any, Callable, Sequence


REQUIRED_ASSETS = (
    "split-report-template.xlsx",
    "company-map.json",
    "route-aliases.json",
)

REVIEW_ASSETS = REQUIRED_ASSETS

REQUIRED_REFERENCES = (
    "input-contract.md",
    "reconciliation-rules.md",
)

REQUIRED_SCRIPTS = (
    "agentos_runner.py",
    "build_model_summary.py",
    "build_profit_tables.py",
    "check_runtime.py",
    "run_workflow.py",
    "inspect_inputs.py",
    "build_review_workbook.py",
    "build_final_reports.py",
    "validate_outputs.py",
    "reconciliation/__init__.py",
    "reconciliation/calculations.py",
    "reconciliation/excel_io.py",
    "reconciliation/integrity.py",
    "reconciliation/matching.py",
    "reconciliation/model_summary.py",
    "reconciliation/monthly_comparison.py",
    "reconciliation/normalize.py",
    "reconciliation/profit_audit.py",
    "reconciliation/template_discovery.py",
    "reconciliation/templates.py",
)


def _check_readable_files(root: Path, names: Sequence[str], label: str, errors: list[str]) -> int:
    readable = 0
    for name in names:
        path = root / name
        try:
            if not path.is_file():
                raise FileNotFoundError(path)
            with path.open("rb") as handle:
                handle.read(1)
            readable += 1
        except OSError as exc:
            errors.append(f"{label}不可读: {name} ({exc})")
    return readable


def run_checks(
    workspace: Path | str,
    bundle_root: Path | str,
    *,
    phase: str = "final",
    import_module: Callable[[str], Any] = importlib.import_module,
    require_agentos_layout: bool = False,
) -> dict[str, Any]:
    workspace = Path(workspace).resolve()
    bundle_root = Path(bundle_root).resolve()
    checks: dict[str, Any] = {}
    errors: list[str] = []
    checks["dependency_mode"] = "offline-bundled"

    modules: dict[str, Any] = {}
    for name in ("openpyxl", "et_xmlfile"):
        try:
            module = import_module(name)
            modules[name] = module
            checks[f"{name}_version"] = getattr(module, "__version__", "unknown")
        except Exception as exc:  # dependency diagnostics must be serialized
            errors.append(f"依赖不可用: {name} ({exc})")

    if "openpyxl" in modules:
        try:
            workbook = modules["openpyxl"].Workbook()
            workbook.active["A1"] = "agentos"
            buffer = BytesIO()
            workbook.save(buffer)
            buffer.seek(0)
            loaded = modules["openpyxl"].load_workbook(buffer, read_only=True)
            try:
                checks["xlsx_roundtrip"] = loaded.active["A1"].value == "agentos"
            finally:
                loaded.close()
            if not checks["xlsx_roundtrip"]:
                errors.append("xlsx 内存往返值不一致")
        except Exception as exc:
            checks["xlsx_roundtrip"] = False
            errors.append(f"xlsx 内存往返失败: {exc}")
    else:
        checks["xlsx_roundtrip"] = False

    required_assets = REVIEW_ASSETS if phase == "review" else REQUIRED_ASSETS
    checks["readable_assets"] = _check_readable_files(
        bundle_root / "assets", required_assets, "Skill 资产", errors
    )
    checks["readable_references"] = _check_readable_files(
        bundle_root / "references", REQUIRED_REFERENCES, "Skill 规则", errors
    )
    checks["readable_scripts"] = _check_readable_files(
        bundle_root / "scripts", REQUIRED_SCRIPTS, "Skill 脚本", errors
    )

    if require_agentos_layout:
        expected = Path("/workspace/skills/post-office-monthly-reconciliation")
        if bundle_root != expected:
            errors.append(f"Skill 包路径错误: {bundle_root}，应为 {expected}")
    checks["bundle_root"] = str(bundle_root)

    output = workspace / "output"
    try:
        output.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(prefix=".agentos-write-check-", dir=output, delete=False) as handle:
            probe = Path(handle.name)
            handle.write(b"ok")
        probe.unlink()
        checks["output_writable"] = True
    except OSError as exc:
        checks["output_writable"] = False
        errors.append(f"output 目录不可写: {output} ({exc})")

    return {"ok": not errors, "checks": checks, "errors": errors}


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="检查月度核对 Skill 的 AgentOS 离线运行环境")
    parser.add_argument("--workspace", default="/workspace")
    parser.add_argument("--bundle-root", default=str(Path(__file__).resolve().parents[1]))
    parser.add_argument("--phase", choices=("review", "final"), default="final")
    args = parser.parse_args(argv)
    result = run_checks(
        args.workspace,
        args.bundle_root,
        phase=args.phase,
        require_agentos_layout=Path(args.workspace) == Path("/workspace"),
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
