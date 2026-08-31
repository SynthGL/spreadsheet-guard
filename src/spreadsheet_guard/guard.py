"""Read-only workbook Guard surface backed by the bundled audit engine."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from spreadsheet_guard._engine import run_guard as _run_guard


class GuardError(RuntimeError):
    """Base error for the public Guard surface."""


class GuardExecutionError(GuardError):
    """Raised when the audit engine does not produce a valid Guard report."""


@dataclass(frozen=True)
class GuardOutcome:
    """Stable summary of one read-only workbook comparison."""

    status: str
    output_path: Path
    report: Mapping[str, Any]


def guard_workbooks(
    before: Path,
    after: Path,
    output: Path,
    *,
    policy: Mapping[str, Any] | None = None,
) -> GuardOutcome:
    """Compare two workbooks without modifying either input.

    The bundled engine owns workbook inspection and policy semantics. This
    module owns the stable public invocation and fail-closed result contract.
    """

    runner = _run_guard
    arguments: dict[str, Any] = {
        "before": before,
        "after": after,
        "output": output,
    }
    if policy is not None:
        arguments["policy"] = policy

    try:
        report = runner(**arguments)
    except Exception as exc:
        if isinstance(exc, GuardError):
            raise
        raise GuardExecutionError(
            f"workbook Guard failed ({type(exc).__name__}): {exc}"
        ) from exc

    if not isinstance(report, Mapping):
        raise GuardExecutionError("workbook Guard returned a non-object report")
    status = report.get("status")
    if status not in {"passed", "failed", "unassessed"}:
        raise GuardExecutionError(f"workbook Guard returned invalid status {status!r}")
    if not output.is_file():
        raise GuardExecutionError("workbook Guard did not write the requested report")

    return GuardOutcome(status=status, output_path=output, report=dict(report))
