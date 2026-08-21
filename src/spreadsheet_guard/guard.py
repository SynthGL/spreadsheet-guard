"""Read-only workbook Guard surface backed by WolfXL."""

from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass
from importlib import import_module
from pathlib import Path
from typing import Any, cast


class GuardError(RuntimeError):
    """Base error for the public Guard surface."""


class GuardUnavailableError(GuardError):
    """Raised when the WolfXL Guard runtime is not installed."""


class GuardExecutionError(GuardError):
    """Raised when WolfXL does not produce a valid Guard report."""


@dataclass(frozen=True)
class GuardOutcome:
    """Stable summary of one read-only workbook comparison."""

    status: str
    output_path: Path
    report: Mapping[str, Any]


GuardFunction = Callable[..., dict[str, Any]]


def _load_wolfxl_guard() -> GuardFunction:
    try:
        module = import_module("wolfxl.operations")
    except ImportError as exc:
        raise GuardUnavailableError(
            "WolfXL Guard requires the WolfXL Commercial runtime. "
            "WolfXL Community does not include the operations SDK. "
            "Request evaluation access at https://wolfxl.com"
        ) from exc
    run_guard = getattr(module, "run_guard", None)
    if not callable(run_guard):
        raise GuardUnavailableError("the installed WolfXL runtime does not expose Guard")
    return cast(GuardFunction, run_guard)


def guard_workbooks(
    before: Path,
    after: Path,
    output: Path,
    *,
    policy: Mapping[str, Any] | None = None,
) -> GuardOutcome:
    """Compare two workbooks without modifying either input.

    WolfXL owns workbook inspection and policy semantics. This package owns only
    the stable public invocation and fail-closed result contract.
    """

    runner = _load_wolfxl_guard()
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
            f"WolfXL Guard failed ({type(exc).__name__}): {exc}"
        ) from exc

    if not isinstance(report, Mapping):
        raise GuardExecutionError("WolfXL Guard returned a non-object report")
    status = report.get("status")
    if status not in {"passed", "failed", "unassessed"}:
        raise GuardExecutionError(f"WolfXL Guard returned invalid status {status!r}")
    if not output.is_file():
        raise GuardExecutionError("WolfXL Guard did not write the requested report")

    return GuardOutcome(status=status, output_path=output, report=dict(report))
