"""Read-only workbook Guard surface backed by WolfXL."""

from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass
from importlib import import_module
from pathlib import Path
from typing import Any, cast

EXPECTED_GUARD_CONTRACT_VERSION = "wolfxl-guard-contract-v1"


class GuardError(RuntimeError):
    """Base error for the public Guard surface."""


class GuardUnavailableError(GuardError):
    """Raised when a compatible WolfXL Guard runtime is not installed."""


class GuardExecutionError(GuardError):
    """Raised when WolfXL does not produce a valid Guard result."""


@dataclass(frozen=True)
class GuardOutcome:
    """Stable summary of one read-only workbook comparison."""

    status: str
    output_path: Path
    report: Mapping[str, Any]


GuardFunction = Callable[..., Mapping[str, object]]


def _load_wolfxl_guard() -> GuardFunction:
    """Resolve only WolfXL's versioned public contract, never an internal path."""

    try:
        module = import_module("wolfxl.guard_contract")
    except ImportError as exc:
        raise GuardUnavailableError(
            "Spreadsheet Guard requires a WolfXL runtime exposing "
            f"{EXPECTED_GUARD_CONTRACT_VERSION}"
        ) from exc
    actual_version = getattr(module, "GUARD_CONTRACT_VERSION", None)
    if actual_version != EXPECTED_GUARD_CONTRACT_VERSION:
        raise GuardUnavailableError(
            "installed WolfXL Guard contract is incompatible; "
            f"expected {EXPECTED_GUARD_CONTRACT_VERSION}, got {actual_version!r}"
        )
    runner = getattr(module, "run_guard_contract", None)
    if not callable(runner):
        raise GuardUnavailableError(
            "installed WolfXL runtime does not expose run_guard_contract"
        )
    return cast(GuardFunction, runner)


def guard_workbooks(
    before: Path,
    after: Path,
    output: Path,
    *,
    policy: Mapping[str, Any] | None = None,
) -> GuardOutcome:
    """Compare two workbooks through WolfXL's pinned Guard contract.

    This package owns the stable command/package identity. WolfXL owns policy,
    inspection, report validation, and report-artifact binding.
    """

    runner = _load_wolfxl_guard()
    try:
        result = runner(before=before, after=after, output=output, policy=policy)
    except Exception as exc:
        if isinstance(exc, GuardError):
            raise
        raise GuardExecutionError(
            f"WolfXL Guard failed ({type(exc).__name__}): {exc}"
        ) from exc

    if not isinstance(result, Mapping):
        raise GuardExecutionError("WolfXL Guard contract returned a non-object result")
    if result.get("contract_version") != EXPECTED_GUARD_CONTRACT_VERSION:
        raise GuardExecutionError("WolfXL Guard result contract version drifted")
    status = result.get("status")
    if status not in {"passed", "failed", "unassessed"}:
        raise GuardExecutionError(f"WolfXL Guard returned invalid status {status!r}")
    report = result.get("report")
    if not isinstance(report, Mapping) or report.get("status") != status:
        raise GuardExecutionError("WolfXL Guard returned an inconsistent report")
    output_path_value = result.get("output_path")
    if not isinstance(output_path_value, str):
        raise GuardExecutionError("WolfXL Guard did not return its report path")
    output_path = Path(output_path_value).resolve(strict=False)
    if output_path != output.expanduser().resolve(strict=False):
        raise GuardExecutionError("WolfXL Guard returned an unbound report path")
    if not output_path.is_file():
        raise GuardExecutionError("WolfXL Guard did not write the requested report")

    return GuardOutcome(
        status=str(status),
        output_path=output_path,
        report=dict(report),
    )
