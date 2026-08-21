"""Read-only workbook Guard: compare two Excel workbooks without modifying them."""

from spreadsheet_guard.guard import (
    GuardError,
    GuardExecutionError,
    GuardFunction,
    GuardOutcome,
    GuardUnavailableError,
    guard_workbooks,
)

__all__ = [
    "GuardError",
    "GuardExecutionError",
    "GuardFunction",
    "GuardOutcome",
    "GuardUnavailableError",
    "guard_workbooks",
]
