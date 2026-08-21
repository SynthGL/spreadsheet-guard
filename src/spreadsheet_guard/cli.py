"""Command-line entry point for the read-only Spreadsheet Guard."""

from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any, cast

from spreadsheet_guard.guard import GuardError, guard_workbooks


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="spreadsheet-guard",
        description="Compare two workbooks with WolfXL Guard without modifying them.",
    )
    parser.add_argument("before", type=Path, help="Original workbook")
    parser.add_argument("after", type=Path, help="Candidate workbook")
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="Destination for the complete JSON Guard report",
    )
    parser.add_argument(
        "--policy",
        type=Path,
        help="Optional JSON Guard policy",
    )
    return parser


def _load_policy(path: Path | None) -> Mapping[str, Any] | None:
    if path is None:
        return None
    try:
        payload = json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"cannot read Guard policy: {exc}") from exc
    if not isinstance(payload, dict):
        raise ValueError("Guard policy must be a JSON object")
    return cast(dict[str, Any], payload)


def _emit(payload: Mapping[str, Any], *, error: bool = False) -> None:
    stream = sys.stderr if error else sys.stdout
    print(json.dumps(payload, sort_keys=True), file=stream)


def main(argv: Sequence[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        policy = _load_policy(args.policy)
        outcome = guard_workbooks(
            args.before,
            args.after,
            args.output,
            policy=policy,
        )
    except (GuardError, ValueError) as exc:
        _emit(
            {
                "schema_version": 1,
                "status": "error",
                "error": type(exc).__name__,
                "message": str(exc),
            },
            error=True,
        )
        return 2

    _emit(
        {
            "schema_version": 1,
            "status": outcome.status,
            "report": str(outcome.output_path),
        }
    )
    return 0 if outcome.status == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
