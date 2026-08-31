from __future__ import annotations

import json
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

import spreadsheet_guard.cli as cli_module
import spreadsheet_guard.guard as guard_module
from spreadsheet_guard.guard import (
    GuardExecutionError,
    GuardOutcome,
    GuardUnavailableError,
    guard_workbooks,
)


def test_guard_workbooks_delegates_to_versioned_contract(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    before = tmp_path / "before.xlsx"
    after = tmp_path / "after.xlsx"
    output = tmp_path / "guard.json"
    before.write_bytes(b"before")
    after.write_bytes(b"after")
    received: dict[str, Any] = {}

    def fake_guard(**kwargs: Any) -> dict[str, object]:
        received.update(kwargs)
        report = {"schema_version": 1, "status": "passed"}
        kwargs["output"].write_text(json.dumps(report))
        return {
            "contract_version": "wolfxl-guard-contract-v1",
            "status": "passed",
            "output_path": str(kwargs["output"].resolve()),
            "report": report,
        }

    monkeypatch.setattr(guard_module, "_load_wolfxl_guard", lambda: fake_guard)

    outcome = guard_workbooks(before, after, output)

    assert outcome.status == "passed"
    assert outcome.output_path == output.resolve()
    assert received == {
        "before": before,
        "after": after,
        "output": output,
        "policy": None,
    }
    assert before.read_bytes() == b"before"
    assert after.read_bytes() == b"after"


@pytest.mark.parametrize(
    ("version", "has_runner"),
    [
        (None, True),
        ("wolfxl-guard-contract-v0", True),
        ("wolfxl-guard-contract-v1", False),
    ],
)
def test_loader_fails_closed_on_contract_drift(
    version: object,
    has_runner: bool,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    module = ModuleType("wolfxl.guard_contract")
    if version is not None:
        module.GUARD_CONTRACT_VERSION = version
    if has_runner:
        module.run_guard_contract = lambda **kwargs: kwargs
    monkeypatch.setattr(guard_module, "import_module", lambda name: module)

    with pytest.raises(GuardUnavailableError):
        guard_module._load_wolfxl_guard()


@pytest.mark.parametrize("invalid_status", ["unknown", [], {}])
def test_guard_workbooks_fails_closed_on_invalid_contract_result(
    invalid_status: object,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    output = tmp_path / "guard.json"
    output.write_text("{}")

    monkeypatch.setattr(
        guard_module,
        "_load_wolfxl_guard",
        lambda: lambda **kwargs: {
            "contract_version": "wolfxl-guard-contract-v1",
            "status": invalid_status,
            "output_path": str(output),
            "report": {},
        },
    )

    with pytest.raises(GuardExecutionError, match="invalid status"):
        guard_workbooks(tmp_path / "before.xlsx", tmp_path / "after.xlsx", output)


@pytest.mark.parametrize(
    ("status", "expected_code"),
    [("passed", 0), ("failed", 1), ("unassessed", 1)],
)
def test_guard_cli_exit_code_tracks_guard_status(
    status: str,
    expected_code: int,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    output = tmp_path / "guard.json"

    def fake_guard_workbooks(
        before: Path,
        after: Path,
        requested_output: Path,
        *,
        policy: object,
    ) -> GuardOutcome:
        assert before == tmp_path / "before.xlsx"
        assert after == tmp_path / "after.xlsx"
        assert requested_output == output
        assert policy is None
        return GuardOutcome(
            status=status,
            output_path=output,
            report={"status": status},
        )

    monkeypatch.setattr(cli_module, "guard_workbooks", fake_guard_workbooks)

    code = cli_module.main(
        [
            str(tmp_path / "before.xlsx"),
            str(tmp_path / "after.xlsx"),
            "--output",
            str(output),
        ]
    )

    assert code == expected_code
    payload = json.loads(capsys.readouterr().out)
    assert payload == {"report": str(output), "schema_version": 1, "status": status}


def test_guard_cli_rejects_non_object_policy(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    policy = tmp_path / "policy.json"
    policy.write_text("[]")
    called = False

    def fake_guard_workbooks(*args: object, **kwargs: object) -> GuardOutcome:
        nonlocal called
        called = True
        raise AssertionError("Guard must not run with an invalid policy")

    monkeypatch.setattr(cli_module, "guard_workbooks", fake_guard_workbooks)

    code = cli_module.main(
        [
            str(tmp_path / "before.xlsx"),
            str(tmp_path / "after.xlsx"),
            "--output",
            str(tmp_path / "guard.json"),
            "--policy",
            str(policy),
        ]
    )

    assert code == 2
    assert not called
    payload = json.loads(capsys.readouterr().err)
    assert payload["status"] == "error"
    assert payload["error"] == "ValueError"
