from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path

from spreadsheet_guard import guard_workbooks

PROOF_DIR = Path(__file__).parents[1] / "examples" / "formula-integrity"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _replace_archive_member(
    source: Path,
    destination: Path,
    member: str,
    before: bytes,
    after: bytes,
) -> None:
    replaced = False
    with (
        zipfile.ZipFile(source) as source_archive,
        zipfile.ZipFile(destination, "w") as destination_archive,
    ):
        for info in source_archive.infolist():
            data = source_archive.read(info.filename)
            if info.filename == member:
                assert data.count(before) == 1
                data = data.replace(before, after)
                replaced = True
            destination_archive.writestr(info, data)
    assert replaced


def test_bundled_engine_accepts_an_intact_workbook(tmp_path: Path) -> None:
    before = PROOF_DIR / "before.xlsx"
    after = PROOF_DIR / "after-intact.xlsx"
    before_hash = _sha256(before)
    after_hash = _sha256(after)
    output = tmp_path / "passed-report.json"

    outcome = guard_workbooks(before, after, output)

    assert outcome.status == "passed"
    assert json.loads(output.read_text(encoding="utf-8")) == outcome.report
    assert _sha256(before) == before_hash
    assert _sha256(after) == after_hash


def test_bundled_engine_rejects_a_formula_change(tmp_path: Path) -> None:
    before = PROOF_DIR / "before.xlsx"
    after = PROOF_DIR / "after-formula-damaged.xlsx"
    output = tmp_path / "failed-report.json"

    outcome = guard_workbooks(before, after, output)

    assert outcome.status == "failed"
    formula_decision = outcome.report["policy_decisions"]["formula_integrity"]
    assert formula_decision["status"] == "failed"
    assert formula_decision["findings"] == [
        {
            "kind": "formula_changed",
            "message": "worksheet formulas changed between before and after workbooks",
        }
    ]


def test_bundled_engine_rejects_a_cached_formula_result_change(
    tmp_path: Path,
) -> None:
    before = PROOF_DIR / "before.xlsx"
    after = tmp_path / "after-cached-result-damaged.xlsx"
    output = tmp_path / "failed-cached-result-report.json"
    _replace_archive_member(
        before,
        after,
        "xl/worksheets/sheet1.xml",
        b'<c r="A3"><f>SUM(A1:A2)</f><v>0</v></c>',
        b'<c r="A3"><f>SUM(A1:A2)</f><v>999</v></c>',
    )

    outcome = guard_workbooks(before, after, output)

    assert outcome.status == "failed"
    package_decision = outcome.report["policy_decisions"]["package_integrity"]
    assert package_decision["status"] == "failed"
    assert {finding["kind"] for finding in package_decision["findings"]} == {
        "worksheet_formulas_semantic_drift"
    }
