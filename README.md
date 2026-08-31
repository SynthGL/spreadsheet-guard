# spreadsheet-guard

**A standalone, read-only preservation gate for `.xlsx` and `.xlsm` files.**

Compare an original workbook with a candidate produced by an agent, script, or
library. `spreadsheet-guard` audits OOXML structure, formulas, cached values, and
feature relationships without modifying either input. It emits deterministic
JSON and automation-friendly exit codes.

- MIT licensed
- No proprietary runtime or service dependency
- Local files only
- Fail-closed on invalid or incomplete evidence

## Install and audit

Requires Python 3.11 or newer.

```text
pip install spreadsheet-guard
spreadsheet-guard before.xlsx after.xlsx --output preservation-report.json
```

The command writes the complete report to `preservation-report.json` and a
one-line summary to stdout:

```json
{"report": "preservation-report.json", "schema_version": 1, "status": "passed"}
```

| Status | Exit code | Meaning |
|---|---:|---|
| `passed` | 0 | Every configured preservation dimension passed. |
| `failed` | 1 | At least one configured dimension found a regression. |
| `unassessed` | 1 | Required evidence was unavailable or a dimension was disabled. |
| `error` | 2 | Inputs, policy, or execution were invalid. |

## Run the checked-in proof

From a clone of this repository, these commands target bash or zsh and use
repo-relative paths:

```bash
uv sync

uv run spreadsheet-guard \
  examples/formula-integrity/before.xlsx \
  examples/formula-integrity/after-intact.xlsx \
  --output /tmp/spreadsheet-guard-passed.json

uv run spreadsheet-guard \
  examples/formula-integrity/before.xlsx \
  examples/formula-integrity/after-formula-damaged.xlsx \
  --output /tmp/spreadsheet-guard-failed.json
```

The first command exits 0 with `passed`. The second exits 1 with `failed`
because `SUM(A1:A2)` changed to `SUM(A1:A1)`. The frozen reports are checked in
at:

- [`examples/formula-integrity/passed-report.json`](examples/formula-integrity/passed-report.json)
- [`examples/formula-integrity/failed-report.json`](examples/formula-integrity/failed-report.json)

This compact proof demonstrates the packaged engine and report contract. It does
not establish universal Excel compatibility or business-output correctness.

## What it checks

The default policy requires all five dimensions to remain unchanged:

1. Macro inventory
2. External-link inventory
3. Worksheet inventory and names
4. Formula integrity
5. OOXML package integrity

Package integrity checks part relationships, XML parseability, cached formula
values, and semantic fingerprints for workbook features such as charts, defined
names, data validation, conditional formatting, drawings, external links, and
pivot-related metadata.

The strict default also reports intentional formula or structural changes. The
caller must compare each finding with the intended edit.

## Custom policy

Pass `--policy policy.json` to replace the strict default:

```text
spreadsheet-guard before.xlsx after.xlsx \
  --output preservation-report.json \
  --policy policy.json
```

```json
{
  "macro_inventory": {"mode": "unchanged"},
  "external_link_inventory": {"mode": "unchanged"},
  "worksheet_inventory": {
    "mode": "unchanged",
    "expected_names": ["Inputs", "Model", "Outputs"]
  },
  "formula_integrity": {"mode": "unchanged"},
  "package_integrity": {"mode": "unchanged"}
}
```

A custom policy should enumerate all five dimensions when a `passed` result is
required. A dimension set to `null` is reported as `unassessed`.

## Python API

```python
from pathlib import Path

from spreadsheet_guard import guard_workbooks

outcome = guard_workbooks(
    Path("before.xlsx"),
    Path("after.xlsx"),
    Path("preservation-report.json"),
)
print(outcome.status)
```

## Limits

- The Guard evaluates preservation. It does not determine whether formulas,
  assumptions, or business logic are correct.
- The Guard is read-only. It reports findings and does not repair either
  workbook.
- Inputs are limited to OOXML `.xlsx` and `.xlsm` workbooks.
- A clean report applies only to the configured dimensions and the supplied
  before-and-after pair.

## Write with a commit gate

[`WolfXL`](https://github.com/wolfiesch/wolfxl) applies policy before a workbook
write, runs independent verification, and commits only accepted candidates.
`spreadsheet-guard` is the free, read-only audit surface.

## License

MIT. See [LICENSE](LICENSE).
