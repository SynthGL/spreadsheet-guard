# spreadsheet-guard

Read-only workbook Guard for Excel files. Compare an original workbook against a
candidate and get a machine-readable verdict on whether the candidate damaged
anything: formulas, structure, styles, defined names, or package internals.

Neither input file is ever modified. The tool fails closed: if the comparison
cannot be completed, you get an explicit error instead of a silent pass.

## Why

Agents, scripts, and humans edit workbooks. Most damage is invisible until
someone opens the file weeks later: a broken formula chain, a dropped defined
name, a corrupted pivot cache. `spreadsheet-guard` gives you a before/after
gate you can put in CI, in an agent loop, or in a review step, so a damaged
workbook never becomes the output.

## Install

```bash
pip install spreadsheet-guard
```

The CLI is a thin, stable invocation surface. The comparison itself is executed
by the WolfXL Guard runtime, which ships with the WolfXL Commercial
distribution. The MIT `wolfxl` package on PyPI (WolfXL Community) does not
include the operations SDK, so running the Guard requires a Commercial runtime.
Request evaluation access at [wolfxl.com](https://wolfxl.com).

Without the runtime installed, the CLI reports exactly that, fail-closed, as a
machine-readable error.

## Usage

```bash
spreadsheet-guard before.xlsx after.xlsx --output report.json
```

Optional policy (JSON object) to tighten or scope the comparison:

```bash
spreadsheet-guard before.xlsx after.xlsx --output report.json --policy policy.json
```

Output on stdout is a single JSON line:

```json
{"report": "report.json", "schema_version": 1, "status": "passed"}
```

The complete Guard report, including per-finding diagnostics, is written to the
`--output` path.

## Exit codes

| Code | Meaning |
|------|---------|
| 0 | Guard passed: no damage detected within the evaluated policy |
| 1 | Guard failed or could not fully assess the candidate |
| 2 | Error: invalid inputs, unreadable policy, or Guard runtime unavailable |

Errors are emitted to stderr as JSON with `status`, `error`, and `message`
fields, so wrappers can branch on them without parsing prose.

## Python API

```python
from pathlib import Path
from spreadsheet_guard import guard_workbooks

outcome = guard_workbooks(
    Path("before.xlsx"),
    Path("after.xlsx"),
    Path("report.json"),
)
print(outcome.status)  # "passed" | "failed" | "unassessed"
```

`guard_workbooks` raises `GuardUnavailableError` when no Guard runtime is
installed and `GuardExecutionError` when the runtime does not produce a valid
report.

## License

MIT. See [LICENSE](LICENSE).
