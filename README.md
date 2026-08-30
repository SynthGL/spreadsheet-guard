# spreadsheet-guard

**A read-only audit for agentic spreadsheet edits that tells you whether a candidate workbook silently damaged your file.**

Compare the original workbook with the agent's candidate. The Guard checks preservation without modifying either input.

- In a paired model study, direct openpyxl editing achieved **6/27** useful successes; a verified transaction achieved **26/27**. Exact McNemar p=1.91e-6.
- In an Enron census of **500 real workbooks**, the standard openpyxl write stack package-preserved **0 of 496** written files and silently invalidated cached formula values in **62.5%** of them.

**Falsifier: if your path is fine, the tool says so in one run.**

## Install and audit

Requires Python 3.11 or newer.

```bash
pip install spreadsheet-guard
spreadsheet-guard before.xlsx after.xlsx --output preservation-report.json
```

The command compares `before.xlsx` with `after.xlsx`, writes the complete
preservation report to `preservation-report.json`, and writes this summary to
stdout when it passes:

```json
{"report": "preservation-report.json", "schema_version": 1, "status": "passed"}
```

The package is MIT-licensed. Running the audit requires the WolfXL Commercial
runtime's operations SDK. Without that runtime, the CLI fails closed with a JSON
error and exit code 2. Request evaluation access at [wolfxl.com](https://wolfxl.com).

The optional `--policy policy.json` argument accepts a JSON object that scopes or
tightens the Guard policy:

```bash
spreadsheet-guard before.xlsx after.xlsx --output preservation-report.json --policy policy.json
```

`passed` exits 0. `failed` and `unassessed` exit 1. Invalid inputs, an unreadable
policy, or an unavailable runtime exit 2. The complete report contains the
per-finding diagnostics.

## What it checks

The Guard is a preservation check, not a generic workbook diff:

- **Package parts:** unauthorized changes to OOXML package content.
- **Cached formula values:** values a reviewer may see before any recalculation.
- **Workbook features:** survival of features such as charts, defined names,
  external links, data validation, conditional formatting, custom XML, and
  pivot-related metadata.

## What it does not check

The Guard does **not** determine whether formulas, assumptions, or business
logic are correct. A workbook can be structurally preserved and still contain
the wrong numbers. Functional and business correctness are out of scope.

## Evidence

The evidence bundle ships with the companion `spreadsheet-harness` repository.
Its configured GitHub remote is not publicly readable, so the source reports are
listed as repo-relative paths:

- `results/enron-census-v1/CENSUS-REPORT.md`: 500-workbook Enron mutation census,
  corpus definition, results, and reproduction steps.
- `results/bakeoff/write-path-matrix-v1/bakeoff-matrix.md`: write-path matrix for
  package preservation and feature survival.
- `results/customer-evidence/verified-runtime-v1/customer-story.md`: paired model
  study, transaction safety case, and runtime evidence.

The studies measure preservation under their stated corpus and task contracts.
They do not establish business-output correctness or universal spreadsheet
compatibility.

## Write with WolfXL

[WolfXL](https://github.com/wolfiesch/wolfxl) is the fail-closed verified
transaction runtime for agent writes.

## Limits

- This tool is read-only. It reports damage and does not repair or write either
  workbook.
- The evidence uses synthetic fixtures and the Enron corpus. Enron workbooks are
  from 2001 and are not a proxy for modern bank templates.
- The primary baseline is openpyxl. An Excel-COM write-path comparison runs only
  when desktop Excel is available.

## License

MIT. See [LICENSE](LICENSE).
