"""Compare two local workbooks for bounded OOXML fidelity regressions.

The optional API default is the strict five-dimension policy. An explicit
policy object may set any dimension to ``null`` (unassessed), use
``{"mode": "unchanged"}``, or provide bounded inventory counts. The CLI
requires the same object in a JSON file.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
import re
import sys
import tempfile
from collections.abc import Mapping
from pathlib import Path
from typing import Any
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

from . import _ooxml_fidelity as fidelity_audit

SCHEMA_VERSION = 1
EXIT_PASSED = 0
EXIT_REGRESSION = 1
EXIT_INVALID_INPUT = 2
EXIT_INTERNAL_ERROR = 3
_SUPPORTED_SUFFIXES = frozenset({".xlsx", ".xlsm"})
_POLICY_DIMENSIONS = (
    "macro_inventory",
    "external_link_inventory",
    "worksheet_inventory",
    "formula_integrity",
    "package_integrity",
)
_POLICY_ALIASES = {"ooxml_package_integrity": "package_integrity"}
_INVENTORY_POLICY_KEYS = frozenset({"mode", "max_count", "expected_count"})
_WORKSHEET_POLICY_KEYS = frozenset(
    {"mode", "max_count", "expected_count", "expected_names"}
)
_INTEGRITY_POLICY_KEYS = frozenset({"mode"})
_MODE_ALIASES = {
    "preserve": "unchanged",
    "strict": "unchanged",
    "unchanged": "unchanged",
}
_DEFAULT_POLICY = {dimension: {"mode": "unchanged"} for dimension in _POLICY_DIMENSIONS}

FormulaEntry = tuple[str, list[tuple[str, str]], str | None]


class GuardInputError(ValueError):
    """Raised when Guard cannot safely compare the requested files."""


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _validated_workbook(path: Path, *, label: str) -> Path:
    try:
        resolved = path.expanduser().resolve(strict=True)
    except OSError as exc:
        raise GuardInputError(f"{label} workbook is not a readable file") from exc
    if not resolved.is_file() or resolved.suffix.lower() not in _SUPPORTED_SUFFIXES:
        raise GuardInputError(f"{label} workbook must be an .xlsx or .xlsm file")
    return resolved


def _validated_output(path: Path, *, before: Path, after: Path) -> Path:
    if path.suffix.lower() != ".json":
        raise GuardInputError("output must be a .json file")
    output = path.expanduser().resolve(strict=False)
    if output in {before, after}:
        raise GuardInputError("output must not replace either workbook")
    return output


def _artifact_summary(path: Path) -> dict[str, Any]:
    return {
        "filename": path.name,
        "sha256": _sha256(path),
        "size_bytes": path.stat().st_size,
    }


def _atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8",
            dir=path.parent,
            prefix=f".{path.name}.",
            suffix=".tmp",
            delete=False,
        ) as handle:
            temporary_path = Path(handle.name)
            json.dump(payload, handle, sort_keys=True, separators=(",", ":"))
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_path, path)
        temporary_path = None
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


_POLICY_DEFAULT = object()


def _normalized_policy(policy: object) -> dict[str, dict[str, Any] | None]:
    if policy is _POLICY_DEFAULT:
        policy = _DEFAULT_POLICY
    if not isinstance(policy, Mapping):
        raise GuardInputError("policy must be a JSON object")

    raw_policy = dict(policy)
    invalid_keys = sorted(
        str(key)
        for key in raw_policy
        if not isinstance(key, str)
        or (key not in _POLICY_DIMENSIONS and key not in _POLICY_ALIASES)
    )
    if invalid_keys:
        raise GuardInputError(f"policy has unknown key(s): {', '.join(invalid_keys)}")

    normalized: dict[str, dict[str, Any] | None] = {}
    for key, value in raw_policy.items():
        canonical = _POLICY_ALIASES.get(key, key)
        if canonical in normalized:
            raise GuardInputError(f"policy specifies {canonical!r} more than once")
        if value is None:
            normalized[canonical] = None
            continue
        if not isinstance(value, Mapping):
            raise GuardInputError(f"policy.{canonical} must be null or an object")
        allowed = (
            _WORKSHEET_POLICY_KEYS
            if canonical == "worksheet_inventory"
            else _INTEGRITY_POLICY_KEYS
            if canonical in {"formula_integrity", "package_integrity"}
            else _INVENTORY_POLICY_KEYS
        )
        dimension_policy = dict(value)
        unknown = sorted(
            str(field) for field in dimension_policy if field not in allowed
        )
        if unknown:
            raise GuardInputError(
                f"policy.{canonical} has unknown key(s): {', '.join(unknown)}"
            )
        normalized_dimension: dict[str, Any] = {}
        if "mode" in dimension_policy:
            mode = dimension_policy["mode"]
            if not isinstance(mode, str) or mode not in _MODE_ALIASES:
                raise GuardInputError(
                    f"policy.{canonical}.mode must be 'unchanged' or 'preserve'"
                )
            normalized_dimension["mode"] = _MODE_ALIASES[mode]
        for field in ("max_count", "expected_count"):
            if field not in dimension_policy:
                continue
            count = dimension_policy[field]
            if isinstance(count, bool) or not isinstance(count, int) or count < 0:
                raise GuardInputError(
                    f"policy.{canonical}.{field} must be a non-negative integer"
                )
            normalized_dimension[field] = count
        if "expected_names" in dimension_policy:
            names = dimension_policy["expected_names"]
            if (
                not isinstance(names, list)
                or any(not isinstance(name, str) for name in names)
                or len(set(names)) != len(names)
            ):
                raise GuardInputError(
                    f"policy.{canonical}.expected_names must be a list of unique strings"
                )
            normalized_dimension["expected_names"] = list(names)
        if not normalized_dimension:
            normalized_dimension["mode"] = "unchanged"
        if canonical in {"formula_integrity", "package_integrity"} and set(
            normalized_dimension
        ) != {"mode"}:
            raise GuardInputError(f"policy.{canonical} only accepts the mode field")
        normalized[canonical] = normalized_dimension

    # Package fidelity is the core Guard check. It remains assessed unless a
    # caller explicitly sets package_integrity to null.
    normalized.setdefault("package_integrity", {"mode": "unchanged"})
    for dimension in _POLICY_DIMENSIONS:
        normalized.setdefault(dimension, None)
    return {dimension: normalized[dimension] for dimension in _POLICY_DIMENSIONS}


def _load_policy(path: Path) -> dict[str, dict[str, Any] | None]:
    try:
        policy_path = path.expanduser().resolve(strict=True)
    except OSError as exc:
        raise GuardInputError("policy file is not a readable file") from exc
    if not policy_path.is_file():
        raise GuardInputError("policy file is not a readable file")
    try:
        text = policy_path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as exc:
        raise GuardInputError("policy file is not valid UTF-8 JSON") from exc
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError as exc:
        raise GuardInputError("policy file contains malformed JSON") from exc
    return _normalized_policy(parsed)


def load_guard_policy(path: Path) -> dict[str, dict[str, Any] | None]:
    """Load and validate a Guard policy JSON file."""
    return _load_policy(path)


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


_EXTERNAL_UNQUOTED_SHEET = re.compile(r"[^!\s\[\]\(\)\+\-\*/,:;]+!")
_EXTERNAL_QUOTED_SHEET = re.compile(r"[^']*'!")


def _contains_external_workbook_reference(formula: str) -> bool:
    """Return whether a formula contains an external-workbook reference.

    A workbook token is a bracketed token at a formula boundary followed by a
    sheet reference and ``!``. Structured table references have an identifier
    immediately before their bracket and are excluded. Double-quoted Excel
    string literals are removed before scanning.
    """
    cleaned: list[str] = []
    in_string = False
    index = 0
    while index < len(formula):
        char = formula[index]
        if char == '"':
            if in_string and index + 1 < len(formula) and formula[index + 1] == '"':
                cleaned.extend((" ", " "))
                index += 2
                continue
            in_string = not in_string
            cleaned.append(" ")
        elif in_string:
            cleaned.append(" ")
        else:
            cleaned.append(char)
        index += 1

    text = "".join(cleaned)
    search_from = 0
    while True:
        opening = text.find("[", search_from)
        if opening < 0:
            return False
        closing = text.find("]", opening + 1)
        if closing < 0 or closing == opening + 1:
            return False
        before = opening - 1
        while before >= 0 and text[before].isspace():
            before -= 1
        if before >= 0 and (text[before].isalnum() or text[before] in "_.$"):
            search_from = closing + 1
            continue
        after = text[closing + 1 :]
        if _EXTERNAL_QUOTED_SHEET.match(after) or _EXTERNAL_UNQUOTED_SHEET.match(after):
            return True
        search_from = closing + 1


def _resolve_relationship_target(source: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


def _package_inventory(path: Path) -> dict[str, Any]:
    with ZipFile(path) as archive:
        raw_parts = archive.namelist()
        counts: dict[str, int] = {}
        for part in raw_parts:
            counts[part] = counts.get(part, 0) + 1
        duplicate_parts = sorted(part for part, count in counts.items() if count > 1)
        parts = sorted(counts)
        inventory_error = (
            "OOXML package contains duplicate ZIP members" if duplicate_parts else None
        )
        macro_parts = [
            part
            for part in parts
            if part == "xl/vbaProject.bin" or part == "xl/vbaData.xml"
        ]
        external_link_parts = sorted(
            part
            for part in parts
            if re.fullmatch(r"xl/externalLinks/externalLink\d+\.xml", part)
        )
        macro_signatures = {
            part: hashlib.sha256(archive.read(part)).hexdigest() for part in macro_parts
        }
        external_link_signatures = {
            part: hashlib.sha256(archive.read(part)).hexdigest()
            for part in external_link_parts
        }
        worksheet_parts = sorted(
            part for part in parts if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", part)
        )
        worksheet_names: list[str] | None = None
        worksheet_error: str | None = None
        workbook_root: ElementTree.Element | None = None
        if "xl/workbook.xml" in parts:
            try:
                workbook_root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
            except ElementTree.ParseError:
                worksheet_error = "workbook worksheet inventory could not be read"
        else:
            worksheet_error = "workbook worksheet inventory is unavailable"

        sheet_nodes: list[ElementTree.Element] = []
        if workbook_root is not None:
            sheet_nodes = [
                sheet
                for sheet in workbook_root.iter()
                if _local_name(sheet.tag) == "sheet"
            ]
            worksheet_names = [
                str(sheet.attrib["name"])
                for sheet in sheet_nodes
                if "name" in sheet.attrib
            ]
            if len(worksheet_names) != len(sheet_nodes):
                worksheet_error = "worksheet name mapping is incomplete"
            elif not worksheet_names or not worksheet_parts:
                worksheet_error = "worksheet inventory is empty"
            elif len(set(worksheet_names)) != len(worksheet_names):
                worksheet_error = "worksheet inventory contains duplicate names"

            rels_by_id: dict[str, str] = {}
            rels_path = "xl/_rels/workbook.xml.rels"
            if rels_path not in parts:
                worksheet_error = "workbook worksheet relationships are unavailable"
            else:
                try:
                    rels_root = ElementTree.fromstring(archive.read(rels_path))
                    for relationship in rels_root.iter():
                        if _local_name(relationship.tag) != "Relationship":
                            continue
                        rel_id = relationship.attrib.get("Id")
                        target = relationship.attrib.get("Target")
                        if not rel_id or not target or rel_id in rels_by_id:
                            worksheet_error = (
                                "workbook worksheet relationships are incomplete"
                            )
                            continue
                        rels_by_id[rel_id] = _resolve_relationship_target(
                            "xl/workbook.xml", target
                        )
                except ElementTree.ParseError:
                    worksheet_error = (
                        "workbook worksheet relationships could not be read"
                    )

            mapped_parts: set[str] = set()
            for sheet in sheet_nodes:
                relationship_id = next(
                    (
                        value
                        for key, value in sheet.attrib.items()
                        if _local_name(key) == "id"
                    ),
                    None,
                )
                target = rels_by_id.get(relationship_id or "")
                if target is None or target not in worksheet_parts:
                    worksheet_error = (
                        "worksheet relationship/sheet mapping is incomplete"
                    )
                else:
                    mapped_parts.add(target)
            if set(worksheet_parts) != mapped_parts:
                worksheet_error = "worksheet relationship/sheet mapping is incomplete"

        formula_entries: dict[str, list[FormulaEntry]] = {}
        formula_error: str | None = None
        for part in worksheet_parts:
            try:
                root = ElementTree.fromstring(archive.read(part))
            except ElementTree.ParseError:
                formula_error = "worksheet formula inventory could not be read"
                worksheet_error = worksheet_error or "worksheet XML could not be read"
                continue
            if _local_name(root.tag) != "worksheet":
                formula_error = "worksheet formula inventory could not be read"
                worksheet_error = worksheet_error or "worksheet XML could not be read"
                continue
            entries: list[FormulaEntry] = []
            for cell in root.iter():
                if _local_name(cell.tag) != "c":
                    continue
                formula = next(
                    (child for child in cell if _local_name(child.tag) == "f"),
                    None,
                )
                if formula is None:
                    continue
                attrs: list[tuple[str, str]] = [
                    (name, formula.attrib[name])
                    for name in ("t", "ref", "si", "ca", "bx")
                    if name in formula.attrib
                ]
                entries.append(
                    (
                        cell.attrib.get("r", ""),
                        attrs,
                        (formula.text or "").strip() or None,
                    )
                )
            if entries:
                formula_entries[part] = entries

        external_error: str | None = None
        for part in external_link_parts:
            try:
                ElementTree.fromstring(archive.read(part))
            except ElementTree.ParseError:
                external_error = "external-link XML inventory could not be read"
                break
        macro_error: str | None = None
        duplicate_worksheet = any(
            part in {"xl/workbook.xml", "xl/_rels/workbook.xml.rels"}
            or part in worksheet_parts
            for part in duplicate_parts
        )
        duplicate_macro = any(part in macro_parts for part in duplicate_parts)
        duplicate_external = any(
            part in external_link_parts for part in duplicate_parts
        )
        if duplicate_worksheet:
            worksheet_error = (
                worksheet_error or "worksheet evidence contains duplicate ZIP members"
            )
        if duplicate_macro:
            macro_error = "macro evidence contains duplicate ZIP members"
        if duplicate_external:
            external_error = "external-link evidence contains duplicate ZIP members"
        if worksheet_error is not None:
            formula_error = formula_error or "worksheet/formula evidence is incomplete"
        formulas: dict[str, list[FormulaEntry]] | None = formula_entries
        if formula_error is not None:
            formulas = None
        if formulas is not None:
            for part, entries in formulas.items():
                for cell_reference, _attrs, formula_text in entries:
                    if isinstance(
                        formula_text, str
                    ) and _contains_external_workbook_reference(formula_text):
                        label = f"{part}:{cell_reference}"
                        external_link_parts.append(label)
                        external_link_signatures[label] = hashlib.sha256(
                            formula_text.encode("utf-8")
                        ).hexdigest()
            external_link_parts = sorted(set(external_link_parts))

        if not worksheet_parts:
            worksheet_error = worksheet_error or "worksheet inventory is empty"
            formula_error = (
                formula_error or "worksheet formula inventory is unavailable"
            )
            formulas = None
        return {
            "parts": parts,
            "duplicate_parts": duplicate_parts,
            "inventory_error": inventory_error,
            "macro_parts": macro_parts,
            "external_link_parts": external_link_parts,
            "macro_signatures": macro_signatures,
            "external_link_signatures": external_link_signatures,
            "worksheet_parts": worksheet_parts,
            "worksheet_names": worksheet_names,
            "worksheet_error": worksheet_error,
            "formulas": formulas,
            "formula_error": formula_error,
            "macro_error": macro_error,
            "external_error": external_error,
        }


def _not_assessed(dimension: str) -> dict[str, Any]:
    return {
        "status": "unassessed",
        "findings": [
            {
                "kind": "not_assessed",
                "message": f"{dimension} is not assessed by policy",
            }
        ],
    }


def _inventory_decision(
    dimension: str,
    config: dict[str, Any] | None,
    before_values: list[str] | None,
    after_values: list[str] | None,
    *,
    before_error: str | None = None,
    after_error: str | None = None,
    before_signatures: Mapping[str, str] | None = None,
    after_signatures: Mapping[str, str] | None = None,
) -> dict[str, Any]:
    if config is None:
        return _not_assessed(dimension)
    if (
        before_values is None
        or after_values is None
        or before_error is not None
        or after_error is not None
    ):
        return {
            "status": "unassessed",
            "findings": [
                {
                    "kind": "evidence_unavailable",
                    "message": before_error
                    or after_error
                    or "inventory evidence is unavailable",
                }
            ],
        }
    before_count = len(before_values)
    after_count = len(after_values)
    evidence = {
        "before_count": before_count,
        "after_count": after_count,
        "before_items": list(before_values),
        "after_items": list(after_values),
    }
    findings: list[dict[str, Any]] = []
    if "max_count" in config and after_count > config["max_count"]:
        findings.append(
            {
                "kind": "max_count_exceeded",
                "message": (
                    f"{dimension} after count {after_count} exceeds max_count {config['max_count']}"
                ),
                "max_count": config["max_count"],
                "after_count": after_count,
            }
        )
    if "expected_count" in config and after_count != config["expected_count"]:
        findings.append(
            {
                "kind": "unexpected_count",
                "message": (
                    f"{dimension} after count {after_count} does not equal "
                    f"expected_count {config['expected_count']}"
                ),
                "expected_count": config["expected_count"],
                "after_count": after_count,
            }
        )
    if "expected_names" in config and after_values != config["expected_names"]:
        findings.append(
            {
                "kind": "unexpected_items",
                "message": f"{dimension} after inventory differs from expected items",
                "expected_items": list(config["expected_names"]),
                "after_items": list(after_values),
            }
        )
    signatures_changed = (
        before_signatures is not None
        and after_signatures is not None
        and dict(before_signatures) != dict(after_signatures)
    )
    if config.get("mode") == "unchanged" and (
        before_values != after_values or signatures_changed
    ):
        findings.append(
            {
                "kind": "inventory_changed",
                "message": f"{dimension} changed between before and after workbooks",
                "before_items": list(before_values),
                "after_items": list(after_values),
            }
        )
    return {
        "status": "failed" if findings else "passed",
        "evidence": evidence,
        "findings": findings,
    }


def _formula_decision(
    config: dict[str, Any] | None, before: dict[str, Any], after: dict[str, Any]
) -> dict[str, Any]:
    if config is None:
        return _not_assessed("formula_integrity")
    before_formulas = before["formulas"]
    after_formulas = after["formulas"]
    if before_formulas is None or after_formulas is None:
        return {
            "status": "unassessed",
            "findings": [
                {
                    "kind": "evidence_unavailable",
                    "message": before["formula_error"]
                    or after["formula_error"]
                    or "formula evidence is unavailable",
                }
            ],
        }
    evidence = {
        "before_formula_count": sum(
            len(entries) for entries in before_formulas.values()
        ),
        "after_formula_count": sum(len(entries) for entries in after_formulas.values()),
    }
    if before_formulas != after_formulas:
        return {
            "status": "failed",
            "evidence": evidence,
            "findings": [
                {
                    "kind": "formula_changed",
                    "message": "worksheet formulas changed between before and after workbooks",
                }
            ],
        }
    return {"status": "passed", "evidence": evidence, "findings": []}


def _package_decision(
    config: dict[str, Any] | None,
    audit: dict[str, Any],
    before_inventory: dict[str, Any],
    after_inventory: dict[str, Any],
) -> dict[str, Any]:
    if config is None:
        return _not_assessed("package_integrity")
    issues = [
        {key: issue[key] for key in ("kind", "part", "message") if key in issue}
        for issue in audit.get("issues", [])
    ]
    seen_messages = {str(issue.get("message", "")) for issue in issues}
    for label, inventory in (
        ("before", before_inventory),
        ("after", after_inventory),
    ):
        for error_key in ("inventory_error", "worksheet_error", "formula_error"):
            message = inventory.get(error_key)
            if not message or message in seen_messages:
                continue
            seen_messages.add(message)
            issues.append(
                {
                    "kind": "package_evidence_incomplete",
                    "part": "package",
                    "message": f"{label} workbook: {message}",
                }
            )
    issues.sort(
        key=lambda issue: (
            str(issue.get("kind", "")),
            str(issue.get("part", "")),
            str(issue.get("message", "")),
        )
    )
    return {
        "status": "failed" if issues else "passed",
        "evidence": {"issue_count": len(issues)},
        "findings": issues,
    }


def run_guard(
    *,
    before: Path,
    after: Path,
    output: Path,
    policy: Mapping[str, Any] | None | object = _POLICY_DEFAULT,
) -> dict[str, Any]:
    """Run the fidelity audit and evaluate one explicit bounded policy."""
    before_path = _validated_workbook(before, label="before")
    after_path = _validated_workbook(after, label="after")
    if before_path == after_path:
        raise GuardInputError("before and after must be different files")
    output_path = _validated_output(output, before=before_path, after=after_path)
    normalized_policy = _normalized_policy(policy)

    audit = fidelity_audit.audit(
        before_path,
        after_path,
        compact_semantic_drift=True,
    )
    audit["before"]["path"] = before_path.name
    audit["after"]["path"] = after_path.name
    before_inventory = _package_inventory(before_path)
    after_inventory = _package_inventory(after_path)

    policy_decisions = {
        "macro_inventory": _inventory_decision(
            "macro_inventory",
            normalized_policy["macro_inventory"],
            before_inventory["macro_parts"],
            after_inventory["macro_parts"],
            before_error=before_inventory["macro_error"],
            after_error=after_inventory["macro_error"],
            before_signatures=before_inventory["macro_signatures"],
            after_signatures=after_inventory["macro_signatures"],
        ),
        "external_link_inventory": _inventory_decision(
            "external_link_inventory",
            normalized_policy["external_link_inventory"],
            before_inventory["external_link_parts"],
            after_inventory["external_link_parts"],
            before_error=(
                before_inventory["external_error"]
                or before_inventory["worksheet_error"]
                or before_inventory["formula_error"]
            ),
            after_error=(
                after_inventory["external_error"]
                or after_inventory["worksheet_error"]
                or after_inventory["formula_error"]
            ),
            before_signatures=before_inventory["external_link_signatures"],
            after_signatures=after_inventory["external_link_signatures"],
        ),
        "worksheet_inventory": _inventory_decision(
            "worksheet_inventory",
            normalized_policy["worksheet_inventory"],
            before_inventory["worksheet_names"],
            after_inventory["worksheet_names"],
            before_error=before_inventory["worksheet_error"],
            after_error=after_inventory["worksheet_error"],
        ),
        "formula_integrity": _formula_decision(
            normalized_policy["formula_integrity"], before_inventory, after_inventory
        ),
        "package_integrity": _package_decision(
            normalized_policy["package_integrity"],
            audit,
            before_inventory,
            after_inventory,
        ),
    }
    statuses = [decision["status"] for decision in policy_decisions.values()]
    if "failed" in statuses:
        status = "failed"
    elif "unassessed" in statuses:
        status = "unassessed"
    else:
        status = "passed"
    report = {
        "schema_version": SCHEMA_VERSION,
        "status": status,
        "before": _artifact_summary(before_path),
        "after": _artifact_summary(after_path),
        "policy": normalized_policy,
        "policy_decisions": policy_decisions,
        "fidelity_audit": audit,
        "not_assessed": [
            "intended change authorization",
            "calculation correctness",
            "rendered appearance",
            "macro execution",
        ],
    }
    _atomic_write_json(output_path, report)
    return report


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--before", type=Path, required=True)
    parser.add_argument("--after", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--policy",
        type=Path,
        help="JSON policy object describing bounded Guard checks",
    )
    return parser.parse_args(argv)


def _error_document(
    category: str, message: str, *, exception: str | None = None
) -> str:
    error: dict[str, Any] = {"category": category, "message": message}
    if exception is not None:
        error["exception"] = exception
    return json.dumps(
        {"schema_version": SCHEMA_VERSION, "status": "error", "error": error},
        sort_keys=True,
        separators=(",", ":"),
    )


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        if args.policy is None:
            raise GuardInputError("policy input is required")
        policy = load_guard_policy(args.policy)
        report = run_guard(
            before=args.before,
            after=args.after,
            output=args.output,
            policy=policy,
        )
    except (GuardInputError, BadZipFile) as exc:
        print(_error_document("invalid_input", str(exc)), file=sys.stderr)
        return EXIT_INVALID_INPUT
    except OSError as exc:
        print(
            _error_document(
                "io_error",
                "workbook comparison could not be read or written",
                exception=type(exc).__name__,
            ),
            file=sys.stderr,
        )
        return EXIT_INVALID_INPUT
    except Exception as exc:
        print(
            _error_document(
                "internal_error",
                "workbook comparison failed unexpectedly",
                exception=type(exc).__name__,
            ),
            file=sys.stderr,
        )
        return EXIT_INTERNAL_ERROR

    summary = {
        "status": report["status"],
        "issue_count": report["fidelity_audit"]["issue_count"],
        "output": args.output.name,
    }
    print(json.dumps(summary, sort_keys=True, separators=(",", ":")))
    return EXIT_PASSED if report["status"] == "passed" else EXIT_REGRESSION


if __name__ == "__main__":
    raise SystemExit(main())
