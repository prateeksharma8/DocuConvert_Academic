#!/usr/bin/env python3
"""Compare validation results for multiple DOCX documents."""
import argparse
import json
from pathlib import Path

from services.ieee_validator import validate_ieee_document
from validate_doc import format_status, load_template


def build_comparison(reports):
    checks = list(reports[0].get("checks", {})) if reports else []
    return [
        {
            "check": check,
            "task": next(
                task["task"]
                for task in reports[0].get("tasks", [])
                if task["check"] == check
            ),
            "results": {
                report["file"]: "PASS" if report["checks"].get(check) else "FAIL"
                for report in reports
            },
        }
        for check in checks
    ]


def main():
    parser = argparse.ArgumentParser(description="Compare IEEE validation results for DOCX files")
    parser.add_argument("--file", "-f", nargs="+", required=True, help="Two or more DOCX files to compare")
    parser.add_argument("--template", "-t", required=True, help="YAML validation template/context file")
    parser.add_argument("--out", "-o", help="Optional JSON comparison report path")
    parser.add_argument("--format", choices=["json", "text"], default="text")
    args = parser.parse_args()

    template_path = Path(args.template)
    if not template_path.exists():
        print(f"Template file not found: {template_path}")
        return 2
    if len(args.file) < 2:
        print("Provide at least two DOCX files to compare.")
        return 2

    template = load_template(template_path)
    reports = [validate_ieee_document(path, template.get("validation", {})) for path in args.file]
    comparison = {
        "template": str(template_path),
        "documents": reports,
        "checks": build_comparison(reports),
    }

    if args.out:
        output_path = Path(args.out)
        output_path.write_text(json.dumps(comparison, indent=2), encoding="utf-8")
        print(f"Wrote comparison report to {output_path}")

    if args.format == "json":
        print(json.dumps(comparison))
        return 0

    print(f"Document comparison ({len(reports)} files)")
    print(f"Template: {template.get('template', '(unknown)')}")
    print("\nDocuments:")
    for report in reports:
        summary = report.get("summary", {})
        print(f" - {Path(report['file']).name}: {summary.get('score', 'n/a')} ({report['overall_status']})")

    print("\nCheck comparison:")
    for check in comparison["checks"]:
        statuses = " | ".join(
            f"{Path(file).name}: {format_status(status)}"
            for file, status in check["results"].items()
        )
        print(f" - {check['task']}: {statuses}")

    print("\nFailure reasons:")
    for report in reports:
        failures = report.get("notes", [])
        if failures:
            print(f" - {Path(report['file']).name}:")
            for note in failures:
                print(f"   - {note}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
