#!/usr/bin/env python3
"""Simple CLI to validate a DOCX file and emit a report.

Usage:
  python validate_doc.py --file paper.docx [--template context.yaml] [--out report.json]
"""
import argparse
import json
import sys
from pathlib import Path


def format_status(status: str) -> str:
    if status == "PASS":
        value = "✅ PASS"
        color = "\033[32m"
    else:
        value = "❌ FAIL"
        color = "\033[31m"
    reset = "\033[0m" if sys.stdout.isatty() else ""
    return f"{color}{value}{reset}"

def load_template(path: Path):
    try:
        import yaml
    except Exception:
        raise RuntimeError("PyYAML is required to load template YAML files. Please install via 'pip install pyyaml'.")
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def main():
    parser = argparse.ArgumentParser(description="Validate a DOCX file (IEEE-like) and produce a report")
    parser.add_argument("--file", "-f", required=True, help="Path to the .docx file to validate")
    parser.add_argument("--template", "-t", help="Optional template/context YAML to guide validation")
    parser.add_argument("--out", "-o", help="Optional path to write JSON report")
    parser.add_argument("--format", choices=["json","text"], default="text", help="Output format for stdout")

    args = parser.parse_args()
    doc_path = Path(args.file)
    template = None
    if args.template:
        template_path = Path(args.template)
        if not template_path.exists():
            print(f"Template file not found: {template_path}")
            return 2
        template = load_template(template_path)

    # Import validator
    from services.ieee_validator import validate_ieee_document

    report = validate_ieee_document(doc_path, (template or {}).get("validation", {}))
    # Attach template info if provided
    if template is not None:
        report["template"] = template

    # Write JSON output if requested
    if args.out:
        outp = Path(args.out)
        with open(outp, "w", encoding="utf-8") as f:
            json.dump(report, f, indent=2)
        print(f"Wrote report to {outp}")

    # Print human-friendly summary
    if args.format == "json":
        print(json.dumps(report))
    else:
        print("Validation report for:", doc_path)
        print("Overall status:", report.get("overall_status"))
        summary = report.get("summary", {})
        if summary:
            print(f"Score: {summary['score']} ({summary['passed']} passed, {summary['failed']} failed)")
        print("Tasks:")
        for task in report.get("tasks", []):
                print(f" - {task['task']}: {format_status(task['status'])}")
        notes = report.get("notes", [])
        if notes:
            print("Notes:")
            for n in notes:
                print("  -", n)
        if template is not None:
            print("Template:", template.get("template", "(unknown)"))

if __name__ == "__main__":
    raise SystemExit(main())
