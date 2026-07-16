"""
ci_check.py
-------------------------------------------------------------------------
Run by the GitHub Action (.github/workflows/hsn_check.yml) on a schedule.
No Flask, no live server - this just does one check and writes the
results into docs/data/ as plain files, which the Action then commits
back to the repo. GitHub Pages serves docs/ as a static site, so the
dashboard (docs/index.html) just reads these files directly - there's no
backend running at all once this script finishes.

Config comes from ci_config.json (committed, non-secret - just
key_col/compare_cols/sheets, no credentials). Email credentials, if you
want alerts, come from GitHub Secrets exposed as environment variables
(SMTP_HOST/SMTP_USER/SMTP_PASSWORD) - see notifier.py's load_config().

Usage:
    HSN_DATA_DIR=docs/data python ci_check.py
-------------------------------------------------------------------------
"""

import json
import sys
from pathlib import Path

import hsn_module
import notifier

CONFIG_PATH = Path(__file__).parent / "ci_config.json"


def main():
    if not CONFIG_PATH.exists():
        print(f"ERROR: {CONFIG_PATH} not found. Copy ci_config.example.json to ci_config.json.")
        sys.exit(1)

    hsn_config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))

    print(f"Running HSN/SAC check, writing output to: {hsn_module.BASE_DIR}")
    summary = hsn_module.run_hsn_check(hsn_config)

    for name, s in summary["sheets"].items():
        status = "baseline" if s["is_baseline"] else f"+{s['counts']['added']}/-{s['counts']['deleted']}/~{s['counts']['modified']}"
        print(f"  {name}: {status} ({s['total_codes']} total codes)")

    email_result = notifier.send_hsn_email_alert(summary)
    print(f"Email: {email_result}")

    zip_bytes = hsn_module.build_all_versions_zip()
    if zip_bytes is not None:
        (hsn_module.BASE_DIR / "hsn_all_versions.zip").write_bytes(zip_bytes)
        print(f"Wrote hsn_all_versions.zip ({len(zip_bytes)} bytes)")


if __name__ == "__main__":
    main()
