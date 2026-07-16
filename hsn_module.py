"""
hsn_module.py
-------------------------------------------------------------------------
Server-side glue between hsn_core.py (pure comparison/validation logic)
and the Flask app.

CRITICAL: the portal's HSN_SAC.xlsx may contain multiple sheets (e.g.
one for HSN, one for SAC). These are DIFFERENT code systems and are
NEVER merged or cross-compared - every sheet gets its own independent
snapshot, its own independent comparison against ITS OWN previous
version, its own independent history, and its own independent downloads.
An "HSN" row is only ever compared to a previous "HSN" row; a "SAC" row
only ever to a previous "SAC" row.

No reference file is required. For each sheet, every run:

  1. Saves a permanent, timestamped snapshot of that sheet alone.
  2. Compares that sheet's snapshot against THAT SAME SHEET's previous
     snapshot (if one exists) - the primary comparison. First run for a
     given sheet has nothing to compare against yet, so it's a baseline.
  3. OPTIONALLY, if hsn_config["reference_file"] is set, ALSO compares
     against a same-named sheet in that file, as a secondary comparison.

Every run is archived under hsn_versions/<run_id>/, with one snapshot
file per sheet - nothing is ever overwritten.

Per-sheet files written on every run (● = sheet name, sanitized for filenames):
  hsn_current_version__●.xlsx    - always available, even on a baseline run
  hsn_delta__●.xlsx / .xml        - only once there's a previous version for ● to diff against
  hsn_validation_issues__●.csv     - only if issues found for ●
Shared files:
  hsn_delta.json                    - latest-run summary, nested per sheet
  hsn_history.json                   - append-only index of every run, nested per sheet
-------------------------------------------------------------------------
"""

import json
import re
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Optional

import hsn_core as core

BASE_DIR = Path(__file__).parent
DELTA_JSON = BASE_DIR / "hsn_delta.json"
HISTORY_FILE = BASE_DIR / "hsn_history.json"
VERSIONS_DIR = BASE_DIR / "hsn_versions"


def _safe_name(sheet_name: str) -> str:
    """Turns a sheet name into something safe to use in a filename."""
    return re.sub(r"\W+", "_", sheet_name).strip("_") or "sheet"


def _current_version_path(sheet_name: str) -> Path:
    return BASE_DIR / f"hsn_current_version__{_safe_name(sheet_name)}.xlsx"


def _delta_xlsx_path(sheet_name: str) -> Path:
    return BASE_DIR / f"hsn_delta__{_safe_name(sheet_name)}.xlsx"


def _delta_xml_path(sheet_name: str) -> Path:
    return BASE_DIR / f"hsn_delta__{_safe_name(sheet_name)}.xml"


def _issues_csv_path(sheet_name: str) -> Path:
    return BASE_DIR / f"hsn_validation_issues__{_safe_name(sheet_name)}.csv"


def _issues_xlsx_path(sheet_name: str) -> Path:
    return BASE_DIR / f"hsn_validation_issues__{_safe_name(sheet_name)}.xlsx"


def _full_validation_csv_path(sheet_name: str) -> Path:
    return BASE_DIR / f"hsn_full_validation__{_safe_name(sheet_name)}.csv"


def _full_validation_xlsx_path(sheet_name: str) -> Path:
    return BASE_DIR / f"hsn_full_validation__{_safe_name(sheet_name)}.xlsx"


def _df_to_records(df, cols):
    """Small helper: DataFrame -> list of plain dicts, JSON-safe."""
    return df[cols].where(df[cols].notna(), "").to_dict(orient="records")


def _load_history() -> list:
    if not HISTORY_FILE.exists():
        return []
    try:
        return json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []


def _save_history(history: list):
    HISTORY_FILE.write_text(json.dumps(history, indent=2, ensure_ascii=False), encoding="utf-8")


def _most_recent_previous_run():
    """Returns (run_id, run_dir) for the last archived run, or (None, None)
    if this is the first run ever."""
    history = _load_history()
    if not history:
        return None, None
    last = sorted(history, key=lambda h: h["run_id"])[-1]
    run_dir = VERSIONS_DIR / last["run_id"]
    if run_dir.exists():
        return last["run_id"], run_dir
    return None, None


def _validate_columns(df, key_col, compare_cols, label):
    """Raises a clear, actionable error instead of a cryptic KeyError when
    key_col/compare_cols in config.json don't match the sheet's actual
    column headers."""
    actual_columns = list(df.columns)
    missing = [c for c in [key_col] + list(compare_cols) if c not in actual_columns]
    if missing:
        raise ValueError(
            f"config.json's hsn.key_col/compare_cols don't match the actual columns "
            f"in the {label}. Missing: {missing}. "
            f"Actual columns found: {actual_columns}. "
            f"Update config.json so key_col/compare_cols exactly match one of the "
            f"columns listed above (spelling, spacing, and case all matter), then restart."
        )


def _run_comparison(old_df, new_df, key_col, compare_cols, auto_clean_xml):
    """Runs one comparison (old vs new) within a SINGLE sheet's data only.
    Returns everything needed to display it, export it, and validate it."""
    _validate_columns(old_df, key_col, compare_cols, "previous version of this sheet")
    _validate_columns(new_df, key_col, compare_cols, "current version of this sheet")
    added, deleted, modified = core.compare_dataframes(old_df, new_df, key_col, compare_cols)

    key_new_col = f"{key_col}_src" if f"{key_col}_src" in added.columns else key_col
    key_old_col = f"{key_col}_ref" if f"{key_col}_ref" in deleted.columns else key_col

    added_cols = [c for c in added.columns if c.endswith("_src") or c == key_new_col]
    deleted_cols = [c for c in deleted.columns if c.endswith("_ref") or c == key_old_col]
    modified_cols = [c for c in modified.columns if c.endswith("_ref") or c.endswith("_src") or c == "_changed_fields"]

    excel_buf = core.export_highlighted_excel(
        added, deleted, modified,
        {"added": added_cols, "deleted": deleted_cols, "modified": modified_cols},
    )
    delta_for_xml = core.build_delta_for_xml(added, modified, key_new_col)
    issues_df = core.detect_invisible_issues(delta_for_xml, delta_for_xml.columns, "Added/Modified delta")
    xml_bytes = core.build_xml(delta_for_xml, delta_for_xml.columns, auto_clean=auto_clean_xml)

    return {
        "counts": {"added": len(added), "deleted": len(deleted), "modified": len(modified)},
        "added_records": _df_to_records(added, added_cols),
        "deleted_records": _df_to_records(deleted, deleted_cols),
        "modified_records": _df_to_records(modified, modified_cols),
        "excel_bytes": excel_buf.getvalue(),
        "xml_bytes": xml_bytes,
        "issues_df": issues_df,
        "delta_for_xml": delta_for_xml,
    }


def _resolve_sheet_columns(hsn_config: dict, sheet_name: str):
    """Figures out which key_col/compare_cols apply to a given sheet.
    Different sheets can use entirely different column names (e.g. the
    portal's HSN sheet uses HSN_CD/HSN_Description, while its SAC sheet
    uses SAC_CD/SAC_Description) - config.json can specify a per-sheet
    override under hsn.sheets.<sheet_name>, and falls back to the
    top-level hsn.key_col/hsn.compare_cols if no override exists for
    that particular sheet."""
    per_sheet = (hsn_config.get("sheets") or {}).get(sheet_name)
    if per_sheet and "key_col" in per_sheet and "compare_cols" in per_sheet:
        return per_sheet["key_col"], per_sheet["compare_cols"]

    if "key_col" in hsn_config and "compare_cols" in hsn_config:
        return hsn_config["key_col"], hsn_config["compare_cols"]

    raise ValueError(
        f"No key_col/compare_cols configured for sheet '{sheet_name}', and no "
        f"top-level fallback is set either. Add an entry under config.json's "
        f"hsn.sheets.\"{sheet_name}\" with its own key_col/compare_cols "
        f"(different sheets can use different column names)."
    )


def run_hsn_check(hsn_config: dict) -> dict:
    """hsn_config expects EITHER top-level key_col/compare_cols (applied to
    every sheet), OR a 'sheets' dict giving each sheet its own key_col/
    compare_cols - needed when sheets use different column names (e.g.
    HSN_CD vs SAC_CD). auto_clean_xml and reference_file are optional.
    Every sheet in the downloaded workbook is processed and compared
    completely independently of every other sheet."""
    auto_clean_xml = hsn_config.get("auto_clean_xml", False)
    reference_file = hsn_config.get("reference_file")

    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = VERSIONS_DIR / run_id
    run_dir.mkdir(parents=True, exist_ok=True)

    prev_run_id, prev_run_dir = _most_recent_previous_run()

    src_bytes = core.fetch_source_excel()
    sheets = core.load_excel_sheets(src_bytes)  # {sheet_name: DataFrame}, kept separate

    # Optional reference file - also loaded per-sheet if provided, matched
    # by sheet name. A reference file with only one sheet, or without a
    # matching name, simply won't have a reference comparison for sheets
    # that don't match - never an error, just skipped for that sheet.
    reference_sheets = {}
    if reference_file and Path(reference_file).exists():
        try:
            reference_sheets = core.load_excel_sheets(reference_file)
        except Exception:
            reference_sheets = {}

    sheet_summaries = {}
    for sheet_name, src_df in sheets.items():
        safe = _safe_name(sheet_name)
        total_codes = len(src_df)

        # Each sheet gets its OWN key_col/compare_cols - they are not
        # assumed to share the same column names.
        key_col, compare_cols = _resolve_sheet_columns(hsn_config, sheet_name)

        # Catch a bad key_col/compare_cols config immediately, even on a
        # baseline run, instead of waiting for a comparison to fail.
        _validate_columns(src_df, key_col, compare_cols, f"'{sheet_name}' sheet (downloaded source)")

        # Archive this sheet's snapshot, permanently, on its own.
        snap_buf = BytesIO()
        src_df.to_excel(snap_buf, index=False)
        snap_bytes = snap_buf.getvalue()
        (run_dir / f"{safe}_snapshot.xlsx").write_bytes(snap_bytes)

        # Always keep the current version of THIS sheet downloadable on
        # its own, regardless of whether there's a previous version yet.
        _current_version_path(sheet_name).write_bytes(snap_bytes)

        # FULL-DATASET description validation - scans every row currently
        # in this sheet (not just what changed) for invisible/special
        # characters in the description-type fields (compare_cols), since
        # codes themselves (key_col) aren't free text and don't need this
        # check. This runs every time, independent of whether there's a
        # previous version to diff against, so it's a standalone data-
        # quality report, not tied to the delta.
        full_issues_df = core.detect_invisible_issues(src_df, compare_cols, sheet_name)
        if not full_issues_df.empty:
            full_issues_text = full_issues_df.to_csv(index=False)
            _full_validation_csv_path(sheet_name).write_text(full_issues_text, encoding="utf-8")
            (run_dir / f"{safe}_full_validation.csv").write_text(full_issues_text, encoding="utf-8")
        elif _full_validation_csv_path(sheet_name).exists():
            _full_validation_csv_path(sheet_name).unlink()
        full_validation_issue_count = int(len(full_issues_df))

        # Same scan, but as a filtered + cell-highlighted Excel workbook -
        # only the problem rows, with the exact offending cell colored red,
        # so the issue is visible in context rather than just listed as text.
        full_issues_xlsx_bytes = core.export_validation_highlighted_excel(src_df, key_col, compare_cols)
        if full_validation_issue_count > 0:
            _full_validation_xlsx_path(sheet_name).write_bytes(full_issues_xlsx_bytes)
            (run_dir / f"{safe}_full_validation.xlsx").write_bytes(full_issues_xlsx_bytes)
        elif _full_validation_xlsx_path(sheet_name).exists():
            _full_validation_xlsx_path(sheet_name).unlink()

        counts_vs_previous = None
        added_records = deleted_records = modified_records = []
        validation_issue_count = 0
        is_baseline_for_sheet = True

        prev_snapshot_path = (prev_run_dir / f"{safe}_snapshot.xlsx") if prev_run_dir else None
        if prev_snapshot_path is not None and prev_snapshot_path.exists():
            is_baseline_for_sheet = False
            prev_df = core.load_excel(prev_snapshot_path)
            result = _run_comparison(prev_df, src_df, key_col, compare_cols, auto_clean_xml)
            counts_vs_previous = result["counts"]
            added_records = result["added_records"]
            deleted_records = result["deleted_records"]
            modified_records = result["modified_records"]

            _delta_xlsx_path(sheet_name).write_bytes(result["excel_bytes"])
            (run_dir / f"{safe}_delta_vs_previous.xlsx").write_bytes(result["excel_bytes"])
            _delta_xml_path(sheet_name).write_bytes(result["xml_bytes"])
            (run_dir / f"{safe}_delta_vs_previous.xml").write_bytes(result["xml_bytes"])

            issues_df = result["issues_df"]
            if not issues_df.empty:
                issues_text = issues_df.to_csv(index=False)
                _issues_csv_path(sheet_name).write_text(issues_text, encoding="utf-8")
                (run_dir / f"{safe}_validation_issues.csv").write_text(issues_text, encoding="utf-8")

                # Same delta-only issues, but as a filtered + highlighted
                # Excel workbook (only Added/Modified rows that changed
                # THIS check, with the offending cell colored).
                delta_for_xml = result["delta_for_xml"]
                key_new_col = f"{key_col}_src" if f"{key_col}_src" in delta_for_xml.columns else key_col
                issues_xlsx_bytes = core.export_validation_highlighted_excel(delta_for_xml, key_new_col, compare_cols)
                _issues_xlsx_path(sheet_name).write_bytes(issues_xlsx_bytes)
                (run_dir / f"{safe}_validation_issues.xlsx").write_bytes(issues_xlsx_bytes)
            else:
                if _issues_csv_path(sheet_name).exists():
                    _issues_csv_path(sheet_name).unlink()
                if _issues_xlsx_path(sheet_name).exists():
                    _issues_xlsx_path(sheet_name).unlink()
            validation_issue_count = int(len(issues_df))
        else:
            # Baseline for this sheet (either the very first run ever, or
            # a sheet that just appeared for the first time this run) -
            # no previous version of THIS sheet to diff against yet.
            for p in (_delta_xlsx_path(sheet_name), _delta_xml_path(sheet_name), _issues_csv_path(sheet_name), _issues_xlsx_path(sheet_name)):
                if p.exists():
                    p.unlink()

        # Optional secondary comparison vs a same-named sheet in the
        # reference file, if one was configured and matches.
        counts_vs_reference = None
        if sheet_name in reference_sheets:
            try:
                ref_df = reference_sheets[sheet_name]
                ref_result = _run_comparison(ref_df, src_df, key_col, compare_cols, auto_clean_xml)
                counts_vs_reference = ref_result["counts"]
                (run_dir / f"{safe}_delta_vs_reference.xlsx").write_bytes(ref_result["excel_bytes"])
                (run_dir / f"{safe}_delta_vs_reference.xml").write_bytes(ref_result["xml_bytes"])
            except ValueError:
                counts_vs_reference = None  # reference sheet has incompatible columns - skip silently

        sheet_summaries[sheet_name] = {
            "sheet_name": sheet_name,
            "key_col": key_col,
            "compare_cols": compare_cols,
            "total_codes": total_codes,
            "is_baseline": is_baseline_for_sheet,
            "counts": counts_vs_previous or {"added": 0, "deleted": 0, "modified": 0},
            "counts_vs_reference": counts_vs_reference,
            "added": added_records,
            "deleted": deleted_records,
            "modified": modified_records,
            "validation_issue_count": validation_issue_count,
            "full_validation_issue_count": full_validation_issue_count,
            "has_files": {
                "current": _current_version_path(sheet_name).exists(),
                "excel": _delta_xlsx_path(sheet_name).exists(),
                "xml": _delta_xml_path(sheet_name).exists(),
                "issues_csv": _issues_csv_path(sheet_name).exists(),
                "issues_xlsx": _issues_xlsx_path(sheet_name).exists(),
                "full_issues_csv": _full_validation_csv_path(sheet_name).exists(),
                "full_issues_xlsx": _full_validation_xlsx_path(sheet_name).exists(),
            },
        }

    run_summary = {
        "run_id": run_id,
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "compared_to_run_id": prev_run_id,
        "sheet_names": list(sheets.keys()),
        "sheets": {
            name: {
                "key_col": s["key_col"],
                "compare_cols": s["compare_cols"],
                "total_codes": s["total_codes"],
                "is_baseline": s["is_baseline"],
                "counts": s["counts"],
                "counts_vs_reference": s["counts_vs_reference"],
                "validation_issue_count": s["validation_issue_count"],
                "full_validation_issue_count": s["full_validation_issue_count"],
            }
            for name, s in sheet_summaries.items()
        },
    }
    (run_dir / "summary.json").write_text(json.dumps(run_summary, indent=2, ensure_ascii=False), encoding="utf-8")

    history = _load_history()
    history.append(run_summary)
    _save_history(history)

    full_summary = {
        "run_id": run_id,
        "ran_at": run_summary["timestamp"],
        "compared_to_run_id": prev_run_id,
        "sheet_names": list(sheets.keys()),
        "sheets": sheet_summaries,
    }

    DELTA_JSON.write_text(json.dumps(full_summary, indent=2, ensure_ascii=False), encoding="utf-8")
    return full_summary


def load_last_hsn_result() -> dict:
    if not DELTA_JSON.exists():
        return {"ran_at": None, "sheet_names": [], "sheets": {}}
    return json.loads(DELTA_JSON.read_text(encoding="utf-8"))


def load_history() -> list:
    """Returns all archived runs, most recent first. Each entry has a
    per-sheet breakdown under 'sheets'."""
    history = _load_history()
    return sorted(history, key=lambda h: h["run_id"], reverse=True)


def get_run_file(run_id: str, sheet_name: str, filetype: str) -> Optional[Path]:
    """Used by the download API to serve a specific archived run's file
    for a specific sheet."""
    allowed = {
        "source": "{safe}_snapshot.xlsx",
        "excel": "{safe}_delta_vs_previous.xlsx",
        "xml": "{safe}_delta_vs_previous.xml",
        "issues": "{safe}_validation_issues.csv",
        "issues_xlsx": "{safe}_validation_issues.xlsx",
        "full_issues": "{safe}_full_validation.csv",
        "full_issues_xlsx": "{safe}_full_validation.xlsx",
        "excel_vs_reference": "{safe}_delta_vs_reference.xlsx",
        "xml_vs_reference": "{safe}_delta_vs_reference.xml",
    }
    if filetype not in allowed:
        return None
    if not run_id.replace("_", "").isdigit() or len(run_id) != 15:
        return None
    safe = _safe_name(sheet_name)
    filename = allowed[filetype].format(safe=safe)
    filepath = VERSIONS_DIR / run_id / filename
    return filepath if filepath.exists() else None


def get_current_version_file(sheet_name: str) -> Optional[Path]:
    p = _current_version_path(sheet_name)
    return p if p.exists() else None


def get_delta_file(sheet_name: str, filetype: str) -> Optional[Path]:
    paths = {
        "excel": _delta_xlsx_path(sheet_name),
        "xml": _delta_xml_path(sheet_name),
        "issues": _issues_csv_path(sheet_name),
        "issues_xlsx": _issues_xlsx_path(sheet_name),
        "full_issues": _full_validation_csv_path(sheet_name),
        "full_issues_xlsx": _full_validation_xlsx_path(sheet_name),
    }
    p = paths.get(filetype)
    return p if (p is not None and p.exists()) else None


def build_all_versions_zip() -> Optional[bytes]:
    """Bundles every archived run's per-sheet source snapshots into one
    ZIP, named by date and sheet."""
    import zipfile

    history = load_history()
    if not history:
        return None

    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for run in history:
            run_dir = VERSIONS_DIR / run["run_id"]
            for sheet_name in run.get("sheet_names", []):
                safe = _safe_name(sheet_name)
                src = run_dir / f"{safe}_snapshot.xlsx"
                if src.exists():
                    zf.write(src, arcname=f"{safe}_{run['run_id']}.xlsx")
    buf.seek(0)
    return buf.getvalue()
