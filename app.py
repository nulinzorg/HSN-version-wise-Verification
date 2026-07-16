"""
app.py — HSN/SAC Delta Validator (standalone)
-------------------------------------------------------------------------
Fully independent from the GST Regulatory Updates dashboard. One Flask
server, one page.

Run with:
    python app.py
Then open:
    http://localhost:8001

A background loop automatically re-checks the GST portal's HSN/SAC file
every HSN_CHECK_INTERVAL_SECONDS (default: 24 hours) for as long as this
process keeps running, and emails subscribers only when there's an
actual change since the last check. A manual "Refresh now" button on
the page triggers an on-demand check any time.

Still $0 cost — everything here is your own machine running open-source
libraries. Nothing hosted, nothing billed.
-------------------------------------------------------------------------
"""

import threading
import time
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory, send_file

from notifier import (
    add_email_subscriber,
    load_config,
    load_subscribers,
    remove_email_subscriber,
    send_hsn_email_alert,
)
import hsn_module

app = Flask(__name__)
BASE_DIR = Path(__file__).parent

HSN_CHECK_INTERVAL_SECONDS = 60 * 60 * 24  # every 24 hours by default

last_status = {"ran_at": None, "sheets": {}, "error": None}


def get_hsn_config():
    config = load_config() or {}
    hsn_config = config.get("hsn")
    if not hsn_config:
        raise RuntimeError(
            "No 'hsn' section found in config.json. Copy config.example.json to "
            "config.json and fill in hsn.key_col / hsn.compare_cols, or hsn.sheets "
            "for per-sheet column names (reference_file is optional)."
        )
    has_global = "key_col" in hsn_config and "compare_cols" in hsn_config
    has_per_sheet = bool(hsn_config.get("sheets"))
    if not has_global and not has_per_sheet:
        raise RuntimeError(
            "config.json's 'hsn' section needs either top-level 'key_col'/'compare_cols' "
            "(applied to every sheet), or a 'sheets' section giving each sheet its own "
            "key_col/compare_cols (needed when sheets use different column names, e.g. "
            "HSN_CD vs SAC_CD). See config.example.json for the format."
        )
    return hsn_config


def run_check_and_notify():
    hsn_config = get_hsn_config()
    summary = hsn_module.run_hsn_check(hsn_config)
    email_result = send_hsn_email_alert(summary)
    summary["email"] = email_result
    return summary


def background_loop():
    while True:
        try:
            summary = run_check_and_notify()
            per_sheet_counts = {name: s["counts"] for name, s in summary["sheets"].items()}
            last_status.update(ran_at=summary["ran_at"], sheets=per_sheet_counts, error=None)
            parts = []
            for name, s in summary["sheets"].items():
                c = s["counts"]
                parts.append(f"{name}: +{c['added']}/-{c['deleted']}/~{c['modified']}" + (" [baseline]" if s["is_baseline"] else ""))
            print(f"[hsn] {summary['ran_at']} — " + ", ".join(parts))
        except Exception as exc:  # noqa: BLE001
            last_status.update(ran_at=datetime.now().isoformat(timespec="seconds"), error=str(exc))
            print(f"[hsn] failed: {exc}")
        time.sleep(HSN_CHECK_INTERVAL_SECONDS)


# --------------------------------------------------------------------------
# Page
# --------------------------------------------------------------------------
@app.route("/")
def home():
    return send_from_directory(".", "hsn_dashboard.html")


# --------------------------------------------------------------------------
# HSN API
# --------------------------------------------------------------------------
@app.route("/api/hsn/data")
def hsn_data():
    return jsonify(hsn_module.load_last_hsn_result())


@app.route("/api/hsn/refresh", methods=["POST"])
def hsn_refresh():
    try:
        summary = run_check_and_notify()
        per_sheet_counts = {name: s["counts"] for name, s in summary["sheets"].items()}
        last_status.update(ran_at=summary["ran_at"], sheets=per_sheet_counts, error=None)
        return jsonify({"success": True, **summary})
    except Exception as exc:  # noqa: BLE001
        last_status.update(ran_at=datetime.now().isoformat(timespec="seconds"), error=str(exc))
        return jsonify({"success": False, "error": str(exc)}), 500


@app.route("/api/hsn/export/current")
def hsn_export_current():
    sheet_name = request.args.get("sheet", "")
    filepath = hsn_module.get_current_version_file(sheet_name)
    if filepath is None:
        return jsonify({"error": "No version downloaded yet for that sheet — run a refresh first."}), 404
    return send_file(filepath, as_attachment=True, download_name=f"HSN_SAC_current_{sheet_name}_{datetime.now().strftime('%Y%m%d')}.xlsx")


@app.route("/api/hsn/export/excel")
def hsn_export_excel():
    sheet_name = request.args.get("sheet", "")
    filepath = hsn_module.get_delta_file(sheet_name, "excel")
    if filepath is None:
        return jsonify({"error": "No delta export available yet for that sheet — run a refresh first (needs at least two checks)."}), 404
    return send_file(filepath, as_attachment=True, download_name=f"HSN_SAC_delta_{sheet_name}_{datetime.now().strftime('%Y%m%d')}.xlsx")


@app.route("/api/hsn/export/xml")
def hsn_export_xml():
    sheet_name = request.args.get("sheet", "")
    filepath = hsn_module.get_delta_file(sheet_name, "xml")
    if filepath is None:
        return jsonify({"error": "No delta export available yet for that sheet — run a refresh first (needs at least two checks)."}), 404
    return send_file(filepath, as_attachment=True, download_name=f"HSN_SAC_delta_{sheet_name}_{datetime.now().strftime('%Y%m%d')}.xml")


@app.route("/api/hsn/export/issues")
def hsn_export_issues():
    sheet_name = request.args.get("sheet", "")
    filepath = hsn_module.get_delta_file(sheet_name, "issues")
    if filepath is None:
        return jsonify({"error": "No validation issues on file for that sheet (either none were found, or no refresh has run yet)."}), 404
    return send_file(filepath, as_attachment=True, download_name=f"HSN_SAC_validation_issues_{sheet_name}_{datetime.now().strftime('%Y%m%d')}.csv")


@app.route("/api/hsn/export/issues_excel")
def hsn_export_issues_excel():
    sheet_name = request.args.get("sheet", "")
    filepath = hsn_module.get_delta_file(sheet_name, "issues_xlsx")
    if filepath is None:
        return jsonify({"error": "No validation issues on file for that sheet (either none were found, or no refresh has run yet)."}), 404
    return send_file(filepath, as_attachment=True, download_name=f"HSN_SAC_validation_issues_highlighted_{sheet_name}_{datetime.now().strftime('%Y%m%d')}.xlsx")


@app.route("/api/hsn/export/full_issues")
def hsn_export_full_issues():
    sheet_name = request.args.get("sheet", "")
    filepath = hsn_module.get_delta_file(sheet_name, "full_issues")
    if filepath is None:
        return jsonify({"error": "No issues found in this sheet's description field (or no refresh has run yet)."}), 404
    return send_file(filepath, as_attachment=True, download_name=f"HSN_SAC_description_validation_{sheet_name}_{datetime.now().strftime('%Y%m%d')}.csv")


@app.route("/api/hsn/export/full_issues_excel")
def hsn_export_full_issues_excel():
    sheet_name = request.args.get("sheet", "")
    filepath = hsn_module.get_delta_file(sheet_name, "full_issues_xlsx")
    if filepath is None:
        return jsonify({"error": "No issues found in this sheet's description field (or no refresh has run yet)."}), 404
    return send_file(filepath, as_attachment=True, download_name=f"HSN_SAC_description_validation_highlighted_{sheet_name}_{datetime.now().strftime('%Y%m%d')}.xlsx")


@app.route("/api/hsn/history")
def hsn_history():
    return jsonify(hsn_module.load_history())


@app.route("/api/hsn/history/<run_id>/export/<filetype>")
def hsn_history_export(run_id, filetype):
    sheet_name = request.args.get("sheet", "")
    filepath = hsn_module.get_run_file(run_id, sheet_name, filetype)
    if filepath is None:
        return jsonify({"error": "File not found for that run/sheet."}), 404
    return send_file(filepath, as_attachment=True, download_name=f"{run_id}_{filepath.name}")


@app.route("/api/hsn/history/export_all")
def hsn_history_export_all():
    zip_bytes = hsn_module.build_all_versions_zip()
    if zip_bytes is None:
        return jsonify({"error": "No archived versions yet."}), 404
    from io import BytesIO
    return send_file(
        BytesIO(zip_bytes),
        as_attachment=True,
        download_name=f"HSN_SAC_all_versions_{datetime.now().strftime('%Y%m%d')}.zip",
        mimetype="application/zip",
    )


@app.route("/api/subscribe", methods=["POST"])
def subscribe():
    email = (request.get_json(silent=True) or {}).get("email", "").strip()
    if not email or "@" not in email:
        return jsonify({"success": False, "error": "Please provide a valid email address"}), 400
    data = add_email_subscriber(email)
    return jsonify({"success": True, "subscribers": data["emails"]})


@app.route("/api/subscribe", methods=["DELETE"])
def unsubscribe():
    email = (request.get_json(silent=True) or {}).get("email", "").strip()
    data = remove_email_subscriber(email)
    return jsonify({"success": True, "subscribers": data["emails"]})


@app.route("/api/subscribers")
def subscribers():
    return jsonify(load_subscribers())


@app.route("/api/status")
def status():
    return jsonify({**last_status, "interval_seconds": HSN_CHECK_INTERVAL_SECONDS})


@app.route("/<path:filename>")
def other_static(filename):
    return send_from_directory(".", filename)


if __name__ == "__main__":
    try:
        get_hsn_config()
        threading.Thread(target=background_loop, daemon=True).start()
        note = f"Auto-checking every {HSN_CHECK_INTERVAL_SECONDS // 3600}h in the background"
    except RuntimeError as e:
        note = f"Background auto-check NOT started — {e}"

    print("Starting HSN/SAC Delta Validator at http://localhost:8001")
    print(f"  {note}")
    app.run(host="0.0.0.0", port=8001, debug=False)
