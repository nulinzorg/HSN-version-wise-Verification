# HSN / SAC Delta Validator — standalone, $0 cost

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

Fully independent tool. No dependency on any other dashboard.

**HSN and SAC are tracked as completely separate lists.** The portal's
file has them on separate sheets, and this tool never merges or
cross-compares them — an HSN row is only ever compared to a previous HSN
row, a SAC row only ever to a previous SAC row. Each sheet gets its own
independent snapshot history, its own independent Added/Deleted/Modified
counts, and its own independent downloads. The dashboard has a sheet
selector at the top to switch between them.

```
app.py                  → Flask server (the whole app)
hsn_core.py               → download + per-sheet comparison/validation logic
hsn_module.py               → per-sheet archiving, version history, ZIP export
hsn_dashboard.html            → the page itself, with a sheet selector
notifier.py                     → email alerts (per-sheet breakdown)
subscribers.example.json         → template — copy to subscribers.json
config.example.json                → template — copy to config.json
.gitignore                           → keeps secrets and generated data out of git
```

## 0. Get the code

```bash
git clone <your-repo-url>
cd hsn_standalone
```

(Or download/extract the ZIP from GitHub's "Code" button if you don't use git.)

## 1. Set up (one-time)

```bash
pip install -r requirements.txt
```
(On Windows, if `pip` isn't recognized, use `python -m pip install -r requirements.txt` instead.)

## 2. Configure

```bash
cp config.example.json config.json
cp subscribers.example.json subscribers.json
```
(Windows: use `copy` instead of `cp`.)

Both `config.json` and `subscribers.json` are in `.gitignore` on purpose —
one holds your email password, the other will hold real people's email
addresses once you start subscribing. Neither should ever be committed.

Edit `config.json`:

- **SMTP settings** (`smtp_host`, `smtp_user`, `smtp_password`) — optional,
  only needed if you want email alerts. For Gmail: turn on 2-Step
  Verification, create an App Password at
  `myaccount.google.com/apppasswords`, use that (not your normal
  password).
- **`hsn` section** — **different sheets can use different column
  names** (the portal's actual file has separate `HSN_MSTR` and
  `SAC_MSTR` sheets with their own column names), so give each sheet its
  own `key_col`/`compare_cols` under `sheets`, keyed by the exact sheet
  name:

  ```json
  "hsn": {
    "sheets": {
      "HSN_MSTR": { "key_col": "HSN_CD", "compare_cols": ["HSN_Description"] },
      "SAC_MSTR": { "key_col": "SAC_CD", "compare_cols": ["SAC_Description"] }
    },
    "auto_clean_xml": false,
    "reference_file": null
  }
  ```

  If you're not sure of the exact sheet/column names, run it once anyway
  — a failed check's error message lists the real sheet name and every
  real column found in it, so you can copy the exact values back into
  `config.json`.
  - **No reference file is needed.** The tool keeps its own permanent
    record of every version it pulls and automatically compares each new
    pull against the previous one, per sheet.
  - Optionally set `hsn.reference_file` if you *also* want to compare
    against a fixed master list you maintain — this adds a secondary
    comparison alongside the automatic one. Leave as `null` otherwise.

## 3. Run

```bash
python app.py
```

Open **http://localhost:8001**

(Runs on port 8001 by default so it doesn't clash with the GST Regulatory
dashboard if that's also running on port 8000. To change the port, edit
the `app.run(..., port=8001, ...)` line near the bottom of `app.py`.)

## 4. Use it

- **Sheet selector** at the top shows a pill per sheet found in the file
  (e.g. `HSN_MSTR`, `SAC_MSTR`), each with a live status badge (baseline
  / no changes / N changes). Click a pill to switch — everything below
  reflects whichever sheet is selected. HSN and SAC are never mixed.
- **The page checks the portal automatically the moment it loads** —
  you don't need to click anything for the first check of a session.
  It shows whatever was last saved instantly, then immediately kicks off
  a live check in the background and updates once that finishes (you'll
  briefly see "Checking GST portal for the latest version…").
- Click **"Refresh now"** any time afterward for an on-demand re-check.
  The very first check ever (per sheet) has nothing to compare against
  yet and is recorded as a baseline; every check after that shows real
  Added/Deleted/Modified counts vs the previous check.
- **"Download current version (Excel)"** is always available after any
  run, including the very first baseline run — the full current sheet
  as pulled from the portal, nothing filtered out.
- **"Highlighted Excel (vs previous check)"** and **"Delta XML"** only
  appear once there's an actual previous version for that sheet to
  compare against (i.e. not on the baseline run).
- **Description validation** (invisible characters, stray whitespace,
  special/XML-reserved characters) comes in two independent scopes:
  - **Delta only** — scans just this check's Added/Modified rows. Only
    appears if something in *this specific check* actually has an issue.
  - **Full list** — scans every row currently in the sheet, every time,
    regardless of what changed. Always available after any check.
  Each scope has both a plain CSV and a **highlighted Excel** version —
  the Excel is filtered to only the problem rows, with the exact
  offending cell colored, so the issue is visible in context rather than
  just listed as text.
- **Added / Deleted / Modified** tabs show the selected sheet's delta.
- **Version History** tab lists every run ever made for the selected
  sheet, with download links for that day's source snapshot and delta
  files, plus a **"Download ALL versions, all sheets (ZIP)"** button to
  grab everything at once.
- **Email me when a delta is found** box subscribes an address — only
  fires when there's an actual change on any sheet (never on baseline or
  no-change runs).

## 5. What happens automatically

While `python app.py` is running, it re-checks the portal every
**24 hours** in the background and emails subscribers only if something
actually changed. Stop it with `Ctrl+C` — the background loop stops with it.

### Changing the interval

Open `app.py`, near the top:
```python
HSN_CHECK_INTERVAL_SECONDS = 60 * 60 * 24  # every 24 hours
```
Change the value (in seconds) and restart.

## 6. Version history — permanent, downloadable record

No reference file required for any of this. Every check:

1. Downloads the source and saves a permanent, timestamped snapshot —
   independent of whatever the portal does with its own copy afterward.
2. Compares it against the **previous run's snapshot** (the primary
   comparison — "what changed since the last check").
3. Optionally also compares against your **reference file**, if configured.
4. Archives everything under `hsn_versions/<run_id>/`:
   - `source_snapshot.xlsx` — the exact file downloaded that day
   - `delta_vs_previous_version.xlsx` / `.xml`
   - `delta_vs_reference.xlsx` / `.xml` — only if a reference file is configured
   - `validation_issues.csv` — if any invisible-character issues were found
   - `summary.json`
5. Appends an entry to `hsn_history.json`.

Nothing is ever overwritten — `hsn_versions/` grows by one dated folder
per run, forever, so you always have a complete downloadable archive.

## Troubleshooting

**"403 Forbidden" or refresh fails immediately:** already fixed in this
version — `hsn_core.py` sends real browser-like headers with every
request (some government sites block the default Python `requests`
identity outright). If it still fails, try opening
`https://tutorial.gst.gov.in/downloads/HSN_SAC.xlsx` directly in your own
browser on the same machine — if that also fails, it's a network/firewall
issue on your end, not the script.

**`pip` not recognized (Windows):** use `python -m pip install ...`
instead, or reinstall Python from python.org with "Add python.exe to
PATH" checked.

## Cost breakdown

| Piece | Cost |
|---|---|
| Running the check (manual or scheduled) | $0 — local Python |
| Serving the dashboard | $0 — Flask dev server on your own machine |
| Storing every version | $0 — files on disk |
| Scheduling (built-in background loop) | $0 |
| Email alerts | $0 — your own SMTP account |
| **Total** | **$0** |

## Contributing

Bug reports, feature suggestions, and pull requests are welcome — see
[CONTRIBUTING.md](CONTRIBUTING.md).

## Disclaimer

This is an independent, unofficial tool. It is **not affiliated with,
endorsed by, or officially connected to** the Goods and Services Tax
Network (GSTN), the Government of India, or the GST portal in any way.

It works by downloading the publicly available HSN/SAC Excel file that
the GST portal itself already offers via a "Download" link on its
public-facing HSN/SAC search page — the same file anyone can get by
clicking that link in a browser. This tool automates that download and
compares versions over time; it does not access any private, restricted,
or authenticated data.

The government portal's URL, file structure, column names, and available
sheets can change at any time without notice (this project has already
needed several column-mapping fixes for exactly that reason). Always
verify important compliance decisions against the official portal
directly — **use this tool's output as a convenience aid, not as a sole
source of truth for regulatory compliance.**

No warranty is provided — see [LICENSE](LICENSE) (MIT).

## License

MIT — see [LICENSE](LICENSE). Free to use, modify, and redistribute.
