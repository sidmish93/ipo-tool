"""
Local web interface for the IPO shareholding Excel generator.

Run:  python app.py
Then open http://127.0.0.1:5000 in your browser.
"""

import os
import json
import threading
import datetime
import traceback

from flask import (Flask, render_template, request, jsonify,
                   send_from_directory, send_file)

from pipeline import (run_pipeline, lookup_company, split_terms,
                   _write_workbook, MCAP_THRESHOLD, Cancelled)

BASE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE, "output")
os.makedirs(OUT_DIR, exist_ok=True)
JOB_PATH = os.path.join(OUT_DIR, "job.json")

app = Flask(__name__)
app.config["TEMPLATES_AUTO_RELOAD"] = True

# single-job state (this tool is meant for one user at a time)
_job = {
    "running": False,
    "log": [],
    "current": 0,
    "total": 0,
    "stage": "",
    "finished": False,
    "error": None,
    "file": None,
    "count": None,
    "cancelled": False,
}
# one browser operation at a time (generation OR single lookup) — protects RAM
_gate = threading.Lock()
_cancel = threading.Event()
# last successful on-screen lookup, reused by Download Excel
_last_lookup = {"key": "", "results": []}


def _lookup_key(q):
    return ";".join(t.lower() for t in split_terms(q))


def _found_companies(results):
    return [r for r in (results or []) if r.get("found")]


def _save_job():
    try:
        tmp = JOB_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(_job, fh)
        os.replace(tmp, JOB_PATH)
    except OSError:
        pass


def _reset():
    _job.update(running=False, log=[], current=0, total=0, stage="",
                finished=False, error=None, file=None, count=None,
                cancelled=False)
    _save_job()


def _progress(msg, current=None, total=None, stage=None, file=None, count=None,
              **_extra):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    _job["log"].append(f"[{ts}] {msg}")
    if current is not None:
        _job["current"] = current
    if total is not None:
        _job["total"] = total
    if stage is not None:
        _job["stage"] = stage
    if file is not None:
        _job["file"] = file
    if count is not None:
        _job["count"] = count
    _save_job()


def _recover_job():
    """If Render OOM-killed us mid-run, keep the log and mark it failed."""
    try:
        with open(JOB_PATH, encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, ValueError):
        return
    if not isinstance(data, dict):
        return
    for key in _job:
        if key in data:
            _job[key] = data[key]
    if _job.get("running"):
        _job["running"] = False
        _job["finished"] = True
        _job["error"] = (
            "Server restarted during the run. On Render this is usually "
            "Chrome running out of memory when BSE opens. Use a 2 GB "
            "instance (Docker runtime) and try again."
        )
        _job.setdefault("log", [])
        if not isinstance(_job["log"], list):
            _job["log"] = []
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        _job["log"].append(f"[{ts}] ERROR: {_job['error']}")
        _save_job()


_recover_job()


def _parse_mcap_min(raw):
    """UI threshold in Rs cr. Blank → 3000. Reject negatives / junk."""
    if raw is None or raw == "":
        return MCAP_THRESHOLD
    try:
        value = float(raw)
    except (TypeError, ValueError):
        raise ValueError("Market-cap threshold must be a number.")
    if value < 0:
        raise ValueError("Market-cap threshold cannot be negative.")
    return value


def _worker(from_str, to_str, mcap_min):
    fname = f"IPO_Companies_{from_str}_to_{to_str}.xlsx"
    out_path = os.path.join(OUT_DIR, fname)
    try:
        companies = run_pipeline(from_str, to_str, out_path,
                                 progress_cb=_progress, mcap_min=mcap_min,
                                 cancel_cb=_cancel.is_set)
        _job["file"] = fname
        _job["count"] = len(companies)
    except Cancelled:
        _job["cancelled"] = True
        _progress("Cancelled.")
    except Exception as e:
        _job["error"] = str(e)
        _progress("ERROR: " + str(e))
        traceback.print_exc()
    finally:
        _job["finished"] = True
        _job["running"] = False
        _save_job()
        _gate.release()


@app.route("/healthz")
def healthz():
    return "ok", 200


@app.route("/")
def index():
    if not _job.get("running") and _job.get("cancelled"):
        _reset()
    resp = app.make_response(render_template("index.html"))
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/start", methods=["POST"])
def start():
    data = request.get_json(force=True)
    from_str = data.get("from")
    to_str = data.get("to")
    try:
        d1 = datetime.datetime.strptime(from_str, "%Y-%m-%d").date()
        d2 = datetime.datetime.strptime(to_str, "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return jsonify({"error": "Please provide valid start and end dates."}), 400
    if d1 > d2:
        return jsonify({"error": "Start date must be on or before end date."}), 400
    try:
        mcap_min = _parse_mcap_min(data.get("mcap_min"))
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    if not _gate.acquire(blocking=False):
        return jsonify({"error": "Server is busy (a generation or lookup is "
                                 "running). Please try again shortly."}), 409
    _reset()
    _cancel.clear()
    _job["running"] = True
    _job["mcap_min"] = mcap_min
    _save_job()

    threading.Thread(target=_worker, args=(from_str, to_str, mcap_min),
                     daemon=True).start()
    return jsonify({"ok": True})


@app.route("/company", methods=["POST"])
def company():
    data = request.get_json(force=True)
    q = (data.get("q") or "").strip()
    terms = split_terms(q)
    if not terms:
        return jsonify({"error": "Please enter a company name."}), 400
    if len(terms) > 15:
        return jsonify({"error": "Please look up at most 15 companies at once."}), 400

    if not _gate.acquire(blocking=False):
        return jsonify({"error": "Server is busy (a generation or lookup is "
                                 "running). Please try again shortly."}), 409
    try:
        results = lookup_company(q)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500
    finally:
        _gate.release()
    _last_lookup["key"] = _lookup_key(q)
    _last_lookup["results"] = results
    return jsonify({
        "ok": True,
        "queries": terms,
        "results": results,
        "found_count": sum(1 for r in results if r.get("found")),
        "missing": [r["query"] for r in results if not r.get("found")],
    })


@app.route("/company.xlsx", methods=["POST"])
def company_xlsx():
    data = request.get_json(force=True)
    q = (data.get("q") or "").strip()
    terms = split_terms(q)
    if not terms:
        return jsonify({"error": "Please enter a company name."}), 400
    if len(terms) > 15:
        return jsonify({"error": "Please look up at most 15 companies at once."}), 400

    key = _lookup_key(q)
    if _last_lookup["key"] == key and _last_lookup["results"]:
        results = _last_lookup["results"]
    else:
        if not _gate.acquire(blocking=False):
            return jsonify({"error": "Server is busy (a generation or lookup is "
                                     "running). Please try again shortly."}), 409
        try:
            results = lookup_company(q)
        except Exception as e:
            traceback.print_exc()
            return jsonify({"error": str(e)}), 500
        finally:
            _gate.release()
        _last_lookup["key"] = key
        _last_lookup["results"] = results

    found = _found_companies(results)
    if not found:
        missing = ", ".join(r.get("query") or "" for r in results if not r.get("found"))
        return jsonify({"error": "No BSE match to put in Excel"
                                 + (f" for: {missing}" if missing else ".")}), 404

    today = datetime.date.today().isoformat()
    fname = f"Shareholding_{len(found)}_companies_{today}.xlsx"
    out_path = os.path.join(OUT_DIR, fname)
    _write_workbook(found, out_path, sheet1_title="Companies")
    return send_file(out_path, as_attachment=True, download_name=fname)


@app.route("/cancel", methods=["POST"])
def cancel():
    if not _job.get("running"):
        return jsonify({"ok": True, "running": False})
    _cancel.set()
    _progress("Cancel requested — stopping after the current company…")
    return jsonify({"ok": True, "running": True})


@app.route("/status")
def status():
    return jsonify({
        "running": _job["running"],
        "log": _job["log"][-400:],
        "current": _job["current"],
        "total": _job["total"],
        "stage": _job["stage"],
        "finished": _job["finished"],
        "error": _job["error"],
        "file": _job["file"],
        "count": _job["count"],
        "mcap_min": _job.get("mcap_min"),
        "cancelled": _job.get("cancelled"),
    })


@app.route("/download/<path:fname>")
def download(fname):
    return send_from_directory(OUT_DIR, fname, as_attachment=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, threaded=True, debug=False)
