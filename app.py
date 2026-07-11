"""
Local web interface for the IPO shareholding Excel generator.

Run:  python app.py
Then open http://127.0.0.1:5000 in your browser.
"""

import os
import threading
import datetime
import traceback

from flask import (Flask, render_template, request, jsonify,
                   send_from_directory)

from pipeline import run_pipeline

BASE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE, "output")
os.makedirs(OUT_DIR, exist_ok=True)

app = Flask(__name__)

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
}
_lock = threading.Lock()


def _reset():
    _job.update(running=False, log=[], current=0, total=0, stage="",
                finished=False, error=None, file=None, count=None)


def _progress(msg, current=None, total=None, stage=None):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    _job["log"].append(f"[{ts}] {msg}")
    if current is not None:
        _job["current"] = current
    if total is not None:
        _job["total"] = total
    if stage is not None:
        _job["stage"] = stage


def _worker(from_str, to_str):
    fname = f"IPO_Companies_{from_str}_to_{to_str}.xlsx"
    out_path = os.path.join(OUT_DIR, fname)
    try:
        companies = run_pipeline(from_str, to_str, out_path,
                                 progress_cb=_progress)
        _job["file"] = fname
        _job["count"] = len(companies)
    except Exception as e:
        _job["error"] = str(e)
        _progress("ERROR: " + str(e))
        traceback.print_exc()
    finally:
        _job["finished"] = True
        _job["running"] = False


@app.route("/")
def index():
    return render_template("index.html")


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

    with _lock:
        if _job["running"]:
            return jsonify({"error": "A generation job is already running."}), 409
        _reset()
        _job["running"] = True

    threading.Thread(target=_worker, args=(from_str, to_str),
                     daemon=True).start()
    return jsonify({"ok": True})


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
    })


@app.route("/download/<path:fname>")
def download(fname):
    return send_from_directory(OUT_DIR, fname, as_attachment=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, threaded=True, debug=False)
