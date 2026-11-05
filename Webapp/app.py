#!/usr/bin/env python3
"""Flask front end for load_data_status_report.py  ->  python app.py  (opens http://127.0.0.1:5000)"""
import json
import logging
import threading
import webbrowser
from datetime import date, datetime
from pathlib import Path

from flask import Flask, abort, jsonify, render_template, request, send_from_directory
from werkzeug.utils import secure_filename

import load_data_status_report as lr

BASE = Path(__file__).resolve().parent
CONFIG = BASE / "config.json"
UPLOADS = BASE / "uploads"
UPLOADS.mkdir(exist_ok=True)

app = Flask(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

LOCK = threading.Lock()
JOB = {"state": "idle", "done": 0, "total": 0, "error": None, "result": None}


def get_cfg():
    cfg = lr.load_config(CONFIG)                       # re-read on every call: config edits apply at once
    out = Path(cfg.get("output_dir", "reports"))
    cfg["output_dir"] = str(out if out.is_absolute() else BASE / out)
    Path(cfg["output_dir"]).mkdir(parents=True, exist_ok=True)
    return cfg


def parse_day(txt):
    return datetime.strptime(txt, "%Y-%m-%d").date()


@app.get("/")
def index():
    cfg = get_cfg()
    history = lr.load_history(Path(cfg["output_dir"]) / "status_history.json")
    rep = cfg.get("report", {})
    prev = rep.get("previous_checked_date") or (sorted(history)[-1] if history else "")
    cur = rep.get("current_checked_date") or date.today().isoformat()
    clients = [{"id": c["client"], "name": c["name"], "channels": sum(1 for ch in c["channels"] if ch.get("enabled", True))}
               for c in cfg["clients"] if c.get("enabled", True)]
    return render_template("index.html", clients=clients, prev=prev, cur=cur, has_history=bool(history))


@app.post("/api/run")
def api_run():
    try:
        prev, cur = parse_day(request.form["previous"]), parse_day(request.form["current"])
    except (KeyError, ValueError):
        return jsonify(error="Enter both dates."), 400
    if prev >= cur:
        return jsonify(error="The previous check must be earlier than the current check."), 400
    ids = [int(x) for x in request.form.getlist("clients") if x.isdigit()]
    demo = request.form.get("demo") == "1"
    prev_report = None
    f = request.files.get("prev_report")
    if f and f.filename:
        prev_report = UPLOADS / secure_filename(f.filename)
        f.save(prev_report)

    with LOCK:
        if JOB["state"] == "running":
            return jsonify(error="A check is already running."), 409
        JOB.update(state="running", done=0, total=0, error=None, result=None)

    def work():
        try:
            def progress(done, total):
                JOB.update(done=done, total=total)
            res = lr.run_report(get_cfg(), prev, cur, clients=ids or None, demo=demo,
                                prev_report=str(prev_report) if prev_report else None, progress=progress)
            JOB.update(state="done", result={"file": res["file"], "summary": res["summary"]})
        except BaseException as e:                       # SystemExit too
            logging.exception("Run failed")
            JOB.update(state="error", error=str(e) or e.__class__.__name__)

    threading.Thread(target=work, daemon=True).start()
    return jsonify(ok=True)


@app.get("/api/status")
def api_status():
    return jsonify(JOB)


@app.get("/api/reports")
def api_reports():
    out = Path(get_cfg()["output_dir"])
    items = []
    for j in sorted(out.glob("Load_Status_*.json"), key=lambda p: p.stat().st_mtime, reverse=True):
        try:
            m = json.loads(j.read_text(encoding="utf-8"))
        except ValueError:
            continue
        items.append({"file": j.with_suffix(".xlsx").name, "previous": m["previous"], "current": m["current"],
                      "generated": m["generated"], "demo": m.get("demo", False), "summary": m["summary"]})
    return jsonify(items)


@app.get("/api/report/<name>")
def api_report(name):
    out = Path(get_cfg()["output_dir"])
    p = out / (Path(secure_filename(name)).stem + ".json")
    if not p.exists():
        abort(404)
    return app.response_class(p.read_text(encoding="utf-8"), mimetype="application/json")


@app.get("/download/<name>")
def download(name):
    name = secure_filename(name)
    if not name.endswith(".xlsx"):
        abort(404)
    return send_from_directory(get_cfg()["output_dir"], name, as_attachment=True)


if __name__ == "__main__":
    threading.Timer(1.0, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
