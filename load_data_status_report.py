#!/usr/bin/env python3
"""
Load-data status report automation
==================================
For every enabled client/channel in config.json the script pulls 15-min load data
from the enexion API ONLY for the window  previous-checked-date -> current-checked-date
(one API call per channel) and writes an Excel report:

  Client | Meter Name | Channel Name | Previous Status | Current Status | Reason |
  Previous checked date | Current checked date

* Current Status  = evaluated from the data in the window.
* Previous Status = taken from the last saved run (reports/status_history.json),
                    or from an older Excel report (config "report.previous_report"
                    or --prev-report) the first time you run it.

Dates come from config.json -> "report". If previous_checked_date is empty the date
of the last saved run is used; if current_checked_date is empty today is used.

Usage
-----
  python load_data_status_report.py                      # all clients
  python load_data_status_report.py --client 27          # only client id 27
  python load_data_status_report.py --probe 27 17        # test ONE channel (client, channel)
  python load_data_status_report.py --prev-report old.xlsx
  python load_data_status_report.py --demo               # synthetic data, no API needed
  python load_data_status_report.py --no-save            # do not update status_history.json

Requires: pip install pandas numpy openpyxl requests
"""
import argparse
import io
import json
import logging
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

log = logging.getLogger("loadreport")

OK, BAD, NONE = "\u221a", "\u00d7", "-"      # √  ×  -

# Rules for deciding √ / ×. Override by adding "status_rules" to config.json.
DEFAULT_RULES = {
    "max_lag_days": 7,          # × if last data point is older than the window end by more than this
    "min_coverage_pct": 90,     # below this % (gaps) status stays √ but Reason is written in red
    "spike_iqr_factor": 3.0,    # daily max above Q3 + factor*IQR of daily maxes => "high pick"
    "gap_min_minutes": 60,      # list gaps (missing 15-min data) of at least this length in the Reason
    "gap_max_listed": 5,        # max number of gaps written into the Reason (largest ones)
    "gas_interval_minutes": 60, # Gas channels: hourly values
    "gas_day_start_hour": 6,    # Gas day = 06:00 -> 06:00 next day (last value 05:00)
    "low_level_pct": 50,        # weekday mean below this % of the window median => "data is low"
    "request_timeout": 60,
    "retries": 3,
    "workers": 8,
}

# Field-name hints used to find timestamp / value columns in the API reply.
TS_HINTS = ("time", "date", "stamp", "ts", "zeit", "datum", "from", "start")
VAL_HINTS = ("value", "val", "load", "kw", "kwh", "wert", "power", "energy", "v")


# =============================================================================
# Config / dates / history
# =============================================================================
def load_config(path):
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    cfg["rules"] = {**DEFAULT_RULES, **cfg.get("status_rules", {})}
    return cfg


def fmt_d(d):
    return pd.Timestamp(d).strftime("%d-%m-%Y")


def key_of(meter_name, channel_name):
    return f"{str(meter_name).strip()} || {str(channel_name).strip()}"


def load_history(path):
    if Path(path).exists():
        try:
            return json.loads(Path(path).read_text(encoding="utf-8"))
        except ValueError:
            log.warning("Could not read %s - starting with empty history", path)
    return {}


def norm_status(v):
    t = str(v or "").strip().lower()
    if t in ("\u221a", "\u2713", "\u2714", "v", "ok", "yes"):
        return OK
    if t in ("\u00d7", "x", "\u2717", "\u2718", "no"):
        return BAD
    return NONE


def read_report_statuses(path):
    """Read Current Status column of an older report -> {key: status}."""
    ws = load_workbook(path, data_only=True).active
    out, meter = {}, None
    for row in ws.iter_rows(min_row=4, values_only=True):
        if row[1]:
            meter = str(row[1]).strip()          # merged cells: only first row holds the name
        if row[2] and meter:
            out[key_of(meter, row[2])] = norm_status(row[4])
    return out


def resolve_dates(cfg, history, today):
    """Return (previous_checked, current_checked) as dates."""
    rep = cfg.get("report", {})
    d = lambda x: datetime.strptime(x, "%Y-%m-%d").date()
    cur = d(rep["current_checked_date"]) if rep.get("current_checked_date") else today
    if rep.get("previous_checked_date"):
        prev = d(rep["previous_checked_date"])
    else:
        earlier = sorted(k for k in history if k < cur.isoformat())
        if not earlier:
            sys.exit('Set "report": {"previous_checked_date": "YYYY-MM-DD"} in config.json '
                     "(no earlier run saved yet).")
        prev = d(earlier[-1])
    if prev >= cur:
        sys.exit(f"previous_checked_date ({prev}) must be before current_checked_date ({cur}).")
    return prev, cur


def previous_statuses(history, prev):
    """Statuses of the latest saved run dated on/before the previous checked date."""
    dates = sorted(k for k in history if k <= prev.isoformat())
    return history[dates[-1]] if dates else {}


# =============================================================================
# API access + parsing
# =============================================================================
def _find_records(obj):
    """Dig through nested JSON until we find a list of records."""
    if isinstance(obj, str):
        try:
            return _find_records(json.loads(obj))
        except ValueError:
            return None
    if isinstance(obj, list):
        return obj
    if isinstance(obj, dict):
        # {"d": ...} wrapper from ASMX, or {"data":[...]}, or {timestamp: value}
        for k in ("d", "data", "Data", "values", "Values", "result", "Result", "items"):
            if k in obj:
                r = _find_records(obj[k])
                if r is not None:
                    return r
        if obj and all(isinstance(v, (int, float, type(None))) for v in obj.values()):
            return [{"ts": k, "value": v} for k, v in obj.items()]
        for v in obj.values():
            if isinstance(v, (list, dict, str)):
                r = _find_records(v)
                if r:
                    return r
    return None


def _records_from_csv_text(inner):
    """enexion format: title line, 'Datum;Zeit;kWh;...' header, then ';' separated rows."""
    lines = [l.strip() for l in inner.replace("\r", "").split("\n")]
    h = next((i for i, l in enumerate(lines) if ";" in l and "datum" in l.lower()), None)
    if h is None:
        return None
    header = [c.strip() for c in lines[h].split(";")]
    low = [c.lower() for c in header]
    di = low.index("datum")
    zi = low.index("zeit") if "zeit" in low else None
    recs = []
    for l in lines[h + 1:]:
        if not l or l.lower().startswith("no data") or ";" not in l:
            continue
        parts = l.split(";")
        parts += [""] * (len(header) - len(parts))
        d, t = parts[di].strip(), (parts[zi].strip() if zi is not None else "")
        if t.startswith("24:00"):                      # 24:00 = midnight of the next day
            try:
                d = (pd.to_datetime(d, dayfirst="." in d, format="mixed") + pd.Timedelta(days=1)).strftime("%d.%m.%Y")
            except Exception:
                pass
            t = "00:00"
        rec = {"Timestamp": f"{d} {t}".strip()}
        for i, c in enumerate(header):
            if i not in (di, zi):
                rec[c] = parts[i].strip()
        recs.append(rec)
    return recs


def _records_from_xml(text):
    import xml.etree.ElementTree as ET
    root = ET.fromstring(text)
    inner = (root.text or "").strip()
    if ";" in inner and "datum" in inner.lower():
        return _records_from_csv_text(inner)
    if inner[:1] in ("[", "{"):                       # <string>[...json...]</string>
        return _find_records(inner)
    recs = []
    for el in root.iter():
        kids = list(el)
        if kids and all(len(k) == 0 for k in kids) and any((k.text or "").strip() for k in kids):
            recs.append({k.tag.split("}")[-1]: k.text for k in kids})
    if not recs:                                  # <V t="2026-10-01T00:00" v="1.5"/> style
        for el in root.iter():
            if len(el) == 0 and len(el.attrib) >= 2:
                recs.append(dict(el.attrib))
    return recs


_GENUINE_EMPTY = ("", "[]", "{}", "null", '""', '{"d":[]}', '{"d":null}', '{"d":"[]"}', '{"d":""}')


def _empty(text):
    s = pd.Series(dtype=float)
    compact = "".join(text.split())
    if "nodataavailable" in compact.lower():
        s.attrs.update(n_records=0, n_valid_ts=0, n_values=0, unparsed=False)
        return s
    s.attrs.update(n_records=0, n_valid_ts=0, n_values=0,
                   unparsed=compact not in _GENUINE_EMPTY and not (compact.startswith("<") and len(compact) < 120))
    return s


def parse_response(text):
    """Turn whatever the API returned into a pandas Series (index=timestamp, values=float)."""
    text = (text or "").strip()
    if not text:
        return _empty(text)

    recs = None
    if text[0] in "[{\"":
        recs = _find_records(text)
    elif text[0] == "<":
        recs = _records_from_xml(text)
    else:                                       # maybe CSV
        try:
            df = pd.read_csv(io.StringIO(text), sep=None, engine="python")
            recs = df.to_dict("records")
        except Exception:
            recs = None
    if not recs:
        return _empty(text)

    # list of [ts, value] pairs
    if isinstance(recs[0], (list, tuple)):
        ts = [r[0] for r in recs]
        vals = [r[1] if len(r) > 1 else None for r in recs]
    else:
        keys = list(recs[0].keys())
        lower = {k: str(k).lower() for k in keys}
        ts_key = next((k for k in keys if any(lower[k] == h for h in TS_HINTS)), None) or \
                 next((k for k in keys if any(h in lower[k] for h in TS_HINTS)), keys[0])
        skip = ("id", "channel", "client", "meter", "status", "quality", "flag", "unit", "name", "malo")
        others = [k for k in keys if k != ts_key and not any(x in lower[k] for x in skip)] or \
                 [k for k in keys if k != ts_key]
        hinted = [k for k in others if lower[k] in VAL_HINTS or
                  any(h in lower[k] for h in VAL_HINTS if len(h) > 1)]

        def numeric_share(k):
            v = _to_num(pd.Series([r.get(k) for r in recs[:500]]))
            return v.notna().mean() if len(v) else 0
        cands = hinted + [k for k in others if k not in hinted]
        val_key = next((k for k in cands if numeric_share(k) > 0), None)
        ts = [r.get(ts_key) for r in recs]
        vals = [r.get(val_key) for r in recs] if val_key else [None] * len(recs)

    try:
        s = pd.Series(_to_num(pd.Series(vals)).values, index=_to_dt(pd.Series(ts)))
    except ValueError:                      # timestamps missing / not found
        return _empty(text)
    n_rec = len(s)
    s = s[~s.index.isna()]
    if getattr(s.index, "tz", None) is not None:
        s.index = s.index.tz_convert("Europe/Berlin").tz_localize(None)
    s = s.sort_index()
    s.attrs.update(n_records=n_rec, n_valid_ts=len(s), n_values=int(s.notna().sum()))
    return s


def _to_num(ser):
    """Numbers, also German style '1.234,5' / '1,5' and empty strings."""
    ser = ser.astype(object).where(ser.notna(), None)
    def conv(v):
        if v is None or v == "":
            return np.nan
        if isinstance(v, (int, float, np.number)):
            return float(v)
        t = str(v).strip().replace("\u00a0", "")
        if "," in t:
            t = t.replace(".", "").replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return np.nan
    return ser.map(conv).astype(float)


def _to_dt(ser):
    """ISO, /Date(ms)/, epoch and German dd.mm.yyyy timestamps."""
    ser = ser.astype(object)
    sample = ser.dropna().astype(str)
    if sample.empty:
        return pd.DatetimeIndex([])
    if sample.str.match(r"^/Date\(").all():
        ms = pd.to_numeric(ser.astype(str).str.extract(r"(-?\d+)")[0], errors="coerce")
        return pd.DatetimeIndex(pd.to_datetime(ms, unit="ms", utc=True).dt.tz_convert("Europe/Berlin").dt.tz_localize(None))
    num = pd.to_numeric(sample, errors="coerce")
    if num.notna().all():
        unit = "ms" if num.abs().median() > 1e11 else "s"
        return pd.DatetimeIndex(pd.to_datetime(pd.to_numeric(ser, errors="coerce"), unit=unit))
    if sample.str.match(r"^\d{1,2}\.\d{1,2}\.\d{4}").mean() > 0.5:   # German format
        return pd.DatetimeIndex(pd.to_datetime(ser, errors="coerce", dayfirst=True, format="mixed"))
    return pd.DatetimeIndex(pd.to_datetime(ser, errors="coerce", format="mixed"))


_warned = [0]


def fetch_series(cfg, client_id, channel_id, start, end, debug_path=None):
    api = cfg["api"]
    params = {
        "token": api["token"],
        "fr": api.get("resolution", "min15"),
        "fillValue": "",
        "to": end.isoformat(),
        "from": start.isoformat(),
        "channel": channel_id,
        "client": client_id,
    }
    rules = cfg["rules"]
    last_err = None
    for attempt in range(1, rules["retries"] + 1):
        try:
            r = requests.get(api["url"], params=params, timeout=rules["request_timeout"])
            r.raise_for_status()
            if debug_path:
                Path(debug_path).write_text(r.text[:200000], encoding="utf-8")
            ser = parse_response(r.text)
            ser.attrs["raw"] = r.text[:5000]
            if ser.empty and len(r.text.strip()) > 20 and _warned[0] < 3:
                _warned[0] += 1
                log.warning("Reply had content but 0 points parsed (client=%s ch=%s). Start of reply: %r",
                            client_id, channel_id, r.text[:300])
            return ser
        except Exception as e:               # network / HTTP / parse
            last_err = e
            time.sleep(1.5 * attempt)
    raise RuntimeError(f"API failed for client={client_id} channel={channel_id}: {last_err}")


def demo_series(client_id, channel_id, start, end, step=15):
    """Synthetic data so the pipeline can be tested without the API."""
    rng = np.random.default_rng(client_id * 1000 + channel_id + start.toordinal())
    idx = pd.date_range(start, end, freq=f"{step}min", inclusive="left")
    if channel_id % 7 == 0:
        return pd.Series(dtype=float)                                  # no data
    base = 100 + 40 * np.sin((idx.hour + idx.minute / 60 - 6) / 24 * 2 * np.pi)
    base *= np.where(idx.dayofweek >= 5, 0.6, 1.0)
    s = pd.Series(base + rng.normal(0, 5, len(idx)), index=idx)
    if channel_id % 5 == 0:
        s = s[s.index < pd.Timestamp(end) - pd.Timedelta(days=12)]                      # stale feed
    if channel_id % 11 == 0:
        s.iloc[len(s) // 2 : len(s) // 2 + 24 * 60 // step] *= 4                    # spike day
    return s


# =============================================================================
# Evaluation (current window only)
# =============================================================================
def _dur(minutes):
    d, rem = divmod(int(minutes), 1440)
    h, m = divmod(rem, 60)
    return " ".join(x for x in (f"{d}d" if d else "", f"{h}h" if h else "", f"{m}m" if m else "") if x) or "0m"


def find_gaps(clean, start, last, min_minutes, step=15):
    """Missing stretches (grid = step minutes) between window start and last received value.
    Returns list of (first_missing, last_missing, minutes_missing)."""
    freq = f"{step}min"
    have = pd.DatetimeIndex(clean.index.floor(freq)).unique()
    full = pd.date_range(pd.Timestamp(start), pd.Timestamp(last), freq=freq)
    missing = full.difference(have)
    if len(missing) == 0:
        return []
    m = missing.to_series()
    grp = (m.diff() != pd.Timedelta(minutes=step)).cumsum()
    gaps = []
    for _, g in m.groupby(grp):
        if len(g) * step >= min_minutes:
            gaps.append((g.iloc[0], g.iloc[-1], len(g) * step))
    return gaps


def fmt_gaps(gaps, rules):
    if not gaps:
        return ""
    top = sorted(gaps, key=lambda g: g[2], reverse=True)[: rules["gap_max_listed"]]
    top.sort(key=lambda g: g[0])
    items = [f"{a:%d-%m-%Y %H:%M} to {b:%d-%m-%Y %H:%M} ({_dur(mins)})" for a, b, mins in top]
    more = len(gaps) - len(top)
    return "Data gaps: " + "; ".join(items) + (f" (+{more} more)" if more > 0 else "") + "."


def _dump_raw(raw, tag):
    try:
        d = Path("debug_nodata")
        d.mkdir(exist_ok=True)
        (d / f"{tag}.txt").write_text(raw or "", encoding="utf-8")
    except OSError:
        pass


def evaluate_window(s, start, end, rules, today, tag="x", step=15, offset=None):
    """Availability check of the window -> dict(status, reason, last, clean)."""
    nodata = lambda why: {"status": BAD, "reason": why, "last": None, "clean": pd.Series(dtype=float)}
    if s is None or len(s) == 0:
        a = getattr(s, "attrs", {}) or {}
        if a.get("unparsed"):
            _dump_raw(a.get("raw", ""), tag)
            return nodata(f"No data - API reply could not be read by script (raw saved: debug_nodata/{tag}.txt)")
        return nodata("No data")
    if s.notna().sum() == 0:
        _dump_raw(s.attrs.get("raw", ""), tag)
        return nodata(f"No data - {len(s)} timestamps received but all values empty/unreadable "
                      f"(raw saved: debug_nodata/{tag}.txt)")
    clean = s.dropna()
    first_all, last_all = clean.index.min(), clean.index.max()
    clean = clean[(clean.index >= pd.Timestamp(start)) & (clean.index < pd.Timestamp(end))]
    if clean.empty:
        return nodata(f"No data in window - API data only from {fmt_d(first_all)} to {fmt_d(last_all)}")

    last = clean.index.max()
    expected_end = min(pd.Timestamp(end), pd.Timestamp(today) + (offset or pd.Timedelta(0)))
    # coverage = gaps inside the received range (lateness is handled by max_lag_days)
    expected = max(int((last - pd.Timestamp(start)) / pd.Timedelta(minutes=step)) + 1, 1)
    coverage = len(clean) / expected * 100
    lag_days = (expected_end - last).total_seconds() / 86400

    gtxt = fmt_gaps(find_gaps(clean, start, last, rules["gap_min_minutes"], step), rules)
    if lag_days > rules["max_lag_days"]:
        return {"status": BAD, "last": last, "clean": clean, "gaps": gtxt,
                "reason": (f"load data received till {fmt_d(last)}. Data missing for last {lag_days:.0f} days. "
                           + gtxt).strip()}
    cov_note = ""
    if coverage < rules["min_coverage_pct"]:       # gaps in the data: status stays √, reason is shown in red
        cov_note = f"Only {coverage:.0f}% of intervals available."
    return {"status": OK, "reason": "", "last": last, "clean": clean, "gaps": gtxt, "cov_note": cov_note}


def pattern_notes(clean, rules, offset=None):
    """High/low pick detection inside the window."""
    if offset is not None:                       # gas day: shift so the day runs 06:00 -> 06:00
        clean = clean.copy()
        clean.index = clean.index - offset
    daily_max = clean.resample("D").max().dropna()
    daily_mean = clean.resample("D").mean().dropna()
    notes = []
    if len(daily_max) >= 7:
        q1, q3 = daily_max.quantile([0.25, 0.75])
        limit = q3 + rules["spike_iqr_factor"] * (q3 - q1)
        spikes = daily_max[daily_max > limit]
        if len(spikes):
            notes.append("high pick found on date " + ", ".join(fmt_d(d) for d in spikes.index[:3]))
        lows = daily_mean[daily_mean < daily_mean.median() * rules["low_level_pct"] / 100]
        lows = lows[lows.index.dayofweek < 5]                              # ignore weekends
        if len(lows):
            notes.append("data range is low on " + ", ".join(fmt_d(d) for d in lows.index[:3]))
    return notes


def is_gas(ch):
    return str(ch.get("commodity", "")).strip().lower() == "gas" or \
           str(ch.get("meter_name", "")).strip().lower().startswith("gas")


def check_channel(cfg, client_id, ch, start, end, today, demo=False, debug_path=None):
    rules = cfg["rules"]
    gas = is_gas(ch)
    step = rules["gas_interval_minutes"] if gas else 15
    offset = pd.Timedelta(hours=rules["gas_day_start_hour"]) if gas else pd.Timedelta(0)
    w_start, w_end = pd.Timestamp(start) + offset, pd.Timestamp(end) + offset   # gas day: 06:00 -> 06:00
    fetch_end = end + timedelta(days=1) if gas else end    # gas day ends 05:00 of the next calendar day
    try:
        if demo:
            s = demo_series(client_id, ch["channel"], start, fetch_end, step)
        else:
            s = fetch_series(cfg, client_id, ch["channel"], start, fetch_end, debug_path=debug_path)
    except Exception as e:
        log.warning("%s", e)
        return {"cur": BAD, "reason": f"API error: {e}"[:200], "red": False}

    ev = evaluate_window(s, w_start, w_end, rules, today, tag=f"{client_id}_{ch['channel']}",
                         step=step, offset=offset)
    reason = ev["reason"]
    if ev["status"] == OK:
        parts = [f"load data received till {fmt_d(ev['last'])}."]
        if ev.get("cov_note"):
            parts.append(ev["cov_note"])
        if ev.get("gaps"):
            parts.append(ev["gaps"])
        pk = pattern_notes(ev["clean"], rules, offset if gas else None)
        parts.append("; ".join(pk) + "." if pk else "no high /low pick found. Its annual data pattern.")
        reason = " ".join(parts)
    red = bool(ev.get("gaps") or ev.get("cov_note"))
    return {"cur": ev["status"], "reason": reason, "red": red}


# =============================================================================
# Excel output
# =============================================================================
def write_excel(rows, path, title, prev_date, cur_date):
    """rows: list of (client_name, [ (meter_name, [ (channel_name, prev, cur, reason), ...]) ])"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    thin = Side(style="thin", color="000000")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    outer_med = Side(style="medium", color="000000")

    headers = ["Client", "Meter Name [eeC]", "Channel Name [eeC]", "Previous Status",
               "Current Status", "Reason", "Previous checked date", "Current checked date"]
    widths = [22, 52, 46, 15, 15, 95, 22, 22]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells("A1:H1")
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=12)
    ws["A1"].fill = PatternFill("solid", start_color="4EA72E")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = Font(bold=True, size=9)
        cell.border = box
    ws.freeze_panes = "A4"

    r = 4
    red = Font(color="FF0000", bold=True)
    for client_name, meters in rows:
        first = r
        for meter_name, chans in meters:
            m_first = r
            for ch_name, prev, cur, reason, red_reason in chans:
                ws.cell(r, 3, ch_name)
                for col, st in ((4, prev), (5, cur)):
                    c = ws.cell(r, col, st)
                    c.alignment = Alignment(horizontal="left", vertical="center")
                    if st == BAD:
                        c.font = red
                ws.cell(r, 6, reason).alignment = Alignment(vertical="center", wrap_text=False)
                if red_reason:                                  # gaps found -> comment in red font
                    ws.cell(r, 6).font = Font(color="FF0000")
                ws.cell(r, 7, prev_date)
                ws.cell(r, 8, cur_date)
                for col in range(1, 9):
                    ws.cell(r, col).border = box
                    if col in (3, 6, 7, 8):
                        ws.cell(r, col).alignment = Alignment(vertical="center")
                r += 1
            ws.cell(m_first, 2, meter_name)
            if r - 1 > m_first:
                ws.merge_cells(start_row=m_first, start_column=2, end_row=r - 1, end_column=2)
            ws.cell(m_first, 2).alignment = Alignment(vertical="center", wrap_text=True)
        ws.cell(first, 1, client_name)
        if r - 1 > first:
            ws.merge_cells(start_row=first, start_column=1, end_row=r - 1, end_column=1)
        ws.cell(first, 1).alignment = Alignment(vertical="center")
        # medium outline around client block
        for rr in range(first, r):
            for cc in range(1, 9):
                b = ws.cell(rr, cc).border
                ws.cell(rr, cc).border = Border(
                    left=outer_med if cc == 1 else b.left,
                    right=outer_med if cc == 8 else b.right,
                    top=outer_med if rr == first else b.top,
                    bottom=outer_med if rr == r - 1 else b.bottom)
        r += 1                                   # blank spacer row

    for row in ws.iter_rows(min_row=4, max_row=r):
        for c in row:
            if c.font and not c.font.bold:
                c.font = Font(size=9, color=c.font.color)
    wb.save(path)


# =============================================================================
# Main
# =============================================================================
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-c", "--config", default="config.json")
    ap.add_argument("--client", type=int, action="append", help="only this client id (repeatable)")
    ap.add_argument("--demo", action="store_true", help="use synthetic data instead of the API")
    ap.add_argument("--debug", action="store_true", help="save first raw API reply to debug_raw.txt")
    ap.add_argument("--probe", nargs=2, type=int, metavar=("CLIENT", "CHANNEL"),
                    help="fetch ONE channel, print raw reply start + parsed result, then exit")
    ap.add_argument("--prev-report", help="older report .xlsx to take Previous Status from (first run)")
    ap.add_argument("--no-save", action="store_true", help="do not update status_history.json")
    ap.add_argument("-o", "--output", help="output .xlsx path (default: <output_dir>/Load_Status_<date>.xlsx)")
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    cfg = load_config(args.config)
    today = date.today()
    out_dir = Path(cfg.get("output_dir", "reports"))
    out_dir.mkdir(parents=True, exist_ok=True)
    hist_path = out_dir / "status_history.json"
    history = load_history(hist_path)

    prev, cur = resolve_dates(cfg, history, today)
    start, end = prev, cur + timedelta(days=1)                # end exclusive
    log.info("Data window: %s -> %s (previous check -> current check)", prev, cur)

    if args.probe:
        cl, chn = args.probe
        api = cfg["api"]
        r = requests.get(api["url"], params={"token": api["token"], "fr": api.get("resolution", "min15"),
                         "fillValue": "", "to": end.isoformat(), "from": start.isoformat(),
                         "channel": chn, "client": cl}, timeout=60)
        print("HTTP", r.status_code, "| bytes:", len(r.text), "| content-type:", r.headers.get("content-type"))
        print("RAW START:\n", r.text[:1500])
        ser = parse_response(r.text)
        print("\nPARSED points:", len(ser), "| non-null:", int(ser.notna().sum()))
        if len(ser):
            print(ser.head(5), "\n...\n", ser.tail(3))
        return

    # Previous status source
    prev_report = args.prev_report or cfg.get("report", {}).get("previous_report")
    if prev_report and prev.isoformat() not in history:
        history[prev.isoformat()] = read_report_statuses(prev_report)
        log.info("Previous statuses loaded from %s (%d channels)", prev_report, len(history[prev.isoformat()]))
    prev_status = previous_statuses(history, prev)
    if not prev_status:
        log.warning("No previous statuses found -> Previous Status will show '%s'. "
                    "Use --prev-report <old report.xlsx> to seed it.", NONE)

    jobs = []
    for c in cfg["clients"]:
        if not c.get("enabled", True) or (args.client and c["client"] not in args.client):
            continue
        for ch in c["channels"]:
            if ch.get("enabled", True):
                jobs.append((c, ch))
    log.info("Checking %d channels ...", len(jobs))

    results = {}
    workers = 1 if args.debug else cfg["rules"]["workers"]
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {}
        for i, (c, ch) in enumerate(jobs):
            dbg = "debug_raw.txt" if (args.debug and i == 0) else None
            futs[ex.submit(check_channel, cfg, c["client"], ch, start, end, today, args.demo, dbg)] = (c["client"], ch["channel"])
        done = 0
        for f in as_completed(futs):
            results[futs[f]] = f.result()
            done += 1
            if done % 25 == 0 or done == len(futs):
                log.info("  %d/%d done", done, len(futs))

    rows, new_status = [], {}
    for c in cfg["clients"]:
        meters = {}
        for ch in c["channels"]:
            res = results.get((c["client"], ch["channel"]))
            if not res:
                continue
            k = key_of(ch["meter_name"], ch["channel_name"])
            new_status[k] = res["cur"]
            meters.setdefault(ch["meter_name"], []).append(
                (ch["channel_name"], prev_status.get(k, NONE), res["cur"], res["reason"],
                 res.get("red", False)))
        if meters:
            rows.append((c["name"], list(meters.items())))

    out = Path(args.output) if args.output else out_dir / f"Load_Status_{cur:%Y-%m-%d}.xlsx"
    write_excel(rows, out, f"{fmt_d(prev)} to {fmt_d(cur)}", fmt_d(prev), fmt_d(cur))

    all_bad = results and all(r["cur"] == BAD for r in results.values())
    if all_bad:
        log.warning("EVERY channel is %s - looks like a failure, so status_history.json was NOT updated.", BAD)
    if not args.demo and not args.no_save and not args.client and not all_bad:
        history[cur.isoformat()] = new_status
        hist_path.write_text(json.dumps(history, indent=1, ensure_ascii=False), encoding="utf-8")
        log.info("Statuses saved to %s", hist_path)
    elif not args.demo and not args.no_save and not all_bad:
        log.info("--client used: history not updated (partial run).")

    bad = sum(1 for r in results.values() if r["cur"] == BAD)
    log.info("Report saved: %s  (%d channels, %d currently %s)", out, len(results), bad, BAD)


if __name__ == "__main__":
    sys.exit(main())