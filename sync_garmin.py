"""
Garmin -> Google Sheets sync for the Sportdata workbook.

What it writes
  1. "Garmin Data"   one row per run (as before), plus new per-run columns:
                     Aerobic TE, Anaerobic TE, Training Load, Avg Power (W),
                     VO2max, Start Time, Activity ID, running dynamics,
                     minutes per heart-rate zone and the weather during the run.
                     Columns are matched by header name, so your manual
                     "Run type" column and any column order are left alone.
                     Missing new columns are added to the header automatically,
                     and recent existing rows are back-filled.
  2. "Daily Metrics" one row per day (created on first run): resting HR, HRV,
                     sleep, Body Battery, stress, Training Readiness, VO2max,
                     training status, load and load focus, lactate threshold HR
                     and pace, and Garmin's race predictions.

Optional environment variables
  DAILY_DAYS      how many past days to refresh each run (default 3; recent days
                  are refreshed because sleep/HRV arrive after the night ends)
  DAILY_BACKFILL  how many days to fill the first time the tab is empty (default 30)
"""

import os
import json
import time
from datetime import date, datetime, timedelta

from garminconnect import Garmin
from google.oauth2.service_account import Credentials
import gspread
from gspread.utils import rowcol_to_a1

# Load environment variables from .env file if it exists (for local testing)
if os.path.exists('.env'):
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        print("Warning: python-dotenv not installed. Install with: pip install python-dotenv")


# --------------------------------------------------------------------------
# Sheet layout
# --------------------------------------------------------------------------

RUN_SHEET = "Garmin Data"
DAILY_SHEET = "Daily Metrics"

# Original columns, in the original order (used only if the sheet has no header yet).
BASE_RUN_HEADERS = [
    "Date", "Activity Name", "Distance (km)", "Duration (min)", "Avg Pace (min/km)",
    "Avg HR", "Max HR", "Calories", "Avg Cadence", "Elevation Gain (m)",
    "Activity Type", "Lap Details", "Run type",
]
# New per-run columns. Added to the right of whatever is already there.
NEW_RUN_HEADERS = [
    "Aerobic TE", "Anaerobic TE", "Training Load", "Avg Power (W)",
    "VO2max", "Start Time", "Activity ID",
    "Ground Contact (ms)", "Vertical Osc (cm)", "Stride (m)", "Vertical Ratio (%)",
]
# These need two extra Garmin requests per run (heart-rate zones and weather).
DETAIL_RUN_HEADERS = [
    "Z1 (min)", "Z2 (min)", "Z3 (min)", "Z4 (min)", "Z5 (min)",
    "Temp (C)", "Feels Like (C)", "Humidity (%)", "Wind (km/h)", "Weather",
]
MAX_DETAIL_BACKFILL = 60   # at most this many existing runs get zones/weather per sync

DAILY_HEADERS = [
    "Date",
    "Resting HR", "HRV Last Night (ms)", "HRV Weekly Avg (ms)", "HRV Status",
    "Sleep Score", "Sleep (h)",
    "Body Battery Wake", "Body Battery High", "Body Battery Low", "Avg Stress",
    "Training Readiness", "Readiness Level",
    "VO2max", "Training Status", "Acute Load", "Chronic Load",
    "LT HR (bpm)", "LT Pace (min/km)",
    "Pred 5K (s)", "Pred 10K (s)", "Pred HM (s)", "Pred Marathon (s)",
    "Load Low Aerobic", "Load High Aerobic", "Load Anaerobic", "Load Focus", "Recovery Time (h)",
]


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------

def format_duration(seconds):
    """Convert seconds to minutes (rounded to 2 decimals)"""
    return round(seconds / 60, 2) if seconds else 0


def format_pace(distance_meters, duration_seconds):
    """Calculate pace in min/km"""
    if not distance_meters or not duration_seconds:
        return 0
    distance_km = distance_meters / 1000
    pace_seconds = duration_seconds / distance_km
    return round(pace_seconds / 60, 2)  # Convert to min/km


def format_pace_str(distance_meters, duration_seconds):
    """Calculate pace as a M:SS string (easier to read in a JSON blob than a decimal)"""
    if not distance_meters or not duration_seconds:
        return None
    distance_km = distance_meters / 1000
    pace_seconds = duration_seconds / distance_km
    m = int(pace_seconds // 60)
    s = int(round(pace_seconds % 60))
    if s == 60:
        m += 1
        s = 0
    return f"{m}:{s:02d}"


def dig(obj, *path, default=None):
    """Safely walk nested dicts/lists: dig(d, 'a', 0, 'b'). Returns default on any miss."""
    cur = obj
    for key in path:
        try:
            if isinstance(cur, dict):
                cur = cur.get(key)
            elif isinstance(cur, (list, tuple)) and isinstance(key, int):
                cur = cur[key]
            else:
                return default
        except (IndexError, KeyError, TypeError):
            return default
        if cur is None:
            return default
    return cur


def rnd(value, digits=1):
    """Round numbers, pass through None as an empty cell."""
    if value is None or value == "":
        return ""
    try:
        return round(float(value), digits)
    except (TypeError, ValueError):
        return ""


def safe(label, fn, *args, **kwargs):
    """Call a Garmin endpoint; one failing metric never blocks the rest."""
    try:
        return fn(*args, **kwargs)
    except Exception as e:  # noqa: BLE001 - Garmin raises many different errors
        print(f"  ⚠️ {label}: {e}")
        return None


def to_number(text):
    """Parse a sheet cell that may use a decimal comma (Dutch locale)."""
    try:
        return float(str(text).replace(",", "."))
    except (TypeError, ValueError):
        return None


# --------------------------------------------------------------------------
# Runs
# --------------------------------------------------------------------------

def get_laps_json(garmin, activity_id):
    """
    Fetch per-lap data for an activity and return it as a JSON string.
    Each lap reflects a manual lap-button press (or auto-lap if enabled),
    so a hill rep / tempo segment / recovery jog show up as distinct rows
    instead of being blended into a single whole-run average.
    Returns '[]' on any failure so a single bad activity never blocks the sync.
    """
    try:
        splits = garmin.get_activity_splits(activity_id)
        laps = splits.get('lapDTOs', []) if splits else []
        lap_list = []
        for i, lap in enumerate(laps, start=1):
            dist_m = lap.get('distance', 0) or 0
            dur_s = lap.get('duration', 0) or 0
            lap_list.append({
                "lap": i,
                "distance_km": round(dist_m / 1000, 3) if dist_m else 0,
                "duration_s": round(dur_s, 1) if dur_s else 0,
                "avg_pace": format_pace_str(dist_m, dur_s),
                "avg_hr": lap.get('averageHR'),
                "max_hr": lap.get('maxHR'),
                "elevation_gain_m": round(lap.get('elevationGain', 0), 1) if lap.get('elevationGain') else 0,
                "cadence": round(lap['averageRunCadence']) if lap.get('averageRunCadence') else None,
                "power": round(lap['averagePower']) if lap.get('averagePower') else None,
            })
        return json.dumps(lap_list)
    except Exception as e:
        print(f"  ⚠️ Could not fetch laps for activity {activity_id}: {e}")
        return "[]"


def run_values(activity):
    """All per-run values that come straight from the activity list (no extra API call)."""
    distance_m = activity.get('distance', 0) or 0
    duration_s = activity.get('duration', 0) or 0
    start_local = activity.get('startTimeLocal', '') or ''
    return {
        "Date": start_local[:10],
        "Activity Name": activity.get('activityName', 'Run'),
        "Distance (km)": round(distance_m / 1000, 2) if distance_m else 0,
        "Duration (min)": format_duration(duration_s),
        "Avg Pace (min/km)": format_pace(distance_m, duration_s),
        "Avg HR": activity.get('averageHR', 0) or 0,
        "Max HR": activity.get('maxHR', 0) or 0,
        "Calories": activity.get('calories', 0) or 0,
        "Avg Cadence": activity.get('averageRunningCadenceInStepsPerMinute', 0) or 0,
        "Elevation Gain (m)": round(activity.get('elevationGain', 0), 1) if activity.get('elevationGain') else 0,
        "Activity Type": dig(activity, 'activityType', 'typeKey', default='running'),
        # New columns
        "Aerobic TE": rnd(activity.get('aerobicTrainingEffect'), 1),
        "Anaerobic TE": rnd(activity.get('anaerobicTrainingEffect'), 1),
        "Training Load": rnd(activity.get('activityTrainingLoad'), 0),
        "Avg Power (W)": rnd(activity.get('avgPower'), 0),
        "VO2max": rnd(activity.get('vO2MaxValue'), 1),
        "Start Time": start_local[11:16],
        "Activity ID": str(activity.get('activityId') or ''),
        "Ground Contact (ms)": rnd(activity.get('avgGroundContactTime'), 0),
        "Vertical Osc (cm)": rnd(activity.get('avgVerticalOscillation'), 1),
        "Stride (m)": rnd((activity.get('avgStrideLength') or 0) / 100, 2) if activity.get('avgStrideLength') else "",
        "Vertical Ratio (%)": rnd(activity.get('avgVerticalRatio'), 1),
    }


def detail_values(garmin, activity_id):
    """Heart-rate zones and weather for one run (two extra Garmin requests)."""
    out = {name: "" for name in DETAIL_RUN_HEADERS}
    zones = safe(f"hr zones {activity_id}", garmin.get_activity_hr_in_timezones, activity_id) or []
    for z in zones if isinstance(zones, list) else []:
        n = z.get('zoneNumber')
        if n in (1, 2, 3, 4, 5) and z.get('secsInZone') is not None:
            out[f"Z{n} (min)"] = round(z['secsInZone'] / 60, 1)
    w = safe(f"weather {activity_id}", garmin.get_activity_weather, activity_id) or {}
    if isinstance(w, dict):
        f_to_c = lambda f: round((f - 32) * 5 / 9, 1)          # Garmin reports Fahrenheit
        if w.get('temp') is not None:
            out["Temp (C)"] = f_to_c(w['temp'])
        if w.get('apparentTemp') is not None:
            out["Feels Like (C)"] = f_to_c(w['apparentTemp'])
        if w.get('relativeHumidity') is not None:
            out["Humidity (%)"] = w['relativeHumidity']
        if w.get('windSpeed') is not None:
            out["Wind (km/h)"] = round(w['windSpeed'] * 1.609, 0)   # Garmin reports mph
        out["Weather"] = dig(w, 'weatherTypeDTO', 'desc', default="") or ""
    time.sleep(0.3)
    return out


def ensure_headers(sheet, wanted, base=None):
    """Return the header row, adding any missing column names to the right."""
    header = sheet.row_values(1)
    if not header:
        header = list(base or wanted)
        sheet.update(range_name="A1", values=[header])
    missing = [h for h in wanted if h not in header]
    if missing:
        first_col = len(header) + 1
        needed_cols = len(header) + len(missing)
        if sheet.col_count < needed_cols:
            sheet.add_cols(needed_cols - sheet.col_count)
        sheet.update(range_name=rowcol_to_a1(1, first_col), values=[missing])
        header = header + missing
        print(f"  Added columns to '{sheet.title}': {', '.join(missing)}")
    return header


def sync_runs(garmin, spreadsheet):
    print("\nFetching recent activities...")
    try:
        activities = garmin.get_activities(0, 50)  # Get last 50 activities
        print(f"Found {len(activities)} total activities")
    except Exception as e:
        print(f"❌ Failed to fetch activities: {e}")
        return

    running = [
        a for a in activities
        if dig(a, 'activityType', 'typeKey', default='').lower() in ['running', 'treadmill_running', 'trail_running']
    ]
    print(f"Found {len(running)} running activities")
    if not running:
        return

    try:
        sheet = spreadsheet.worksheet(RUN_SHEET)
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.sheet1

    header = ensure_headers(sheet, BASE_RUN_HEADERS[:12] + NEW_RUN_HEADERS + DETAIL_RUN_HEADERS,
                            base=BASE_RUN_HEADERS + NEW_RUN_HEADERS + DETAIL_RUN_HEADERS)
    col = {name: i for i, name in enumerate(header)}  # 0-based

    rows = sheet.get_all_values()[1:]

    def cell(row, name):
        i = col.get(name)
        return row[i] if i is not None and i < len(row) else ""

    existing_ids = {cell(r, "Activity ID") for r in rows if cell(r, "Activity ID")}
    print(f"Found {len(rows)} existing entries")
    matched_rows = set()

    # ---- back-fill the new columns on existing rows (matched by Activity ID, else date + distance)
    by_id = {str(a.get('activityId')): a for a in running if a.get('activityId')}
    updates, detail_budget = [], MAX_DETAIL_BACKFILL
    for row_number, r in enumerate(rows, start=2):
        r_id = cell(r, "Activity ID")
        if r_id:
            match = by_id.get(r_id)
        else:
            r_date, r_km = cell(r, "Date"), to_number(cell(r, "Distance (km)"))
            match = next((a for a in running
                          if (a.get('startTimeLocal', '') or '')[:10] == r_date
                          and r_km is not None
                          and abs(round((a.get('distance', 0) or 0) / 1000, 2) - r_km) <= 0.03), None)
        if not match:
            continue
        values = run_values(match)
        if detail_budget > 0 and not any(cell(r, name) for name in DETAIL_RUN_HEADERS):
            values.update(detail_values(garmin, match.get('activityId')))
            detail_budget -= 1
        for name in NEW_RUN_HEADERS + DETAIL_RUN_HEADERS:
            if not cell(r, name) and values.get(name, "") != "":
                updates.append({"range": rowcol_to_a1(row_number, col[name] + 1), "values": [[values[name]]]})
        existing_ids.add(values["Activity ID"])
        matched_rows.add(row_number)
    if updates:
        sheet.batch_update(updates, value_input_option="RAW")
        print(f"  Back-filled {len(updates)} cells on existing rows")

    # Rows from the old script that could not be matched have no Activity ID: fall back to the date for those.
    dates_without_id = {cell(r, "Date") for n, r in enumerate(rows, start=2)
                        if cell(r, "Date") and not cell(r, "Activity ID") and n not in matched_rows}

    # ---- append new runs (oldest first, so the sheet stays in date order)
    new_rows = []
    for activity in sorted(running, key=lambda a: a.get('startTimeLocal', '') or ''):
        try:
            values = run_values(activity)
            if values["Activity ID"] in existing_ids:
                continue
            if values["Date"] in dates_without_id:
                # an old-style row for this date that could not be matched by distance
                print(f"Skipping {values['Date']} - already exists")
                continue
            activity_id = activity.get('activityId')
            values["Lap Details"] = get_laps_json(garmin, activity_id) if activity_id else "[]"
            if activity_id:
                values.update(detail_values(garmin, activity_id))
            row = [values.get(name, "") for name in header]  # "Run type" and unknown columns stay empty
            new_rows.append(row)
            n_laps = len(json.loads(values["Lap Details"]))
            print(f"✅ New: {values['Date']} - {values['Activity Name']} ({values['Distance (km)']} km, {n_laps} laps)")
        except Exception as e:
            print(f"❌ Error processing activity: {e}")

    if new_rows:
        sheet.append_rows(new_rows, value_input_option="RAW", table_range="A1")
        print(f"\n🎉 Successfully added {len(new_rows)} new running activities!")
    else:
        print("\n✓ No new activities to add")


# --------------------------------------------------------------------------
# Daily metrics
# --------------------------------------------------------------------------

def lactate_threshold(garmin):
    """Latest lactate threshold as (heart rate, pace in min/km)."""
    data = safe("lactate threshold", garmin.get_lactate_threshold, latest=True)
    hr = dig(data, 'speed_and_heart_rate', 'heartRate')
    speed = dig(data, 'speed_and_heart_rate', 'speed')
    pace = ""
    if speed:
        speed = float(speed)
        if speed < 1:          # Garmin stores this as metres per second divided by 10
            speed *= 10
        if speed > 0:
            pace = round(1000 / speed / 60, 2)
    return (hr or ""), pace


def race_predictions(garmin):
    data = safe("race predictions", garmin.get_race_predictions)
    if isinstance(data, list):
        data = data[-1] if data else {}
    data = data or {}
    return [data.get(k, "") or "" for k in ("time5K", "time10K", "timeHalfMarathon", "timeMarathon")]


def training_status(garmin, day):
    """Training status, acute/chronic load, VO2max and load focus for a day."""
    data = safe(f"training status {day}", garmin.get_training_status, day) or {}
    out = {"status": "", "acute": "", "chronic": "", "low": "", "high": "", "anaerobic": "", "focus": "",
           "vo2": dig(data, 'mostRecentVO2Max', 'generic', 'vo2MaxPreciseValue') or dig(data, 'mostRecentVO2Max', 'generic', 'vo2MaxValue')}
    per_device = dig(data, 'mostRecentTrainingStatus', 'latestTrainingStatusData', default={}) or {}
    for dev in per_device.values():
        if not isinstance(dev, dict):
            continue
        out["status"] = dev.get('trainingStatusFeedbackPhrase') or out["status"]
        out["acute"] = dig(dev, 'acuteTrainingLoadDTO', 'dailyTrainingLoadAcute', default=out["acute"])
        out["chronic"] = dig(dev, 'acuteTrainingLoadDTO', 'dailyTrainingLoadChronic', default=out["chronic"])
        if dev.get('primaryTrainingDevice'):
            break
    balance = dig(data, 'mostRecentTrainingLoadBalance', 'metricsTrainingLoadBalanceDTOMap', default={}) or {}
    for dev in balance.values():
        if not isinstance(dev, dict):
            continue
        out["low"] = dev.get('monthlyLoadAerobicLow', out["low"])
        out["high"] = dev.get('monthlyLoadAerobicHigh', out["high"])
        out["anaerobic"] = dev.get('monthlyLoadAnaerobic', out["anaerobic"])
        out["focus"] = dev.get('trainingBalanceFeedbackPhrase') or out["focus"]
        if dev.get('primaryTrainingDevice'):
            break
    return out


def daily_row(garmin, day, lt, preds, is_today):
    """Collect one day's metrics. `day` is 'YYYY-MM-DD'."""
    stats = safe(f"stats {day}", garmin.get_stats, day) or {}
    sleep = safe(f"sleep {day}", garmin.get_sleep_data, day) or {}
    hrv = safe(f"hrv {day}", garmin.get_hrv_data, day) or {}
    readiness = safe(f"readiness {day}", garmin.get_training_readiness, day) or []
    ts = training_status(garmin, day)
    status, acute, chronic, vo2 = ts["status"], ts["acute"], ts["chronic"], ts["vo2"]
    if not vo2:
        mm = safe(f"max metrics {day}", garmin.get_max_metrics, day)
        vo2 = dig(mm, 0, 'generic', 'vo2MaxPreciseValue') or dig(mm, 'generic', 'vo2MaxPreciseValue')

    sleep_seconds = dig(sleep, 'dailySleepDTO', 'sleepTimeSeconds')
    # Readiness is a list with the newest reading first; take the first one that has a score.
    ready = next((r for r in readiness if isinstance(r, dict) and r.get('score') is not None), {}) if isinstance(readiness, list) else (readiness or {})

    row = {
        "Date": day,
        "Resting HR": stats.get('restingHeartRate') or dig(sleep, 'restingHeartRate') or "",
        "HRV Last Night (ms)": dig(hrv, 'hrvSummary', 'lastNightAvg') or dig(sleep, 'avgOvernightHrv') or "",
        "HRV Weekly Avg (ms)": dig(hrv, 'hrvSummary', 'weeklyAvg') or "",
        "HRV Status": dig(hrv, 'hrvSummary', 'status') or "",
        "Sleep Score": dig(sleep, 'dailySleepDTO', 'sleepScores', 'overall', 'value') or "",
        "Sleep (h)": round(sleep_seconds / 3600, 2) if sleep_seconds else "",
        "Body Battery Wake": stats.get('bodyBatteryAtWakeTime') or "",
        "Body Battery High": stats.get('bodyBatteryHighestValue') or "",
        "Body Battery Low": stats.get('bodyBatteryLowestValue') or "",
        "Avg Stress": stats.get('averageStressLevel') if (stats.get('averageStressLevel') or 0) > 0 else "",
        "Training Readiness": ready.get('score', "") if ready else "",
        "Readiness Level": ready.get('level', "") if ready else "",
        "VO2max": rnd(vo2, 1),
        "Training Status": status,
        "Acute Load": rnd(acute, 0),
        "Chronic Load": rnd(chronic, 0),
        # Threshold and predictions are "latest" values, so they are only stamped on today's row.
        "LT HR (bpm)": lt[0] if is_today else "",
        "LT Pace (min/km)": lt[1] if is_today else "",
        "Pred 5K (s)": preds[0] if is_today else "",
        "Pred 10K (s)": preds[1] if is_today else "",
        "Pred HM (s)": preds[2] if is_today else "",
        "Pred Marathon (s)": preds[3] if is_today else "",
        "Load Low Aerobic": rnd(ts["low"], 0),
        "Load High Aerobic": rnd(ts["high"], 0),
        "Load Anaerobic": rnd(ts["anaerobic"], 0),
        "Load Focus": ts["focus"],
        # Garmin gives recovery time in minutes, on watches that report it.
        "Recovery Time (h)": round(ready['recoveryTime'] / 60, 1) if ready and ready.get('recoveryTime') else "",
    }
    return row


def sync_daily(garmin, spreadsheet):
    print("\nSyncing daily metrics...")
    try:
        sheet = spreadsheet.worksheet(DAILY_SHEET)
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=DAILY_SHEET, rows=400, cols=len(DAILY_HEADERS))
        print(f"  Created tab '{DAILY_SHEET}'")

    header = ensure_headers(sheet, DAILY_HEADERS)
    col = {name: i for i, name in enumerate(header)}
    rows = sheet.get_all_values()[1:]
    row_of = {r[col["Date"]]: n for n, r in enumerate(rows, start=2) if r and len(r) > col["Date"] and r[col["Date"]]}

    days_back = int(os.environ.get("DAILY_DAYS", "3"))
    if not row_of:
        days_back = int(os.environ.get("DAILY_BACKFILL", "30"))
        print(f"  Empty tab: back-filling {days_back} days")

    today = date.today()
    lt = lactate_threshold(garmin)
    preds = race_predictions(garmin)

    updates, appends = [], []
    for offset in range(days_back - 1, -1, -1):          # oldest first
        day = (today - timedelta(days=offset)).isoformat()
        values = daily_row(garmin, day, lt, preds, is_today=(offset == 0))
        if day in row_of:
            existing = rows[row_of[day] - 2]
            line = []
            for i, name in enumerate(header):
                new = values.get(name, "")
                old = existing[i] if i < len(existing) else ""
                line.append(new if new != "" else old)       # never blank out a value we already had
            updates.append({"range": rowcol_to_a1(row_of[day], 1), "values": [line]})
        else:
            appends.append([values.get(name, "") for name in header])
        filled = sum(1 for k, v in values.items() if k != "Date" and v != "")
        print(f"  {day}: {filled} metrics")
        time.sleep(0.5)                                      # be gentle with Garmin

    if updates:
        sheet.batch_update(updates, value_input_option="RAW")
    if appends:
        sheet.append_rows(appends, value_input_option="RAW", table_range="A1")
    print(f"✅ Daily metrics: {len(appends)} new day(s), {len(updates)} refreshed")


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    print("Starting Garmin sync...")

    # Get credentials from environment variables
    garmin_email = os.environ.get('GARMIN_EMAIL')
    garmin_password = os.environ.get('GARMIN_PASSWORD')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    sheet_id = os.environ.get('SHEET_ID')

    # For local testing: try to load from credentials.json file
    if not google_creds_json and os.path.exists('credentials.json'):
        print("Loading Google credentials from credentials.json...")
        with open('credentials.json', 'r') as f:
            google_creds_json = f.read()

    if not all([garmin_email, garmin_password, google_creds_json, sheet_id]):
        print("❌ Missing required environment variables")
        print(f"  GARMIN_EMAIL: {'✓' if garmin_email else '✗'}")
        print(f"  GARMIN_PASSWORD: {'✓' if garmin_password else '✗'}")
        print(f"  GOOGLE_CREDENTIALS: {'✓' if google_creds_json else '✗'}")
        print(f"  SHEET_ID: {'✓' if sheet_id else '✗'}")
        return

    # Connect to Garmin
    print("Connecting to Garmin...")
    try:
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()
        print("✅ Connected to Garmin")
    except Exception as e:
        print(f"❌ Failed to connect to Garmin: {e}")
        return

    # Connect to Google Sheets
    print("Connecting to Google Sheets...")
    try:
        creds_dict = json.loads(google_creds_json)
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=[
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
        )
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(sheet_id)   # open by spreadsheet ID, not by title
        print(f"✅ Connected to Google Sheets: {spreadsheet.title} ({spreadsheet.id})")
    except Exception as e:
        print(f"❌ Failed to connect to Google Sheets: {e}")
        print("  Check that the GOOGLE_CREDENTIALS secret is valid, the service account has Editor access, and SHEET_ID is the spreadsheet ID from the sheet URL.")
        return

    # The two parts are independent: a failure in one never blocks the other.
    try:
        sync_runs(garmin, spreadsheet)
    except Exception as e:
        print(f"❌ Run sync failed: {e}")
    try:
        sync_daily(garmin, spreadsheet)
    except Exception as e:
        print(f"❌ Daily metrics sync failed: {e}")

    print(f"\nDone at {datetime.now().isoformat(timespec='seconds')}")


if __name__ == "__main__":
    main()
