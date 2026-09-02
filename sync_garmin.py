import os
import json
from garminconnect import Garmin
from google.oauth2.service_account import Credentials
import gspread
from datetime import datetime, timedelta

# Load environment variables from .env file if it exists (for local testing)
if os.path.exists('.env'):
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        print("Warning: python-dotenv not installed. Install with: pip install python-dotenv")
        pass


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
            })
        return json.dumps(lap_list)
    except Exception as e:
        print(f"  ⚠️ Could not fetch laps for activity {activity_id}: {e}")
        return "[]"


def main():
    print("Starting Garmin running activities sync...")

    # Get credentials from environment variables
    garmin_email = os.environ.get('GARMIN_EMAIL')
    garmin_password = os.environ.get('GARMIN_PASSWORD')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    sheet_id = os.environ.get('SHEET_ID')  # Add sheet ID from environment

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

    # Get recent activities
    print("Fetching recent activities...")
    try:
        activities = garmin.get_activities(0, 50)  # Get last 50 activities
        print(f"Found {len(activities)} total activities")
    except Exception as e:
        print(f"❌ Failed to fetch activities: {e}")
        return

    # Filter for running activities only
    running_activities = [
        activity for activity in activities
        if activity.get('activityType', {}).get('typeKey', '').lower() in ['running', 'treadmill_running', 'trail_running']
    ]
    print(f"Found {len(running_activities)} running activities")

    if not running_activities:
        print("No running activities found in recent data")
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

        # Open by spreadsheet ID, not by workbook title.
        spreadsheet = client.open_by_key(sheet_id)

        # Prefer the worksheet called "Garmin Data" if it exists.
        try:
            sheet = spreadsheet.worksheet("Garmin Data")
        except gspread.WorksheetNotFound:
            sheet = spreadsheet.sheet1

        print(f"✅ Connected to Google Sheets: {spreadsheet.title} ({spreadsheet.id})")
    except Exception as e:
        print(f"❌ Failed to connect to Google Sheets: {e}")
        print("  Check that the GOOGLE_CREDENTIALS secret is valid, the service account has Editor access, and SHEET_ID is the spreadsheet ID from the sheet URL.")
        return

    # Get existing dates to avoid duplicates
    try:
        existing_data = sheet.get_all_values()
        existing_dates = set()
        if len(existing_data) > 1:
            for row in existing_data[1:]:
                if row and row[0]:
                    existing_dates.add(row[0])
        print(f"Found {len(existing_dates)} existing entries")
    except Exception as e:
        print(f"Warning: Could not check existing data: {e}")
        existing_dates = set()

    # Process each running activity
    new_entries = 0
    for activity in running_activities:
        try:
            activity_date = activity.get('startTimeLocal', '')[:10]

            if activity_date in existing_dates:
                print(f"Skipping {activity_date} - already exists")
                continue

            activity_id = activity.get('activityId')
            activity_name = activity.get('activityName', 'Run')
            distance_meters = activity.get('distance', 0)
            distance_km = round(distance_meters / 1000, 2) if distance_meters else 0
            duration_seconds = activity.get('duration', 0)
            duration_min = format_duration(duration_seconds)
            avg_pace = format_pace(distance_meters, duration_seconds)
            avg_hr = activity.get('averageHR', 0) or 0
            max_hr = activity.get('maxHR', 0) or 0
            calories = activity.get('calories', 0) or 0
            avg_cadence = activity.get('averageRunningCadenceInStepsPerMinute', 0) or 0
            elevation_gain = round(activity.get('elevationGain', 0), 1) if activity.get('elevationGain') else 0
            activity_type = activity.get('activityType', {}).get('typeKey', 'running')

            # NEW: fetch per-lap breakdown so hill reps / tempo segments / recovery
            # jogs show up distinctly instead of being averaged into one row.
            laps_json = get_laps_json(garmin, activity_id) if activity_id else "[]"

            row = [
                activity_date,
                activity_name,
                distance_km,
                duration_min,
                avg_pace,
                avg_hr,
                max_hr,
                calories,
                avg_cadence,
                elevation_gain,
                activity_type,
                laps_json  # NEW: column 12
            ]

            sheet.append_row(row)
            n_laps = len(json.loads(laps_json))
            print(f"✅ Added: {activity_date} - {activity_name} ({distance_km} km, {n_laps} laps)")
            new_entries += 1

        except Exception as e:
            print(f"❌ Error processing activity: {e}")
            continue

    if new_entries > 0:
        print(f"\n🎉 Successfully added {new_entries} new running activities!")
    else:
        print("\n✓ No new activities to add")


if __name__ == "__main__":
    main()
