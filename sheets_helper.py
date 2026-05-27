import os
import re
import json
import gspread
from google.oauth2 import service_account
from gspread_formatting import *
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

# Load environment variables
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Config
SPREADSHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1GEzRkk6SArh_TfvCXw_SgWLyveQLFUMgHDjJIQo1wA4")
GOOGLE_CREDS_FILE = os.getenv("GOOGLE_CREDS_FILE", "linen-rex-436411-r4-9bba0db0c720.json")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

if os.path.isabs(GOOGLE_CREDS_FILE):
    CREDENTIALS_PATH = GOOGLE_CREDS_FILE
else:
    CREDENTIALS_PATH = os.path.join(SCRIPT_DIR, GOOGLE_CREDS_FILE)

# Styling Colors
C_HDR_BG   = "1F3864"   # dark navy
C_HDR_FG   = "FFFFFF"   # white
C_TITLE_BG = "2E75B6"   # slate blue
C_YES_BG   = "E2EFDA"   # soft green
C_YES_FG   = "375623"
C_NO_BG    = "FCE4D6"   # soft pink
C_NO_FG    = "C00000"
C_ALT_BG   = "F9FBFD"   # very light blue alternate
C_BORDER   = "D9D9D9"

def hex_to_color(hex_str):
    """Convert hex string (e.g. 1F3864) to a gspread_formatting Color object."""
    r = int(hex_str[0:2], 16) / 255.0
    g = int(hex_str[2:4], 16) / 255.0
    b = int(hex_str[4:6], 16) / 255.0
    return Color(r, g, b)

def get_google_client():
    """Authenticate and return the gspread client."""
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    
    # 1. Try to load credentials from raw JSON string environment variable (perfect for Render/Heroku)
    creds_json_str = os.getenv("GOOGLE_CREDS_JSON")
    if creds_json_str:
        try:
            creds_info = json.loads(creds_json_str)
            creds = service_account.Credentials.from_service_account_info(
                creds_info,
                scopes=scopes
            )
            return gspread.Client(auth=creds)
        except Exception as e:
            print(f"Error loading credentials from GOOGLE_CREDS_JSON: {e}")
            
    # 2. Fall back to loading from local JSON key file (for local development)
    if os.path.exists(CREDENTIALS_PATH):
        creds = service_account.Credentials.from_service_account_file(
            CREDENTIALS_PATH,
            scopes=scopes
        )
        return gspread.Client(auth=creds)
    else:
        raise FileNotFoundError(
            f"Google Credentials file not found at {CREDENTIALS_PATH}. "
            "Please ensure the file exists locally, or set the GOOGLE_CREDS_JSON environment variable."
        )

def push_to_google_sheet(flat_rows, sis_id, course_name):
    """Pushes audit data to a dedicated worksheet, formatting it to match the Excel style."""
    print("  Connecting to Google Sheets...")
    gc = get_google_client()
    sh = gc.open_by_key(SPREADSHEET_ID)

    # Sanitize sheet title (Excel & sheets compatibility)
    clean_title = re.sub(r'[:\\\/\?\*\[\]]', '_', sis_id)
    safe_title = clean_title[:31]

    try:
        ws = sh.worksheet(safe_title)
        print(f"  Worksheet {safe_title} exists. Recreating worksheet to clear all formats and contents...")
        sh.del_worksheet(ws)
        ws = sh.add_worksheet(title=safe_title, rows="200", cols="20")
    except gspread.exceptions.WorksheetNotFound:
        print(f"  Worksheet {safe_title} not found. Creating...")
        ws = sh.add_worksheet(title=safe_title, rows="200", cols="20")

    # Construct Values Matrix
    values = []

    # Row 1: Title Block
    values.append([f"EDGEWOOD DISCUSSION BOARD AUDIT  |  {course_name} ({sis_id})"] + [""] * 13)

    # Row 2: Headers
    headers = [
        "Cohort", "Course Name", "Topic", "Learner Name", "Student ID", "Learner Email", 
        "Created On", "Replied", "Replied On", "Duration (Hours)", "SLA Status", 
        "First Responder", "Replied By (Name)", "Replied By (Email)"
    ]
    values.append(headers)

    # Row 3 to N+2: Data
    if not flat_rows:
        values.append(["No discussion boards or learner posts found in this course."] + [""] * 13)
        last_row = 3
    else:
        for r in flat_rows:
            row_vals = [
                r.get("Cohort", "N/A"),
                r.get("Course Name", "N/A"),
                r.get("Topic", ""),
                r.get("Learner Name", ""),
                r.get("Student ID", "N/A"),
                r.get("Learner Email", ""),
                r.get("Created On", ""),
                r.get("Replied", "N/A"),
                r.get("Replied On", ""),
                r.get("Duration (Hours)", ""),
                r.get("SLA Status", "N/A"),
                r.get("First Responder", ""),
                r.get("Replied By (Name)", ""),
                r.get("Replied By (Email)", "")
            ]
            values.append(row_vals)
        last_row = len(flat_rows) + 2

    # Spacer
    values.append([""] * 14)
    values.append([""] * 14)

    # Overview calculations
    unique_learners = len(set(r.get("Learner Name") for r in flat_rows if r.get("Learner Name") and r.get("Learner Name") != "(No student posts yet)"))
    total_queries   = len([r for r in flat_rows if r.get("Replied") != "N/A" and r.get("Learner Name") != "(No student posts yet)"])
    total_replied   = len([r for r in flat_rows if r.get("Replied") == "Yes"])
    pending_queries = len([r for r in flat_rows if r.get("Replied") == "No"])
    durations       = [r.get("Duration (Hours)") for r in flat_rows if isinstance(r.get("Duration (Hours)"), (int, float))]
    avg_course_time = round(sum(durations)/len(durations), 2) if durations else 0

    sum_start = last_row + 3

    values.append(["COURSE OVERVIEW"] + [""] * 13)
    values.append(["Total Unique Learners", unique_learners] + [""] * 12)
    values.append(["Total Queries Found", total_queries] + [""] * 12)
    values.append(["Total Replied (TAs)", total_replied] + [""] * 12)
    values.append(["Pending Queries", pending_queries] + [""] * 12)
    values.append(["Avg Course Response Time (Hrs)", avg_course_time] + [""] * 12)

    # Spacer
    values.append([""] * 14)
    values.append([""] * 14)

    # TA Performance breakdown calculations
    ta_stats = {}
    for r in flat_rows:
        ta_name = r.get("First Responder")
        if not ta_name: continue
        if ta_name not in ta_stats:
            ta_stats[ta_name] = {"replies": 0, "ontime": 0, "delayed": 0, "durs": []}
        
        ta_stats[ta_name]["replies"] += 1
        if r.get("SLA Status") == "On Time Response": ta_stats[ta_name]["ontime"] += 1
        if r.get("SLA Status") == "Delayed Response": ta_stats[ta_name]["delayed"] += 1
        if isinstance(r.get("Duration (Hours)"), (int, float)):
            ta_stats[ta_name]["durs"].append(r["Duration (Hours)"])

    ta_start = sum_start + 8
    values.append(["TA PERFORMANCE BREAKDOWN (TREND ANALYSIS)"] + [""] * 13)
    values.append(["TA Name", "Replies", "On-Time", "Delayed", "Avg Time (Hrs)", "SLA Violation %"] + [""] * 8)

    for ta_name, s in ta_stats.items():
        avg_t = round(sum(s["durs"])/len(s["durs"]), 2) if s["durs"] else 0
        sla_v = round((s["delayed"]/s["replies"])*100, 1) if s["replies"] > 0 else 0
        values.append([ta_name, s["replies"], s["ontime"], s["delayed"], avg_t, f"{sla_v}%"] + [""] * 8)

    # Write all data at once
    ws.update(range_name=f"A1:N{len(values)}", values=values)

    # --- Apply formatting ---
    print("  Formatting Google Sheet cells (using batch updates)...")
    
    # 1. Merged Regions (Note: Merges are still structural but we execute them first)
    ws.merge_cells("A1:N1")
    if not flat_rows:
        ws.merge_cells("A3:N3")
    
    ws.merge_cells(f"A{sum_start}:D{sum_start}")
    for i in range(1, 6):
        ws.merge_cells(f"B{sum_start+i}:D{sum_start+i}")

    ws.merge_cells(f"A{ta_start}:G{ta_start}")

    # Reusable formatting assets
    color_white = hex_to_color("FFFFFF")
    color_black = hex_to_color("000000")
    color_gray  = hex_to_color("808080")
    border_thin = Border(style='SOLID', color=hex_to_color(C_BORDER))
    borders_all = Borders(top=border_thin, bottom=border_thin, left=border_thin, right=border_thin)

    with batch_updater(sh) as formatter:
        # Format Title (Row 1)
        formatter.format_cell_range(ws, "A1:N1", CellFormat(
            backgroundColor=hex_to_color(C_TITLE_BG),
            textFormat=TextFormat(bold=True, fontSize=12, foregroundColor=color_white),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE"
        ))

        # Format Headers (Row 2)
        formatter.format_cell_range(ws, "A2:N2", CellFormat(
            backgroundColor=hex_to_color(C_HDR_BG),
            textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=color_white),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=borders_all
        ))

        # Data Rows formatting
        if not flat_rows:
            formatter.format_cell_range(ws, "A3:N3", CellFormat(
                textFormat=TextFormat(italic=True, foregroundColor=color_gray),
                horizontalAlignment="CENTER",
                verticalAlignment="MIDDLE"
            ))
        else:
            for ri in range(3, last_row + 1):
                alt_color = C_ALT_BG if ri % 2 == 0 else "FFFFFF"
                
                # Apply base alternating background and borders
                formatter.format_cell_range(ws, f"A{ri}:N{ri}", CellFormat(
                    backgroundColor=hex_to_color(alt_color),
                    textFormat=TextFormat(fontSize=10, foregroundColor=color_black),
                    verticalAlignment="MIDDLE",
                    borders=borders_all
                ))

                # Retrieve from flat_rows in memory instead of ws.cell()!
                r = flat_rows[ri - 3]
                val_replied = r.get("Replied")
                val_sla = r.get("SLA Status")
                
                if val_replied == "Yes":
                    formatter.format_cell_range(ws, f"H{ri}", CellFormat(
                        backgroundColor=hex_to_color(C_YES_BG),
                        textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=hex_to_color(C_YES_FG)),
                        horizontalAlignment="CENTER"
                    ))
                elif val_replied == "No":
                    formatter.format_cell_range(ws, f"H{ri}", CellFormat(
                        backgroundColor=hex_to_color(C_NO_BG),
                        textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=hex_to_color(C_NO_FG)),
                        horizontalAlignment="CENTER"
                    ))

                if val_sla == "On Time Response":
                    formatter.format_cell_range(ws, f"K{ri}", CellFormat(
                        backgroundColor=hex_to_color("C6EFCE"),
                        textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=hex_to_color("006100")),
                        horizontalAlignment="CENTER"
                    ))
                elif val_sla == "Delayed Response":
                    formatter.format_cell_range(ws, f"K{ri}", CellFormat(
                        backgroundColor=hex_to_color("FFC7CE"),
                        textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=hex_to_color("9C0006")),
                        horizontalAlignment="CENTER"
                    ))

        # Format Course Overview Table
        formatter.format_cell_range(ws, f"A{sum_start}:D{sum_start}", CellFormat(
            backgroundColor=hex_to_color("4472C4"),
            textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=color_white),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=borders_all
        ))

        for i in range(1, 6):
            formatter.format_cell_range(ws, f"A{sum_start+i}:D{sum_start+i}", CellFormat(
                textFormat=TextFormat(bold=(i==0), fontSize=10, foregroundColor=color_black),
                verticalAlignment="MIDDLE",
                borders=borders_all
            ))

        # Format TA Breakdown Table
        formatter.format_cell_range(ws, f"A{ta_start}:G{ta_start}", CellFormat(
            backgroundColor=hex_to_color("C00000"),
            textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=color_white),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=borders_all
        ))

        formatter.format_cell_range(ws, f"A{ta_start+1}:F{ta_start+1}", CellFormat(
            backgroundColor=hex_to_color("D9D9D9"),
            textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=color_black),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=borders_all
        ))

        ta_count = len(ta_stats)
        for i in range(2, ta_count + 2):
            formatter.format_cell_range(ws, f"A{ta_start+i}:F{ta_start+i}", CellFormat(
                textFormat=TextFormat(fontSize=10, foregroundColor=color_black),
                verticalAlignment="MIDDLE",
                borders=borders_all
            ))

    # Lock Headers / Freeze panes
    ws.freeze(rows=2)
    
    # Auto-resize columns
    ws.columns_auto_resize(1, 14)
    print("  Google Sheet pushed and formatted successfully!")


def fetch_worksheet_data(ws):
    """Parallel worker to fetch sheet content."""
    try:
        values = ws.get_all_values()
        return ws.title, values
    except Exception:
        return ws.title, None


def parse_date_string(date_str):
    """Robust utility to parse custom date format strings from report."""
    if not date_str:
        return None
    ds = date_str.strip()
    ds = re.sub(r'\s+', ' ', ds) # normalize spaces
    formats = [
        "%m/%d/%y, %I:%M %p",      # e.g. 4/8/25, 10:54 PM
        "%m/%d/%Y, %I:%M %p",      # e.g. 4/8/2025, 10:54 PM
        "%d/%m/%y, %I:%M %p",
        "%d/%m/%Y, %I:%M %p",
        "%b %d, %Y %I:%M %p",      # e.g. Apr 24, 2025 10:32 PM
        "%b %d, %Y, %I:%M %p",
        "%Y-%m-%d %H:%M:%S",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(ds, fmt)
        except ValueError:
            continue
    return None



# In-memory cache for dashboard data
_dashboard_cache = {"data": None, "fetched_at": 0}
CACHE_TTL_SECONDS = 60  # Serve from cache for 60 seconds between Google Sheets calls

def get_dashboard_data():
    """Retrieve all worksheet data concurrently and return a structured raw data payload.
    Results are cached in memory for CACHE_TTL_SECONDS to minimise Google Sheets API calls.
    """
    import time as _time
    now = _time.time()

    # Return cached result if it's still fresh
    if _dashboard_cache["data"] is not None and (now - _dashboard_cache["fetched_at"]) < CACHE_TTL_SECONDS:
        return _dashboard_cache["data"]


    gc = get_google_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheets = sh.worksheets()

    # Read all sheets in parallel using thread pool
    sheet_data_list = []
    with ThreadPoolExecutor(max_workers=8) as executor:
        results = executor.map(fetch_worksheet_data, worksheets)
        for title, values in results:
            if values:
                sheet_data_list.append((title, values))

    raw_queries = []
    unique_courses = set()
    unique_cohorts = set()

    for title, values in sheet_data_list:
        if len(values) < 2:
            continue
        # Verify if it contains valid headers on Row 2
        headers = values[1]
        if len(headers) > 7 and headers[0] == "Cohort" and headers[7] == "Replied":
            # Parse data rows starting at Row 3 (index 2)
            for row in values[2:]:
                if not row or not row[0] or row[0].strip() == "" or "COURSE OVERVIEW" in row[0]:
                    break
                # Ensure row has 14 columns
                padded = row + [""] * (14 - len(row))
                
                cohort = padded[0].strip()
                course_name = padded[1].strip()
                topic = padded[2].strip()
                learner_name = padded[3].strip()
                student_id = padded[4].strip()
                learner_email = padded[5].strip()
                created_on = padded[6].strip()
                replied = padded[7].strip()
                replied_on = padded[8].strip()
                duration_str = padded[9].strip()
                sla_status = padded[10].strip()
                first_responder = padded[11].strip()
                replied_by_name = padded[12].strip()
                replied_by_email = padded[13].strip()

                if learner_name == "(No student posts yet)":
                    # Placeholder, but keep record of course/cohort
                    if course_name and course_name != "N/A":
                        unique_courses.add(course_name)
                    if cohort and cohort != "N/A":
                        unique_cohorts.add(cohort)
                    continue

                # Add to lists
                if course_name and course_name != "N/A":
                    unique_courses.add(course_name)
                if cohort and cohort != "N/A":
                    unique_cohorts.add(cohort)

                # Convert duration
                duration_val = None
                if duration_str:
                    try:
                        duration_val = float(duration_str)
                    except ValueError:
                        pass

                # Parse created date to ISO for frontend sorting/trends
                created_dt = parse_date_string(created_on)
                created_dt_iso = created_dt.isoformat() if created_dt else None

                # Validate first responder using TA_EMAILS_SET
                import Disussion_Automate as da
                r_emails = [e.strip().lower() for e in replied_by_email.split(",") if e.strip()]
                ta_emails_in_row = [e for e in r_emails if e in da.TA_EMAILS_SET]
                
                if not ta_emails_in_row:
                    first_responder = ""
                    if replied == "Yes":
                        replied = "No"
                        duration_val = None
                        sla_status = "Pending Response"

                query_obj = {
                    "cohort": cohort,
                    "course_name": course_name,
                    "topic": topic,
                    "learner_name": learner_name,
                    "student_id": student_id,
                    "learner_email": learner_email,
                    "created_on": created_on,
                    "created_dt_iso": created_dt_iso,
                    "replied": replied,
                    "replied_on": replied_on,
                    "duration_hours": duration_val,
                    "sla_status": sla_status,
                    "first_responder": first_responder if (first_responder and first_responder != "N/A") else "",
                    "replied_by_name": replied_by_name,
                    "replied_by_email": replied_by_email
                }
                raw_queries.append(query_obj)

    result = {
        "queries": raw_queries,
        "courses": sorted(list(unique_courses)),
        "cohorts": sorted(list(unique_cohorts)),
        "spreadsheet_id": SPREADSHEET_ID
    }

    # Store in cache
    import time as _time
    _dashboard_cache["data"] = result
    _dashboard_cache["fetched_at"] = _time.time()

    return result
