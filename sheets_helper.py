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
SPREADSHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1-XJZvmt5hTklebUPn-Sjui0dWkUGoUJF-eIbzm-Y3KA")
GOOGLE_CREDS_FILE = os.getenv("GOOGLE_CREDS_FILE", "linen-rex-436411-r4-9bba0db0c720.json")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

if os.path.isabs(GOOGLE_CREDS_FILE):
    CREDENTIALS_PATH = GOOGLE_CREDS_FILE
else:
    CREDENTIALS_PATH = os.path.join(SCRIPT_DIR, GOOGLE_CREDS_FILE)

# Styling Colors (Aligned with Excel palette)
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

    total_expected_rows = str(max(500, len(flat_rows) + 60))

    try:
        ws = sh.worksheet(safe_title)
        print(f"  Worksheet {safe_title} exists. Recreating worksheet to clear all formats and contents...")
        sh.del_worksheet(ws)
        ws = sh.add_worksheet(title=safe_title, rows=total_expected_rows, cols="20")
    except gspread.exceptions.WorksheetNotFound:
        print(f"  Worksheet {safe_title} not found. Creating...")
        ws = sh.add_worksheet(title=safe_title, rows=total_expected_rows, cols="20")

    # Pre-calculate overview metrics and TA performance stats
    valid_rows = [
        r for r in flat_rows 
        if r.get("Learner Name") and r.get("Learner Name") not in ("(No student posts yet)", "N/A", "Unknown", "User#None", "")
    ]
    unique_learners = len(set(r.get("Learner Name") for r in valid_rows))
    total_queries   = len(valid_rows)
    total_replied   = len([r for r in valid_rows if str(r.get("Replied", "")).startswith("Yes") or "replied" in str(r.get("Replied", "")).lower()])
    pending_queries = len([r for r in valid_rows if str(r.get("Replied", "")).startswith("No")])
    durations       = [r.get("Duration (Hours)") for r in valid_rows if isinstance(r.get("Duration (Hours)"), (int, float))]
    avg_course_time = round(sum(durations)/len(durations), 2) if durations else 0

    within_sla  = len([r for r in valid_rows if "On Time" in str(r.get("SLA Status", ""))])
    outside_sla = len([r for r in valid_rows if "Delayed" in str(r.get("SLA Status", ""))])

    # Pending queries average elapsed duration (hours)
    now_dt = datetime.now()
    pending_durs = []
    for r in valid_rows:
        if str(r.get("Replied", "")).startswith("No"):
            created_dt = parse_date_string(r.get("Created On"))
            if created_dt:
                p_hrs = max(0.0, round((now_dt - created_dt).total_seconds() / 3600.0, 1))
                pending_durs.append(p_hrs)
    avg_pending_time = round(sum(pending_durs)/len(pending_durs), 2) if pending_durs else 0

    # TA Performance breakdown calculations
    ta_stats = {}
    for r in valid_rows:
        assigned_ta = (r.get("Assigned TA (Needs to Reply)") or "").strip()
        first_resp  = (r.get("First Responder") or "").strip()
        sla_st      = str(r.get("SLA Status", "")).strip()
        rep_val     = str(r.get("Replied", "")).strip()
        dur_val     = r.get("Duration (Hours)")

        primary_ta = assigned_ta if (assigned_ta and assigned_ta != "Unassigned") else (first_resp or "Unassigned")
        if primary_ta not in ta_stats:
            ta_stats[primary_ta] = {"pending": 0, "replies": 0, "ontime": 0, "delayed": 0, "durs": []}

        if rep_val.startswith("No") or "Pending" in sla_st:
            ta_stats[primary_ta]["pending"] += 1
        else:
            resp_key = first_resp if first_resp else primary_ta
            if resp_key not in ta_stats:
                ta_stats[resp_key] = {"pending": 0, "replies": 0, "ontime": 0, "delayed": 0, "durs": []}
            ta_stats[resp_key]["replies"] += 1
            if "On Time" in sla_st:
                ta_stats[resp_key]["ontime"] += 1
            if "Delayed" in sla_st:
                ta_stats[resp_key]["delayed"] += 1
            if isinstance(dur_val, (int, float)):
                ta_stats[resp_key]["durs"].append(dur_val)

    # Construct Values Matrix
    values = []

    # Row 1: Title Block
    values.append([f"EDGEWOOD DISCUSSION BOARD AUDIT  |  {course_name} ({sis_id})"] + [""] * 13)

    # Row 2: Spacer
    values.append([""] * 14)

    # Row 3 to 6: COURSE OVERVIEW TABLE (JUST ABOVE LEARNER DETAILS TABLE)
    sum_start = 3
    values.append(["COURSE OVERVIEW"] + [""] * 13)
    values.append(["Total Post", "Total Response", "", "", "Pending Response", ""] + [""] * 8)
    values.append(["Total Queries", "Within SLA", "Outside SLA", "Avg. Response HR", "Pending Count", "Avg"] + [""] * 8)
    values.append([total_queries, within_sla, outside_sla, avg_course_time, pending_queries, avg_pending_time] + [""] * 8)

    # Row 7: Spacer
    values.append([""] * 14)

    # Row 8: Learner Details Table Headers (JUST BELOW SUMMARY TABLE)
    hdr_row = 8
    headers = [
        "Cohort", "Course Name", "Topic", "Learner Name", "Student ID", "Learner Email", 
        "Created On", "Replied", "Replied On", "Duration (Hours)", "SLA Status", 
        "First Responder", "Assigned TA (Needs to Reply)", "Assigned TA Email"
    ]
    values.append(headers)

    # Row 9 to last_row: Learner Data
    if not flat_rows:
        values.append(["No discussion boards or learner posts found in this course."] + [""] * 13)
        last_row = 9
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
                r.get("Assigned TA (Needs to Reply)", r.get("Replied By (Name)", "")),
                r.get("Assigned TA Email", r.get("Replied By (Email)", ""))
            ]
            values.append(row_vals)
        last_row = len(values)

    # Spacer
    values.append([""] * 14)

    # TA Performance Breakdown Table
    ta_start = len(values) + 1
    values.append(["TA PERFORMANCE BREAKDOWN (TREND ANALYSIS)"] + [""] * 13)
    values.append(["TA Name", "Pending Response", "Replies", "On-Time", "Delayed", "Avg Time (Hrs)", "SLA Violation %"] + [""] * 7)

    for ta_name, s in ta_stats.items():
        avg_t = round(sum(s["durs"])/len(s["durs"]), 2) if s["durs"] else 0
        sla_v = round((s["delayed"]/s["replies"])*100, 1) if s["replies"] > 0 else 0
        values.append([ta_name, s["pending"], s["replies"], s["ontime"], s["delayed"], avg_t, f"{sla_v}%"] + [""] * 7)

    # Write all data at once
    ws.update(range_name=f"A1:N{len(values)}", values=values)

    # --- Apply formatting ---
    print("  Formatting Google Sheet cells (using batch updates)...")
    
    # 1. Merged Regions
    ws.merge_cells("A1:N1")
    if not flat_rows:
        ws.merge_cells("A9:N9")
    
    ws.merge_cells(f"A{sum_start}:F{sum_start}")
    ws.merge_cells(f"B{sum_start+1}:D{sum_start+1}")
    ws.merge_cells(f"E{sum_start+1}:F{sum_start+1}")

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

        # Format Learner Table Headers (Row 8)
        formatter.format_cell_range(ws, "A8:N8", CellFormat(
            backgroundColor=hex_to_color(C_HDR_BG),
            textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=color_white),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=borders_all
        ))

        # Data Rows formatting (Row 9+)
        if not flat_rows:
            formatter.format_cell_range(ws, "A9:N9", CellFormat(
                textFormat=TextFormat(italic=True, foregroundColor=color_gray),
                horizontalAlignment="CENTER",
                verticalAlignment="MIDDLE"
            ))
        else:
            for ri in range(9, last_row + 1):
                alt_color = C_ALT_BG if ri % 2 == 0 else "FFFFFF"
                
                # Apply base alternating background and borders
                formatter.format_cell_range(ws, f"A{ri}:N{ri}", CellFormat(
                    backgroundColor=hex_to_color(alt_color),
                    textFormat=TextFormat(fontSize=10, foregroundColor=color_black),
                    verticalAlignment="MIDDLE",
                    borders=borders_all
                ))

                # Retrieve from flat_rows in memory (index = ri - 9)
                r = flat_rows[ri - 9]
                val_replied = str(r.get("Replied", "")).strip()
                val_sla = str(r.get("SLA Status", "")).strip()
                
                if val_replied == "Yes":
                    formatter.format_cell_range(ws, f"H{ri}", CellFormat(
                        backgroundColor=hex_to_color(C_YES_BG),
                        textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=hex_to_color(C_YES_FG)),
                        horizontalAlignment="CENTER"
                    ))
                elif "Other TA" in val_replied or "Replied by another TA" in val_replied:
                    formatter.format_cell_range(ws, f"H{ri}", CellFormat(
                        backgroundColor=hex_to_color("FFF2CC"),
                        textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=hex_to_color("7F6000")),
                        horizontalAlignment="CENTER"
                    ))
                elif val_replied.startswith("No"):
                    formatter.format_cell_range(ws, f"H{ri}", CellFormat(
                        backgroundColor=hex_to_color(C_NO_BG),
                        textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=hex_to_color(C_NO_FG)),
                        horizontalAlignment="CENTER"
                    ))

                if "On Time" in val_sla:
                    formatter.format_cell_range(ws, f"K{ri}", CellFormat(
                        backgroundColor=hex_to_color("C6EFCE"),
                        textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=hex_to_color("006100")),
                        horizontalAlignment="CENTER"
                    ))
                elif "Delayed" in val_sla:
                    formatter.format_cell_range(ws, f"K{ri}", CellFormat(
                        backgroundColor=hex_to_color("FFC7CE"),
                        textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=hex_to_color("9C0006")),
                        horizontalAlignment="CENTER"
                    ))

        # Format Course Overview Table Title
        formatter.format_cell_range(ws, f"A{sum_start}:F{sum_start}", CellFormat(
            backgroundColor=hex_to_color("4472C4"),
            textFormat=TextFormat(bold=True, fontSize=11, foregroundColor=color_white),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=borders_all
        ))

        # Category Headers (Row 1)
        formatter.format_cell_range(ws, f"A{sum_start+1}", CellFormat(
            backgroundColor=hex_to_color("1F3864"),
            textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=color_white),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=borders_all
        ))
        formatter.format_cell_range(ws, f"B{sum_start+1}:D{sum_start+1}", CellFormat(
            backgroundColor=hex_to_color("E2EFDA"),
            textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=hex_to_color("375623")),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=borders_all
        ))
        formatter.format_cell_range(ws, f"E{sum_start+1}:F{sum_start+1}", CellFormat(
            backgroundColor=hex_to_color("FCE4D6"),
            textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=hex_to_color("C00000")),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=borders_all
        ))

        # Sub Headers (Row 2)
        formatter.format_cell_range(ws, f"A{sum_start+2}:F{sum_start+2}", CellFormat(
            backgroundColor=hex_to_color("F2F2F2"),
            textFormat=TextFormat(bold=True, fontSize=9, foregroundColor=color_black),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=borders_all
        ))

        # Data Row (Row 3)
        formatter.format_cell_range(ws, f"A{sum_start+3}:F{sum_start+3}", CellFormat(
            backgroundColor=hex_to_color("FFFFFF"),
            textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=color_black),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=borders_all
        ))

        # Format TA Breakdown Table Title
        formatter.format_cell_range(ws, f"A{ta_start}:G{ta_start}", CellFormat(
            backgroundColor=hex_to_color("C00000"),
            textFormat=TextFormat(bold=True, fontSize=11, foregroundColor=color_white),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=borders_all
        ))

        # TA Column Headers
        formatter.format_cell_range(ws, f"A{ta_start+1}:G{ta_start+1}", CellFormat(
            backgroundColor=hex_to_color("D9D9D9"),
            textFormat=TextFormat(bold=True, fontSize=10, foregroundColor=color_black),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=borders_all
        ))

        ta_count = len(ta_stats)
        for i in range(2, ta_count + 2):
            formatter.format_cell_range(ws, f"A{ta_start+i}:G{ta_start+i}", CellFormat(
                textFormat=TextFormat(fontSize=10, foregroundColor=color_black),
                verticalAlignment="MIDDLE",
                borders=borders_all
            ))

    # Lock Headers / Freeze panes (locks Course Overview & Learner table headers)
    ws.freeze(rows=8)
    
    # Auto-resize columns
    try:
        ws.columns_auto_resize(1, 14)
    except Exception:
        pass
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



# ──────────────────────────────────────────────────────────────────
#  DYNAMIC TA & LEARNER MAPPING (from Google Sheet "ImportData_TA" tab)
# ──────────────────────────────────────────────────────────────────

_import_data_cache = {"data": None, "fetched_at": 0}
IMPORT_DATA_CACHE_TTL = 60  # Cache for 60 seconds

# Local safety fallback for TA emails in case neither ImportData_TA nor TA Details is reachable
LOCAL_SAFETY_TA_EMAILS = frozenset(e.strip().lower() for e in [
    "asharma@edgewood.edu", "PKulkarni@edgewood.edu", "seegupta@edgewood.edu",
    "nsukhani@edgewood.edu", "ppriyadarshani@edgewood.edu", "ssahai@edgewood.edu",
    "vpagaria@edgewood.edu", "ssaluja@edgewood.edu", "JSelvaraj@edgewood.edu",
    "aghosh@edgewood.edu", "pnema@edgewood.edu", "MMehndiratta@edgewood.edu",
    "mnaruka@edgewood.edu", "CMedatwal@edgewood.edu", "kkaur@edgewood.edu",
    "arastogi@edgewood.edu", "PElavarasu@edgewood.edu", "mkaur@edgewood.edu",
    "Mmagry@edgewood.edu", "RUdhani@edgewood.edu", "rmishra@edgewood.edu",
    "RPrasad@edgewood.edu", "SobJose@edgewood.edu", "sonsharma@edgewood.edu",
    "nareshsukhani@gmail.com", "patolejayashree.1980@gmail.com", "aruna.deshpande@gmail.com",
    "Jpatole@edgewood.edu", "jranjan@edgewood.edu", "vpalmisano@edgewood.edu",
    "AnkPandey@edgewood.edu", "LMaxwell@edgewood.edu"
] if e)

def normalize_course_code(course_str):
    """Normalize course codes like 'BUS-940-UC6--Fall-2026', 'BUS 940', 'BUS_940' into 'BUS 940'."""
    if not course_str:
        return ""
    s = str(course_str).strip().upper()
    m = re.search(r'([A-Z]{3})[-_\s]?([0-9]{3})', s)
    if m:
        return f"{m.group(1)} {m.group(2)}"
    clean = re.sub(r'[\s\-_]+', ' ', s).strip()
    return clean

def get_import_data_ta(force_refresh=False):
    """Fetch assigned TA mappings from Google Sheet 'ImportData_TA' tab.
    Columns in sheet:
      A (0): Concat (e.g. '642816BUS 940')
      B (1): Course (e.g. 'BUS 940')
      C (2): Edgewood ID (Student ID) (e.g. '642816')
      D (3): Canvas Name (e.g. 'Doe, John')
      E (4): Edgewood Email / Student Email (e.g. 'jdoe@edgewood.edu')
      F (5): Grader/TA Name (e.g. 'Ashish Sharma')
      G (6): Grader Edgewood Email / TA Email (e.g. 'asharma@edgewood.edu')

    Gracefully falls back to 'TA Details' tab or local safety set if ImportData_TA
    has not yet been populated with IMPORTRANGE data.
    """
    import time as _time
    now = _time.time()
    if not force_refresh and _import_data_cache["data"] and (now - _import_data_cache["fetched_at"]) < IMPORT_DATA_CACHE_TTL:
        return _import_data_cache["data"]

    by_course_email = {}
    by_course_id = {}
    by_email = {}
    all_ta_emails = set()
    ta_name_map = {}

    try:
        gc = get_google_client()
        sh = gc.open_by_key(SPREADSHEET_ID)
        
        # 1. Locate ImportData_TA worksheet
        ws_import = None
        ws_ta_details = None
        for w in sh.worksheets():
            t_clean = w.title.lower().replace(' ', '').replace('_', '')
            if 'importdata' in t_clean and 'ta' in t_clean:
                ws_import = w
            elif 'tadetail' in t_clean:
                ws_ta_details = w

        # Parse ImportData_TA if present and populated
        if ws_import is not None:
            rows = ws_import.get_all_values()
            if len(rows) > 1:
                for r in rows[1:]:  # skip header
                    if len(r) < 7:
                        continue
                    course_raw = r[1].strip()
                    norm_c = normalize_course_code(course_raw)
                    sid = r[2].strip()
                    sname = r[3].strip()
                    semail = r[4].strip().lower()
                    ta_name = r[5].strip()
                    ta_email = r[6].strip().lower()

                    if not norm_c:
                        continue

                    record = {
                        "course": norm_c,
                        "course_raw": course_raw,
                        "student_id": sid,
                        "student_name": sname,
                        "student_email": semail,
                        "ta_name": ta_name,
                        "ta_email": ta_email
                    }

                    if semail:
                        by_course_email[(norm_c, semail)] = record
                        by_email[semail] = record
                    if sid:
                        by_course_id[(norm_c, sid)] = record

                    if ta_email:
                        all_ta_emails.add(ta_email)
                        if ta_name:
                            ta_name_map[ta_email] = ta_name

        # 2. If no TA emails were extracted from ImportData_TA, fall back to 'TA Details' tab
        if not all_ta_emails and ws_ta_details is not None:
            rows_ta = ws_ta_details.get_all_values()
            for r in rows_ta:
                if not r or not r[0]:
                    continue
                email = r[0].strip().lower()
                if "@" in email and "email" not in email:
                    all_ta_emails.add(email)

        # 3. Always include local safety set so operations never fail
        all_ta_emails.update(LOCAL_SAFETY_TA_EMAILS)

        result = {
            "by_course_email": by_course_email,
            "by_course_id": by_course_id,
            "by_email": by_email,
            "all_ta_emails": frozenset(all_ta_emails),
            "ta_name_map": ta_name_map
        }
        _import_data_cache["data"] = result
        _import_data_cache["fetched_at"] = now
        return result

    except Exception as e:
        print(f"  [!] Failed to load ImportData_TA / TA Details: {e}")
        fallback_res = {
            "by_course_email": by_course_email,
            "by_course_id": by_course_id,
            "by_email": by_email,
            "all_ta_emails": frozenset(all_ta_emails or LOCAL_SAFETY_TA_EMAILS),
            "ta_name_map": ta_name_map
        }
        return _import_data_cache["data"] or fallback_res


def get_assigned_ta(course_code, student_email, student_id=None):
    """Retrieve the specifically assigned TA for a learner in a course."""
    directory = get_import_data_ta()
    norm_c = normalize_course_code(course_code)
    semail = (student_email or "").strip().lower()
    sid = str(student_id or "").strip()

    # 1. Match by Course + Student Email
    if semail and (norm_c, semail) in directory["by_course_email"]:
        return directory["by_course_email"][(norm_c, semail)]

    # 2. Match by Course + Student ID
    if sid and (norm_c, sid) in directory["by_course_id"]:
        return directory["by_course_id"][(norm_c, sid)]

    # 3. Fallback: match by Student Email across courses if unique
    if semail and semail in directory["by_email"]:
        return directory["by_email"][semail]

    return None


def get_all_ta_emails(force_refresh=False):
    """Return all unique TA email addresses from ImportData_TA / TA Details."""
    directory = get_import_data_ta(force_refresh=force_refresh)
    return directory["all_ta_emails"]


def get_ta_emails(force_refresh=False):
    """Backward compatibility helper for get_all_ta_emails."""
    return get_all_ta_emails(force_refresh=force_refresh)


# In-memory cache for dashboard data
_dashboard_cache = {"data": None, "fetched_at": 0}
CACHE_TTL_SECONDS = 60  # Serve from cache for 60 seconds between Google Sheets calls

def get_dashboard_data(force_refresh=False):
    """Fetches and parses all worksheets into clean JSON structure for the UI dashboard."""
    import time as _time
    now = _time.time()

    if not force_refresh and _dashboard_cache["data"] is not None and (now - _dashboard_cache["fetched_at"]) < CACHE_TTL_SECONDS:
        return _dashboard_cache["data"]

    gc = get_google_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheets = sh.worksheets()

    sheet_data_list = []
    with ThreadPoolExecutor(max_workers=8) as executor:
        results = executor.map(fetch_worksheet_data, worksheets)
        for title, values in results:
            if not values:
                continue
            clean_title = title.lower().replace(" ", "").replace("_", "")
            # Ignore TA config tabs
            if ("importdata" in clean_title and "ta" in clean_title) or "tadetail" in clean_title:
                continue
            sheet_data_list.append((title, values))

    raw_queries = []
    unique_courses = set()
    unique_cohorts = set()
    ta_emails_set = get_all_ta_emails()

    for title, values in sheet_data_list:
        if len(values) < 2:
            continue
        header_idx = None
        for idx, row in enumerate(values):
            if row and len(row) > 7 and row[0].strip().lower() == "cohort" and row[7].strip().lower() == "replied":
                header_idx = idx
                break
        if header_idx is not None:
            for row in values[header_idx + 1:]:
                if not row or not row[0] or row[0].strip() == "" or "COURSE OVERVIEW" in row[0] or "TA PERFORMANCE" in row[0]:
                    break
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

                if not learner_name or learner_name in ("(No student posts yet)", "N/A", "Unknown", "User#None"):
                    if course_name and course_name != "N/A":
                        unique_courses.add(course_name)
                    if cohort and cohort != "N/A":
                        unique_cohorts.add(cohort)
                    continue

                if course_name and course_name != "N/A":
                    unique_courses.add(course_name)
                if cohort and cohort != "N/A":
                    unique_cohorts.add(cohort)

                duration_val = None
                if duration_str:
                    try:
                        duration_val = float(duration_str)
                    except ValueError:
                        pass

                created_dt = parse_date_string(created_on)
                created_dt_iso = created_dt.isoformat() if created_dt else None

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
                    "replied_by_name": replied_by_name,   # backward compat alias
                    "replied_by_email": replied_by_email, # backward compat alias
                    "assigned_ta_name": replied_by_name,
                    "assigned_ta_email": replied_by_email
                }
                raw_queries.append(query_obj)

    result = {
        "queries": raw_queries,
        "courses": sorted(list(unique_courses)),
        "cohorts": sorted(list(unique_cohorts)),
        "spreadsheet_id": SPREADSHEET_ID
    }

    _dashboard_cache["data"] = result
    _dashboard_cache["fetched_at"] = _time.time()
    return result


def _scan_single_worksheet(sh, course_shell_id, pending_by_ta_agg, ta_stats_agg=None, course_stats_agg=None):
    """Internal helper: scans one worksheet and aggregates pending queries into pending_by_ta_agg dict.
    Also computes full course overview stats and per-TA response metrics.
    Returns (found, ws_title, course_code, local_pending_count)."""
    if ta_stats_agg is None:
        ta_stats_agg = {}
    if course_stats_agg is None:
        course_stats_agg = {
            "total_posts": 0, "total_replied": 0, "within_sla": 0, "outside_sla": 0,
            "pending_count": 0, "durs": [], "pending_durs": []
        }

    clean_target = str(course_shell_id).strip()
    target_norm = normalize_course_code(clean_target)
    safe_target = re.sub(r'[:\\\/\?\*\[\]]', '_', clean_target)[:31].lower()
    alphanumeric_target = re.sub(r'[\s\-_]+', '', clean_target).lower()

    ws = None
    for w in sh.worksheets():
        w_title_lower = w.title.lower()
        w_title_clean = w_title_lower.replace(' ', '').replace('_', '')
        w_title_alpha = re.sub(r'[\s\-_]+', '', w_title_lower)
        if ('importdata' in w_title_clean and 'ta' in w_title_clean) or 'tadetail' in w_title_clean:
            continue
        if w_title_lower == clean_target.lower() or w_title_lower == safe_target:
            ws = w
            break
        if target_norm and target_norm.lower() in w_title_lower:
            ws = w
            break
        if alphanumeric_target and alphanumeric_target in w_title_alpha:
            ws = w
            break

    if not ws:
        return False, course_shell_id, target_norm, 0

    rows = ws.get_all_values()
    local_pending = 0
    if len(rows) < 3:
        return True, ws.title, target_norm, 0

    now_dt = datetime.now()

    header_idx = None
    for idx, r in enumerate(rows):
        if r and len(r) > 7 and r[0].strip().lower() == "cohort" and r[7].strip().lower() == "replied":
            header_idx = idx
            break
    start_row = (header_idx + 1) if header_idx is not None else 2

    for r in rows[start_row:]:
        if not r or not r[0] or "COURSE OVERVIEW" in r[0] or "TA PERFORMANCE" in r[0]:
            break
        padded = r + [""] * (14 - len(r))
        cohort = padded[0].strip()
        c_name = padded[1].strip() or target_norm
        topic = padded[2].strip()
        l_name = padded[3].strip()
        s_id = padded[4].strip()
        l_email = padded[5].strip().lower()
        created_on = padded[6].strip()
        replied = padded[7].strip()
        replied_on = padded[8].strip()
        duration_str = padded[9].strip()
        sla_status = padded[10].strip()
        first_responder = padded[11].strip()
        assigned_ta_col = padded[12].strip() if len(padded) > 12 else ""
        assigned_ta_email_col = padded[13].strip() if len(padded) > 13 else ""

        if not l_name or l_name in ("(No student posts yet)", "N/A", "Unknown", "User#None"):
            continue

        course_stats_agg["total_posts"] += 1

        # Determine assigned TA from mapping or sheet column
        ta_info = get_assigned_ta(c_name, l_email, s_id)
        if ta_info:
            ta_name = ta_info.get("ta_name", "Assigned TA")
            ta_email = ta_info.get("ta_email", "")
        elif assigned_ta_col:
            ta_name = assigned_ta_col
            ta_email = assigned_ta_email_col
        else:
            ta_name = "Unassigned / General Pool"
            ta_email = ""

        if ta_name not in ta_stats_agg:
            ta_stats_agg[ta_name] = {
                "pending": 0, "replies": 0, "ontime": 0, "delayed": 0,
                "durs": [], "pending_durs": []
            }

        is_pending = (replied == "No" or replied.startswith("No") or "Pending" in sla_status)
        if is_pending:
            local_pending += 1
            course_stats_agg["pending_count"] += 1
            ta_stats_agg[ta_name]["pending"] += 1

            c_dt = parse_date_string(created_on)
            if c_dt:
                age_hrs = max(0.0, round((now_dt - c_dt).total_seconds() / 3600.0, 1))
                course_stats_agg["pending_durs"].append(age_hrs)
                ta_stats_agg[ta_name]["pending_durs"].append(age_hrs)

            key = (ta_name, ta_email)
            if key not in pending_by_ta_agg:
                pending_by_ta_agg[key] = {
                    "ta_name": ta_name,
                    "ta_email": ta_email,
                    "queries": []
                }
            pending_by_ta_agg[key]["queries"].append({
                "learner_name": l_name,
                "student_id": s_id,
                "learner_email": l_email,
                "topic": topic,
                "created_on": created_on,
                "cohort": cohort,
                "course_shell": ws.title,
                "course_code": c_name
            })
        else:
            course_stats_agg["total_replied"] += 1
            responder = first_responder if (first_responder and first_responder != "N/A") else ta_name
            if responder not in ta_stats_agg:
                ta_stats_agg[responder] = {
                    "pending": 0, "replies": 0, "ontime": 0, "delayed": 0,
                    "durs": [], "pending_durs": []
                }
            ta_stats_agg[responder]["replies"] += 1

            if "On Time" in sla_status:
                course_stats_agg["within_sla"] += 1
                ta_stats_agg[responder]["ontime"] += 1
            elif "Delayed" in sla_status:
                course_stats_agg["outside_sla"] += 1
                ta_stats_agg[responder]["delayed"] += 1

            if duration_str:
                try:
                    dur_val = float(duration_str)
                    course_stats_agg["durs"].append(dur_val)
                    ta_stats_agg[responder]["durs"].append(dur_val)
                except ValueError:
                    pass

    return True, ws.title, target_norm, local_pending


def get_pending_ta_summary(course_shell_input):
    """Scans the Google Sheet for one or more course shells and groups pending learner queries by assigned TA.
    Supports comma or space separated shells, or 'ALL' to scan every audited tab."""
    import re as _re
    gc = get_google_client()
    sh = gc.open_by_key(SPREADSHEET_ID)

    raw_input = str(course_shell_input).strip()
    is_all = raw_input.upper() == "ALL"

    # Split by comma or semicolon
    shells = [s.strip() for s in _re.split(r'[,;]+', raw_input) if s.strip()]

    if is_all or not shells:
        # Collect all non-config worksheets
        shells = []
        for w in sh.worksheets():
            t_clean = w.title.lower().replace(' ', '').replace('_', '')
            if ('importdata' in t_clean and 'ta' in t_clean) or 'tadetail' in t_clean:
                continue
            shells.append(w.title)

    pending_by_ta_agg = {}  # key=(ta_name, ta_email)
    course_stats_agg = {
        "total_posts": 0,
        "total_replied": 0,
        "within_sla": 0,
        "outside_sla": 0,
        "pending_count": 0,
        "durs": [],
        "pending_durs": []
    }
    ta_stats_agg = {}
    total_pending = 0
    course_shells_found = []
    course_shells_missing = []

    for shell in shells:
        found, ws_title, course_code, local_count = _scan_single_worksheet(
            sh, shell, pending_by_ta_agg, ta_stats_agg, course_stats_agg
        )
        if found:
            course_shells_found.append(ws_title)
            total_pending += local_count
        else:
            course_shells_missing.append(shell)

    if not course_shells_found:
        return {
            "found": False,
            "error": f"No audited worksheets found for: {', '.join(course_shells_missing)}. Please run the audit first.",
            "course_shell": raw_input,
            "ta_groups": [],
            "total_pending": 0,
            "course_overview": {
                "total_posts": 0, "total_replied": 0, "within_sla": 0, "outside_sla": 0,
                "avg_response_hr": 0, "pending_count": 0, "avg_pending_hr": 0
            }
        }

    ta_groups = list(pending_by_ta_agg.values())
    ta_groups.sort(key=lambda x: len(x["queries"]), reverse=True)

    # Attach performance stats to each TA group so pending TA info is never blank
    for g in ta_groups:
        s = ta_stats_agg.get(g["ta_name"]) or {}
        durs = s.get("durs", [])
        reps = s.get("replies", 0)
        delays = s.get("delayed", 0)
        g["stats"] = {
            "pending": s.get("pending", len(g["queries"])),
            "replies": reps,
            "ontime": s.get("ontime", 0),
            "delayed": delays,
            "avg_time": round(sum(durs)/len(durs), 2) if durs else 0,
            "sla_violation_pct": round((delays/reps)*100, 1) if reps > 0 else 0
        }

    c_durs = course_stats_agg["durs"]
    c_p_durs = course_stats_agg["pending_durs"]
    avg_resp = round(sum(c_durs)/len(c_durs), 2) if c_durs else 0
    avg_pend = round(sum(c_p_durs)/len(c_p_durs), 2) if c_p_durs else 0

    course_overview = {
        "total_posts": course_stats_agg["total_posts"],
        "total_replied": course_stats_agg["total_replied"],
        "within_sla": course_stats_agg["within_sla"],
        "outside_sla": course_stats_agg["outside_sla"],
        "avg_response_hr": avg_resp,
        "pending_count": course_stats_agg["pending_count"],
        "avg_pending_hr": avg_pend
    }

    primary_shell = course_shells_found[0] if len(course_shells_found) == 1 else ", ".join(course_shells_found)
    primary_code = normalize_course_code(course_shells_found[0]) if course_shells_found else raw_input

    result = {
        "found": True,
        "course_shell": primary_shell,
        "course_code": primary_code,
        "course_shells_scanned": course_shells_found,
        "course_shells_missing": course_shells_missing,
        "total_pending": total_pending,
        "course_overview": course_overview,
        "ta_groups": ta_groups
    }
    if is_all:
        result["scanned_all"] = True
    return result
