"""
=============================================================================
EDGEWOOD COLLEGE - DISCUSSION BOARD AUDIT TOOL (BUSINESS TRACKER)
Enhanced Canvas API Automation with Dynamic TA Mapping, 24-Hour SLA Validation,
Nested Threaded Replies Support, Multi-Tab Output, and Google Sheets Sync.
=============================================================================
"""

import os
import sys
import re
import requests
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import sheets_helper

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────────────────────────
#  CONFIG & CONSTANTS
# ──────────────────────────────────────────────────────────────────

API_URL      = os.getenv("CANVAS_API_URL", "https://edgewood.instructure.com/api/v1")
ACCESS_TOKEN = os.getenv("CANVAS_ACCESS_TOKEN")

HEADERS = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Accept"       : "application/json"
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))


def extract_all_replies(replies_list):
    """Recursively extracts all replies including nested threaded replies."""
    flat = []
    if not replies_list:
        return flat
    for r in replies_list:
        flat.append(r)
        if r.get("replies"):
            flat.extend(extract_all_replies(r.get("replies")))
    return flat


# ──────────────────────────────────────────────────────────────────
#  CANVAS API HELPERS
# ──────────────────────────────────────────────────────────────────

def check_connection():
    """Verify Canvas API credentials."""
    if not ACCESS_TOKEN:
        print("\n  [!] ERROR: CANVAS_ACCESS_TOKEN is missing from .env")
        return False
    try:
        r = SESSION.get(f"{API_URL}/users/self", timeout=10)
        if r.status_code == 200:
            user = r.json()
            print(f"  Connected to Canvas as: {user.get('name', 'Unknown')} (ID: {user.get('id')})")
            return True
        elif r.status_code == 401:
            print("\n  [!] ERROR: Invalid or expired access token (401 Unauthorized).")
            return False
        else:
            print(f"\n  [!] Connection failed with HTTP {r.status_code}: {r.text[:200]}")
            return False
    except requests.exceptions.RequestException as e:
        print(f"\n  [!] Network error connecting to Canvas: {e}")
        return False


def find_course(sis_id):
    """Find course by SIS ID using targeted API search."""
    # 1. Direct sis_course_id lookup
    url = f"{API_URL}/courses/sis_course_id:{sis_id}"
    try:
        r = SESSION.get(url, timeout=15)
        if r.status_code == 200:
            return r.json()
    except requests.exceptions.RequestException:
        pass

    # 2. Search query fallback
    url = f"{API_URL}/accounts/self/courses"
    params = {"search_term": sis_id, "per_page": 20}
    try:
        r = SESSION.get(url, params=params, timeout=15)
        if r.status_code == 200:
            for c in r.json():
                if (c.get("sis_course_id") or "").strip().lower() == sis_id.strip().lower():
                    return c
                if sis_id.strip().lower() in (c.get("course_code") or "").strip().lower():
                    return c
    except requests.exceptions.RequestException:
        pass

    return None


def get_course_users_with_emails(course_id):
    """
    [PERF] Single API call with include[]=email that extracts BOTH:
      1. user_map  : {user_id: {"name": ..., "email": ...}}
      2. staff_ids : set of user_ids with Teacher/TA/Facilitator enrollment
    Saves one full paginated API traversal per course!
    """
    user_map  = {}
    staff_ids = set()
    staff_types = {"TaEnrollment", "TeacherEnrollment", "DesignerEnrollment"}

    url    = f"{API_URL}/courses/{course_id}/users"
    params = {"per_page": 100, "include[]": ["email", "enrollments"]}

    while url:
        try:
            r = SESSION.get(url, params=params, timeout=20)
            if r.status_code != 200:
                break
            for u in r.json():
                uid   = u["id"]
                name  = u.get("name") or u.get("sortable_name") or f"User#{uid}"
                email = (u.get("email") or u.get("login_id") or "").strip().lower()
                user_map[uid] = {
                    "name"      : name,
                    "email"     : email,
                    "student_id": u.get("sis_user_id") or "N/A"
                }

                # Check enrollments directly from the user record
                for e in u.get("enrollments", []):
                    if e.get("role") in staff_types or e.get("type") in staff_types:
                        staff_ids.add(uid)
                        break

            # Handle pagination
            url = None
            params = {}
            if "next" in r.links:
                url = r.links["next"]["url"]
        except requests.exceptions.RequestException:
            break

    return user_map, staff_ids


def get_discussion_topics(course_id):
    """Get all discussion topics for a course."""
    topics = []
    url    = f"{API_URL}/courses/{course_id}/discussion_topics"
    params = {"per_page": 50, "order_by": "recent_activity"}

    while url:
        try:
            r = SESSION.get(url, params=params, timeout=20)
            if r.status_code != 200:
                break
            topics.extend(r.json())
            url = r.links.get("next", {}).get("url")
            params = {}
        except requests.exceptions.RequestException:
            break

    return topics


def get_full_view(course_id, topic_id):
    """Get full topic view including nested replies."""
    url = f"{API_URL}/courses/{course_id}/discussion_topics/{topic_id}/view"
    try:
        r = SESSION.get(url, timeout=30)
        if r.status_code == 200:
            data = r.json()
            participants = {p["id"]: p for p in data.get("participants", [])}
            view = data.get("view", [])
            return participants, view
    except requests.exceptions.RequestException:
        pass
    return {}, []


def get_course_modules(course_id):
    """Get all modules for a course."""
    modules = []
    url = f"{API_URL}/courses/{course_id}/modules"
    params = {"per_page": 50}

    while url:
        try:
            r = SESSION.get(url, params=params, timeout=20)
            if r.status_code != 200:
                break
            modules.extend(r.json())
            url = r.links.get("next", {}).get("url")
            params = {}
        except requests.exceptions.RequestException:
            break

    return modules


def get_module_items(course_id, module_id):
    """Get items in a module."""
    items = []
    url = f"{API_URL}/courses/{course_id}/modules/{module_id}/items"
    params = {"per_page": 50}

    while url:
        try:
            r = SESSION.get(url, params=params, timeout=20)
            if r.status_code != 200:
                break
            items.extend(r.json())
            url = r.links.get("next", {}).get("url")
            params = {}
        except requests.exceptions.RequestException:
            break

    return items


def find_module_for_discussion(course_id, topic_id):
    """Search modules to see which one contains this discussion topic."""
    modules = get_course_modules(course_id)
    for m in modules:
        items = get_module_items(course_id, m["id"])
        for item in items:
            if item.get("type") == "Discussion" and item.get("content_id") == topic_id:
                return m.get("name", "Unknown Module")
    return None


# ──────────────────────────────────────────────────────────────────
#  DATE & FORMATTING HELPERS
# ──────────────────────────────────────────────────────────────────

def parse_iso(iso_str):
    """Parse ISO datetime string to naive UTC datetime."""
    if not iso_str:
        return None
    try:
        s = iso_str.replace("Z", "+00:00")
        dt = datetime.fromisoformat(s)
        return dt.replace(tzinfo=None)
    except Exception:
        return None


def calculate_duration_hours(start_iso, end_iso):
    """Calculate duration between two timestamps in hours (1 decimal place)."""
    start = parse_iso(start_iso)
    end   = parse_iso(end_iso)
    if start and end and end >= start:
        return round((end - start).total_seconds() / 3600.0, 1)
    return None


def format_created_on(iso_str):
    """Format: '4/8/25, 10:54 PM'"""
    dt = parse_iso(iso_str)
    if not dt:
        return ""
    hour = dt.hour % 12 or 12
    am_pm = "AM" if dt.hour < 12 else "PM"
    year_short = str(dt.year)[-2:]
    return f"{dt.month}/{dt.day}/{year_short}, {hour}:{dt.minute:02d} {am_pm}"


def format_replied_on(iso_str):
    """Format: 'Apr 24, 2025 10:32 PM'"""
    dt = parse_iso(iso_str)
    if not dt:
        return ""
    months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    month_str = months[dt.month - 1]
    hour = dt.hour % 12 or 12
    am_pm = "AM" if dt.hour < 12 else "PM"
    return f"{month_str} {dt.day}, {dt.year} {hour}:{dt.minute:02d} {am_pm}"


# ──────────────────────────────────────────────────────────────────
#  DATA PROCESSING
# ──────────────────────────────────────────────────────────────────

def get_cohort_and_course(sis_id, course_name_api):
    """Extract Cohort (e.g. UC5, UC6, UB1, UED5) and Course (e.g. BUS 940, BUS 501) from SIS ID or Course Name."""
    s_clean = str(sis_id).strip()

    # Cohort Extraction (Supports UC5, UC6, UB1, UED10, etc.)
    cohort_match = re.search(r"(UC|UB|UED)\s*(\d+)", s_clean, re.IGNORECASE)
    if cohort_match:
        cohort = f"{cohort_match.group(1).upper()}{cohort_match.group(2)}"
    else:
        cohort = "N/A"

    # Course Code Extraction (e.g. BUS 940, BUS 501)
    course_match = re.search(r"([a-zA-Z]{3})[- _]?([0-9]{3})", s_clean)
    if course_match:
        course_code = f"{course_match.group(1).upper()} {course_match.group(2)}"
    else:
        clean_match = re.search(r"([a-zA-Z]{3})\s*([0-9]{3})", course_name_api)
        if clean_match:
            course_code = f"{clean_match.group(1).upper()} {clean_match.group(2)}"
        else:
            course_code = course_name_api

    return cohort, course_code


def collect_consolidated_rows(course_id, sis_id, course_name_api):
    cohort, course_code = get_cohort_and_course(sis_id, course_name_api)

    # [PERF] Single API call extracts both user_map AND staff_ids
    user_map, staff_ids = get_course_users_with_emails(course_id)

    print("  Fetching modules...")
    modules = get_course_modules(course_id)
    topic_module_map = {}

    def fetch_items(m):
        items = get_module_items(course_id, m["id"])
        return m.get("name", "Unknown Module"), items

    # Parallelize module fetching for speed
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(fetch_items, m) for m in modules]
        for f in as_completed(futures):
            mod_name, items = f.result()
            for item in items:
                if item.get("type") == "Discussion":
                    topic_module_map[item.get("content_id")] = mod_name

    print("  Fetching discussion topics...")
    topics = get_discussion_topics(course_id)
    flat_rows = []

    for topic in topics:
        topic_title = topic.get("title", "Untitled Topic")
        topic_module_name = topic_module_map.get(topic["id"])

        # Check if topic was created by a student/learner
        topic_author = topic.get("user_name") or (topic.get("author") or {}).get("display_name")
        topic_author_id = (topic.get("author") or {}).get("id") or topic.get("user_id")

        topic_author_info = user_map.get(topic_author_id, {})
        topic_author_name = topic_author_info.get("name") or topic_author or f"User#{topic_author_id}"
        topic_author_email = (topic_author_info.get("email") or "").strip().lower()

        is_student_created = False
        if topic_author_id and topic_author_id not in staff_ids:
            is_student_created = True

        participants, view = get_full_view(course_id, topic["id"])

        if is_student_created:
            # ────────────────────────────────────────────────────────
            # MODE A: Student-created Discussion Topic
            # ────────────────────────────────────────────────────────
            created_on = format_created_on(topic.get("created_at"))

            # Evaluate replies for Mode A
            assigned_info = sheets_helper.get_assigned_ta(course_code, topic_author_email, topic_author_info.get("student_id"))
            all_ta_emails = sheets_helper.get_all_ta_emails()
            assigned_ta_email = assigned_info["ta_email"].lower().strip() if (assigned_info and assigned_info.get("ta_email")) else None
            assigned_ta_name = assigned_info["ta_name"].strip() if (assigned_info and assigned_info.get("ta_name")) else None

            assigned_ta_replies = []
            other_ta_replies = []
            all_replies = []

            for r in extract_all_replies(view):
                rid      = r.get("user_id")
                r_p_info = participants.get(rid, {})
                r_u_info = user_map.get(rid, {})
                r_name   = r_u_info.get("name") or r_p_info.get("display_name") or f"User#{rid}"
                r_email  = (r_u_info.get("email") or r_p_info.get("email") or "").strip().lower()
                r_time   = r.get("created_at", "")

                reply_details = {
                    "name": r_name,
                    "email": r_email,
                    "time": r_time
                }

                if assigned_ta_email and r_email == assigned_ta_email:
                    assigned_ta_replies.append(reply_details)
                elif r_email in all_ta_emails:
                    other_ta_replies.append(reply_details)
                elif rid in staff_ids:
                    other_ta_replies.append(reply_details)
                all_replies.append(reply_details)

            raw_created_on = topic.get("created_at")

            # Business Tracker SLA threshold is 24 hours
            if assigned_info:
                if assigned_ta_replies:
                    sorted_rep = sorted(assigned_ta_replies, key=lambda x: x["time"])
                    raw_replied_on = sorted_rep[0]["time"]
                    first_responder = assigned_ta_name or sorted_rep[0]["name"]
                    replied_on = format_replied_on(raw_replied_on)
                    replied = "Yes"
                    duration = calculate_duration_hours(raw_created_on, raw_replied_on)
                    sla_status = "Delayed Response" if (duration is not None and duration > 24) else "On Time Response"
                    rep_names_str = assigned_ta_name or sorted_rep[0]["name"]
                    rep_emails_str = assigned_ta_email or sorted_rep[0]["email"]
                elif other_ta_replies:
                    sorted_rep = sorted(other_ta_replies, key=lambda x: x["time"])
                    raw_replied_on = sorted_rep[0]["time"]
                    other_name = sorted_rep[0]["name"]
                    first_responder = f"{other_name} (Other TA)"
                    replied_on = format_replied_on(raw_replied_on)
                    replied = "Yes (Replied by another TA)"
                    duration = calculate_duration_hours(raw_created_on, raw_replied_on)
                    sla_status = "Delayed Response (Other TA)" if (duration is not None and duration > 24) else "On Time Response (Other TA)"
                    rep_names_str = f"{other_name} (Other TA)"
                    rep_emails_str = sorted_rep[0]["email"]
                else:
                    replied = "No"
                    replied_on = ""
                    duration = ""
                    sla_status = "Pending Response"
                    first_responder = ""
                    rep_names_str = ""
                    rep_emails_str = ""
            else:
                # No TA assigned to learner in ImportData_TA
                if assigned_ta_replies or other_ta_replies:
                    sorted_rep = sorted(assigned_ta_replies or other_ta_replies, key=lambda x: x["time"])
                    raw_replied_on = sorted_rep[0]["time"]
                    first_responder = f"{sorted_rep[0]['name']} (Unassigned)"
                    replied_on = format_replied_on(raw_replied_on)
                    replied = "Yes (No TA Assigned)"
                    duration = calculate_duration_hours(raw_created_on, raw_replied_on)
                    sla_status = "Delayed Response" if (duration is not None and duration > 24) else "On Time Response"
                    rep_names_str = sorted_rep[0]["name"]
                    rep_emails_str = sorted_rep[0]["email"]
                else:
                    replied = "No (No TA Assigned)"
                    replied_on = ""
                    duration = ""
                    sla_status = "Pending (No TA Assigned)"
                    first_responder = "Unassigned"
                    rep_names_str = ""
                    rep_emails_str = ""

            flat_rows.append({
                "Cohort"            : cohort,
                "Course Name"       : course_code,
                "Topic"             : topic_title,
                "Learner Name"      : topic_author_name,
                "Student ID"        : topic_author_info.get("student_id", "N/A"),
                "Learner Email"     : topic_author_email,
                "Created On"        : created_on,
                "Replied"           : replied,
                "Replied On"        : replied_on,
                "Duration (Hours)"  : duration if duration is not None else "",
                "SLA Status"        : sla_status,
                "First Responder"   : first_responder,
                "Assigned TA (Needs to Reply)" : assigned_ta_name or "Unassigned",
                "Assigned TA Email": assigned_ta_email or ""
            })

        else:
            # ────────────────────────────────────────────────────────
            # MODE B: Instructor/Staff-created Discussion Topic
            # ────────────────────────────────────────────────────────
            if not view:
                continue

            for entry in view:
                uid = entry.get("user_id")

                # Filter out instructor/staff top-level posts
                if uid in staff_ids:
                    continue

                u_info = user_map.get(uid, {})
                l_name = u_info.get("name") or f"User#{uid}"
                l_email = (u_info.get("email") or "").strip().lower()
                l_sid = u_info.get("student_id", "N/A")

                raw_created = entry.get("created_at")
                created_on = format_created_on(raw_created)

                # Recursive extraction captures all nested threaded replies
                raw_replies = extract_all_replies(entry.get("replies", []))

                assigned_info = sheets_helper.get_assigned_ta(course_code, l_email, l_sid)
                all_ta_emails = sheets_helper.get_all_ta_emails()
                assigned_ta_email = assigned_info["ta_email"].lower().strip() if (assigned_info and assigned_info.get("ta_email")) else None
                assigned_ta_name = assigned_info["ta_name"].strip() if (assigned_info and assigned_info.get("ta_name")) else None

                assigned_ta_replies = []
                other_ta_replies = []
                all_replies = []

                for rp in raw_replies:
                    rid = rp.get("user_id")
                    ru_info = user_map.get(rid, {})
                    rp_name = ru_info.get("name") or f"User#{rid}"
                    rp_email = (ru_info.get("email") or "").strip().lower()

                    rep_det = {"name": rp_name, "email": rp_email, "time": rp.get("created_at")}
                    if assigned_ta_email and rp_email == assigned_ta_email:
                        assigned_ta_replies.append(rep_det)
                    elif rp_email in all_ta_emails:
                        other_ta_replies.append(rep_det)
                    elif rid in staff_ids:
                        other_ta_replies.append(rep_det)
                    all_replies.append(rep_det)

                # Business Tracker SLA threshold is 24 hours
                if assigned_info:
                    if assigned_ta_replies:
                        sorted_rep = sorted(assigned_ta_replies, key=lambda x: x["time"])
                        raw_replied = sorted_rep[0]["time"]
                        first_responder = assigned_ta_name or sorted_rep[0]["name"]
                        replied_on = format_replied_on(raw_replied)
                        replied = "Yes"
                        duration = calculate_duration_hours(raw_created, raw_replied)
                        sla_status = "Delayed Response" if (duration is not None and duration > 24) else "On Time Response"
                        rep_names_str = assigned_ta_name or sorted_rep[0]["name"]
                        rep_emails_str = assigned_ta_email or sorted_rep[0]["email"]
                    elif other_ta_replies:
                        sorted_rep = sorted(other_ta_replies, key=lambda x: x["time"])
                        raw_replied = sorted_rep[0]["time"]
                        other_name = sorted_rep[0]["name"]
                        first_responder = f"{other_name} (Other TA)"
                        replied_on = format_replied_on(raw_replied)
                        replied = "Yes (Replied by another TA)"
                        duration = calculate_duration_hours(raw_created, raw_replied)
                        sla_status = "Delayed Response (Other TA)" if (duration is not None and duration > 24) else "On Time Response (Other TA)"
                        rep_names_str = f"{other_name} (Other TA)"
                        rep_emails_str = sorted_rep[0]["email"]
                    else:
                        replied = "No"
                        replied_on = ""
                        duration = ""
                        sla_status = "Pending Response"
                        first_responder = ""
                        rep_names_str = ""
                        rep_emails_str = ""
                else:
                    if assigned_ta_replies or other_ta_replies:
                        sorted_rep = sorted(assigned_ta_replies or other_ta_replies, key=lambda x: x["time"])
                        raw_replied = sorted_rep[0]["time"]
                        first_responder = f"{sorted_rep[0]['name']} (Unassigned)"
                        replied_on = format_replied_on(raw_replied)
                        replied = "Yes (No TA Assigned)"
                        duration = calculate_duration_hours(raw_created, raw_replied)
                        sla_status = "Delayed Response" if (duration is not None and duration > 24) else "On Time Response"
                        rep_names_str = sorted_rep[0]["name"]
                        rep_emails_str = sorted_rep[0]["email"]
                    else:
                        replied = "No (No TA Assigned)"
                        replied_on = ""
                        duration = ""
                        sla_status = "Pending (No TA Assigned)"
                        first_responder = "Unassigned"
                        rep_names_str = ""
                        rep_emails_str = ""

                flat_rows.append({
                    "Cohort"            : cohort,
                    "Course Name"       : course_code,
                    "Topic"             : topic_title,
                    "Learner Name"      : l_name,
                    "Student ID"        : l_sid,
                    "Learner Email"     : l_email,
                    "Created On"        : created_on,
                    "Replied"           : replied,
                    "Replied On"        : replied_on,
                    "Duration (Hours)"  : duration if duration is not None else "",
                    "SLA Status"        : sla_status,
                    "First Responder"   : first_responder,
                    "Assigned TA (Needs to Reply)" : assigned_ta_name or "Unassigned",
                    "Assigned TA Email": assigned_ta_email or ""
                })

    return flat_rows



# ──────────────────────────────────────────────────────────────────
#  EXCEL SHEET BUILDER
# ──────────────────────────────────────────────────────────────────

C_HDR_BG   = "1F3864"   # dark navy
C_HDR_FG   = "FFFFFF"   # white
C_TITLE_BG = "2E75B6"   # slate blue
C_YES_BG   = "E2EFDA"   # soft green
C_YES_FG   = "375623"
C_NO_BG    = "FCE4D6"   # soft pink
C_NO_FG    = "C00000"
C_ALT_BG   = "F9FBFD"   # very light blue alternate
C_BORDER   = "D9D9D9"

_SIDE           = Side(border_style="thin", color=C_BORDER)
CACHED_BORDER   = Border(left=_SIDE, right=_SIDE, top=_SIDE, bottom=_SIDE)
CACHED_HDR_FONT = Font(bold=True, color=C_HDR_FG, size=10)
CACHED_HDR_FILL = PatternFill("solid", fgColor=C_HDR_BG)
CACHED_HDR_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

_FILL_CACHE = {}
_FONT_CACHE = {}
_ALIGN_CACHE = {}

def _get_fill(color):
    if color not in _FILL_CACHE:
        _FILL_CACHE[color] = PatternFill("solid", fgColor=color)
    return _FILL_CACHE[color]

def _get_font(bold, color, size=10):
    key = (bold, color, size)
    if key not in _FONT_CACHE:
        _FONT_CACHE[key] = Font(bold=bold, color=color, size=size)
    return _FONT_CACHE[key]

def _get_align(horizontal, vertical="center", wrap=True):
    key = (horizontal, vertical, wrap)
    if key not in _ALIGN_CACHE:
        _ALIGN_CACHE[key] = Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap)
    return _ALIGN_CACHE[key]

def thin_border():
    return CACHED_BORDER


def build_excel_sheet(wb, flat_rows, sis_id, course_name):
    # Sanitize sheet title (Excel doesn't allow: : \ / ? * [ ])
    clean_title = re.sub(r'[:\\\/\?\*\[\]]', '_', sis_id)
    safe_title = clean_title[:31]

    if "Sheet" in wb.sheetnames and len(wb.sheetnames) == 1 and wb.active.title == "Sheet":
        ws = wb.active
        ws.title = safe_title
    else:
        ws = wb.create_sheet(title=safe_title)

    # ── Title Block ────────────────────────────────────────────────
    ws.merge_cells("A1:N1")
    t = ws["A1"]
    t.value     = f"EDGEWOOD DISCUSSION BOARD AUDIT  |  {course_name} ({sis_id})"
    t.font      = Font(bold=True, size=12, color=C_HDR_FG)
    t.fill      = PatternFill("solid", fgColor=C_TITLE_BG)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # ── Summary Analysis Table (ABOVE LEARNER DETAILS TABLE) ─────────
    valid_rows = [r for r in flat_rows if r.get("Learner Name") and r["Learner Name"] not in ("(No student posts yet)", "N/A", "Unknown", "User#None")]
    unique_learners = len(set(r["Learner Name"] for r in valid_rows))
    total_queries   = len(valid_rows)
    total_replied   = len([r for r in valid_rows if str(r.get("Replied", "")).startswith("Yes") or "replied" in str(r.get("Replied", "")).lower()])
    pending_queries = len([r for r in valid_rows if str(r.get("Replied", "")).startswith("No")])
    durations       = [r["Duration (Hours)"] for r in valid_rows if isinstance(r.get("Duration (Hours)"), (int, float))]
    avg_course_time = round(sum(durations)/len(durations), 2) if durations else 0

    within_sla  = len([r for r in valid_rows if "On Time" in str(r.get("SLA Status", ""))])
    outside_sla = len([r for r in valid_rows if "Delayed" in str(r.get("SLA Status", ""))])

    # Pending queries average elapsed duration (hours)
    now_dt = datetime.now()
    pending_durs = []
    for r in valid_rows:
        if str(r.get("Replied", "")).startswith("No"):
            created_dt = sheets_helper.parse_date_string(r.get("Created On"))
            if created_dt:
                p_hrs = max(0.0, round((now_dt - created_dt).total_seconds() / 3600.0, 1))
                pending_durs.append(p_hrs)
    avg_pending_time = round(sum(pending_durs)/len(pending_durs), 2) if pending_durs else 0

    # TA Breakdown Calculation
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

    # Table 1: Course Overview (Rows 3 to 6)
    sum_start = 3
    ws.merge_cells(f"A{sum_start}:F{sum_start}")
    sh1 = ws[f"A{sum_start}"]; sh1.value = "COURSE OVERVIEW"; sh1.font = Font(bold=True, color="FFFFFF"); sh1.fill = PatternFill("solid", fgColor="4472C4"); sh1.alignment = Alignment(horizontal="center")
    for col in range(1, 7): ws.cell(row=sum_start, column=col).border = thin_border()

    # Category Headers
    c_tp = ws.cell(row=sum_start+1, column=1, value="Total Post"); c_tp.font = Font(bold=True, color="FFFFFF"); c_tp.fill = PatternFill("solid", fgColor="1F3864"); c_tp.alignment = Alignment(horizontal="center"); c_tp.border = thin_border()
    ws.merge_cells(f"B{sum_start+1}:D{sum_start+1}")
    c_tr = ws.cell(row=sum_start+1, column=2, value="Total Response"); c_tr.font = Font(bold=True, color="375623"); c_tr.fill = PatternFill("solid", fgColor="E2EFDA"); c_tr.alignment = Alignment(horizontal="center")
    for col in range(2, 5): ws.cell(row=sum_start+1, column=col).border = thin_border()
    ws.merge_cells(f"E{sum_start+1}:F{sum_start+1}")
    c_pr = ws.cell(row=sum_start+1, column=5, value="Pending Response"); c_pr.font = Font(bold=True, color="C00000"); c_pr.fill = PatternFill("solid", fgColor="FCE4D6"); c_pr.alignment = Alignment(horizontal="center")
    for col in range(5, 7): ws.cell(row=sum_start+1, column=col).border = thin_border()

    # Sub-Headers
    sub_hdrs = ["Total Queries", "Within SLA", "Outside SLA", "Avg. Response HR", "Pending Count", "Avg"]
    for ci, h in enumerate(sub_hdrs, start=1):
        sc = ws.cell(row=sum_start+2, column=ci, value=h)
        sc.font = Font(bold=True, size=9); sc.fill = PatternFill("solid", fgColor="F2F2F2"); sc.alignment = Alignment(horizontal="center"); sc.border = thin_border()

    # Data Values
    data_vals = [total_queries, within_sla, outside_sla, avg_course_time, pending_queries, avg_pending_time]
    for ci, val in enumerate(data_vals, start=1):
        dc = ws.cell(row=sum_start+3, column=ci, value=val)
        dc.font = Font(bold=(ci==1 or ci==5)); dc.alignment = Alignment(horizontal="center"); dc.border = thin_border()

    # Row 7 is a spacer row
    ws.row_dimensions[7].height = 12

    # ── Column Headers for Learner Details Table (Row 8) ───────────
    hdr_row = 8
    headers = [
        "Cohort", "Course Name", "Topic", "Learner Name", "Student ID", "Learner Email", 
        "Created On", "Replied", "Replied On", "Duration (Hours)", "SLA Status", 
        "First Responder", "Assigned TA (Needs to Reply)", "Assigned TA Email"
    ]
    col_widths = [10, 15, 30, 20, 15, 25, 18, 10, 20, 15, 18, 20, 28, 30]

    for ci, (h, w) in enumerate(zip(headers, col_widths), start=1):
        cell           = ws.cell(row=hdr_row, column=ci, value=h)
        cell.font      = CACHED_HDR_FONT
        cell.fill      = CACHED_HDR_FILL
        cell.alignment = CACHED_HDR_ALIGN
        cell.border    = CACHED_BORDER
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.row_dimensions[hdr_row].height = 24

    # ── Data Row Injection (Row 9+) ────────────────────────────────
    if not flat_rows:
        ws.merge_cells("A9:N9")
        empty = ws["A9"]
        empty.value     = "No discussion boards or learner posts found in this course."
        empty.font      = Font(italic=True, color="808080")
        empty.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[9].height = 24
        last_row = 9
    else:
        for ri, r in enumerate(flat_rows, start=9):
            val_replied = str(r.get("Replied", "")).strip()
            val_sla     = str(r.get("SLA Status", "")).strip()
            is_yes      = val_replied == "Yes"
            is_no       = val_replied.startswith("No")
            is_other    = "Other TA" in val_replied or "Replied by another TA" in val_replied
            is_delayed  = "Delayed" in val_sla
            is_ontime   = "On Time" in val_sla
            alt_bg      = C_ALT_BG if ri % 2 == 0 else "FFFFFF"

            if is_other:
                replied_bg = "FFF2CC"
                replied_fg = "7F6000"
            elif is_yes:
                replied_bg = C_YES_BG
                replied_fg = C_YES_FG
            elif is_no:
                replied_bg = C_NO_BG
                replied_fg = C_NO_FG
            else:
                replied_bg = alt_bg
                replied_fg = "7F8C8D"

            sla_bg = alt_bg
            sla_fg = "000000"
            if is_delayed:
                sla_bg = "FFC7CE"; sla_fg = "9C0006"
            elif is_ontime:
                sla_bg = "C6EFCE"; sla_fg = "006100"

            row_cells = [
                (r["Cohort"],             "center", alt_bg, "000000", False),
                (r["Course Name"],        "center", alt_bg, "000000", False),
                (r["Topic"],              "left",   alt_bg, "000000", False),
                (r["Learner Name"],       "left",   alt_bg, "000000", False),
                (r["Student ID"],         "center", alt_bg, "000000", False),
                (r["Learner Email"],      "left",   alt_bg, "000000", False),
                (r["Created On"],         "center", alt_bg, "000000", False),
                (r["Replied"],            "center", replied_bg, replied_fg, True),
                (r["Replied On"],         "center", alt_bg, "000000", False),
                (r["Duration (Hours)"],   "center", alt_bg, "000000", False),
                (r["SLA Status"],         "center", sla_bg, sla_fg, True),
                (r["First Responder"],    "left",   alt_bg, "000000", True),
                (r.get("Assigned TA (Needs to Reply)", r.get("Replied By (Name)", "")),  "left",   alt_bg, "000000", False),
                (r.get("Assigned TA Email", r.get("Replied By (Email)", "")), "left",   alt_bg, "000000", False)
            ]

            for ci, (val, align, bg, fg, bold) in enumerate(row_cells, start=1):
                cell           = ws.cell(row=ri, column=ci, value=val)
                cell.font      = _get_font(bold, fg)
                cell.fill      = _get_fill(bg)
                cell.alignment = _get_align(align)
                cell.border    = CACHED_BORDER

            ws.row_dimensions[ri].height = 20
            last_row = ri

    # ── Table 2: TA Performance Breakdown (below learner rows) ─────
    ta_start = last_row + 2
    ws.merge_cells(f"A{ta_start}:G{ta_start}")
    sh2 = ws[f"A{ta_start}"]; sh2.value = "TA PERFORMANCE BREAKDOWN (TREND ANALYSIS)"; sh2.font = Font(bold=True, color="FFFFFF"); sh2.fill = PatternFill("solid", fgColor="C00000"); sh2.alignment = Alignment(horizontal="center")
    for col in range(1, 8): ws.cell(row=ta_start, column=col).border = thin_border()
    
    ta_hdrs = ["TA Name", "Pending Response", "Replies", "On-Time", "Delayed", "Avg Time (Hrs)", "SLA Violation %"]
    for ci, h in enumerate(ta_hdrs, start=1):
        c = ws.cell(row=ta_start+1, column=ci, value=h)
        c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="D9D9D9"); c.alignment = Alignment(horizontal="center"); c.border = thin_border()

    for i, (name, s) in enumerate(ta_stats.items(), start=2):
        avg_t = round(sum(s["durs"])/len(s["durs"]), 2) if s["durs"] else 0
        sla_v = round((s["delayed"]/s["replies"])*100, 1) if s["replies"] > 0 else 0
        
        cells = [
            (name, "left"),
            (s["pending"], "center"),
            (s["replies"], "center"),
            (s["ontime"], "center"),
            (s["delayed"], "center"),
            (avg_t, "center"),
            (f"{sla_v}%", "center")
        ]
        for ci, (v, a) in enumerate(cells, start=1):
            c = ws.cell(row=ta_start+i, column=ci, value=v)
            c.alignment = Alignment(horizontal=a); c.border = thin_border()

    ws.freeze_panes = "A9"


# ──────────────────────────────────────────────────────────────────
#  MAIN RUNNER
# ──────────────────────────────────────────────────────────────────

def main():
    print("\n" + "="*65)
    print("  EDGEWOOD CANVAS -- ENHANCED DISCUSSION AUDIT (SLA + MULTI-TAB)")
    print("="*65)

    if not check_connection():
        sys.exit(1)

    if len(sys.argv) > 1:
        raw_input = " ".join(sys.argv[1:])
    else:
        print("\n  Enter Course SIS IDs (separated by space or comma):")
        print("  Example: BUS-940-UC6--Fall-2026, BUS-902-UC7--Fall-2026")
        raw_input = input("\n  > ").strip()

    if not raw_input:
        print("  [!] No SIS ID entered. Exiting.")
        sys.exit(1)

    if "," in raw_input:
        sis_ids = [s.strip() for s in raw_input.split(",") if s.strip()]
    else:
        sis_ids = [s.strip() for s in raw_input.split() if s.strip()]
    
    if not sis_ids:
        print("  [!] No valid SIS IDs found.")
        sys.exit(1)

    wb = openpyxl.Workbook()
    processed_count = 0

    for sis_id in sis_ids:
        print(f"\n" + "-"*50)
        print(f"  Processing Course: {sis_id} ...")
        
        course = find_course(sis_id)
        if not course:
            print(f"  [!] Course not found for: '{sis_id}'. Skipping.")
            continue

        course_id   = course["id"]
        course_name = course.get("name", "Unknown Course")
        print(f"  Found --> [{course_id}] {course_name}")

        flat_rows = collect_consolidated_rows(course_id, sis_id, course_name)
        has_real_data = any(r.get("Learner Name") not in ("(No student posts yet)", "N/A", "Unknown", "User#None", "") for r in flat_rows)

        if not flat_rows or not has_real_data:
            print(f"  [!] No student discussions found in this course. Skipping tab.")
            continue

        build_excel_sheet(wb, flat_rows, sis_id, course_name)
        processed_count += 1

    if processed_count == 0:
        print("\n  [!] No courses were successfully processed. Excel not saved.")
        sys.exit(1)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    fname = os.path.join(SCRIPT_DIR, f"MultiCourse_Discussion_Audit_{timestamp}.xlsx")
    
    wb.save(fname)
    print(f"\n" + "="*65)
    print(f"  Excel saved --> {fname}")
    print(f"  Total Courses Processed: {processed_count}")
    print("  Done!\n")

    try:
        os.startfile(fname)
    except Exception:
        pass


if __name__ == "__main__":
    main()