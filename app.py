import os
import io
import re
import time
import uuid
import threading
import builtins
from datetime import datetime
import openpyxl
from flask import Flask, render_template, request, send_file, jsonify, redirect, url_for
import Disussion_Automate as da
import sheets_helper

app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'edgewood_secret_key_123')

# Global state for live updates
task_progress = {}
task_results = {}
current_task_id = None

# Intercept print to stream live updates natively
original_print = builtins.print

def custom_print(*args, **kwargs):
    global current_task_id
    text = " ".join(str(a) for a in args)
    if current_task_id and text.strip():
        clean_text = text.strip()
        # Make the prints look cooler and more readable for the UI
        if "Processing thread" in clean_text:
            clean_text = clean_text.replace("Processing thread:", "Analyzing Topic:")
        elif "Fetching" in clean_text:
            clean_text = f"Connecting to Canvas: {clean_text}"
            
        # Ignore empty lines or simple separators
        if not set(clean_text).issubset({'-', '=', ' '}):
            task_progress[current_task_id] = clean_text
            
    original_print(*args, **kwargs)

builtins.print = custom_print

def run_audit(task_id, course_codes_input):
    global current_task_id
    current_task_id = task_id
    task_progress[task_id] = "Initializing Canvas API Connection..."
    
    try:
        if "," in course_codes_input:
            sis_ids = [s.strip() for s in course_codes_input.split(",") if s.strip()]
        else:
            sis_ids = [s.strip() for s in course_codes_input.split() if s.strip()]

        if not sis_ids:
            task_progress[task_id] = "ERROR: No valid SIS IDs found in input."
            return

        if not da.check_connection():
            task_progress[task_id] = "ERROR: API connection failed. Check your token."
            return

        # Fetch latest TA directory from Google Sheet to ensure new TAs are immediately active
        task_progress[task_id] = "Syncing Active TA Directory from Google Sheet..."
        sheets_helper.get_import_data_ta(force_refresh=True)

        wb = openpyxl.Workbook()
        processed_count = 0

        for sis_id in sis_ids:
            # Business Tracker specific filter
            clean_sis = sis_id.strip()
            if not clean_sis.upper().startswith("BUS"):
                task_progress[task_id] = f"Skipping {sis_id} (Not a Business course)."
                time.sleep(1)
                continue
            sis_id = clean_sis

            task_progress[task_id] = f"Locating Course: {sis_id}..."
            course = da.find_course(sis_id)
            if not course:
                task_progress[task_id] = f"Could not find Course: {sis_id}. Skipping..."
                time.sleep(1)
                continue
                
            course_id = course["id"]
            course_name = course.get("name", "Unknown Course")
            
            # This calls the original function which has all the print statements we intercept!
            flat_rows = da.collect_consolidated_rows(course_id, sis_id, course_name)
            
            has_real_data = any(r.get("Learner Name") not in ("(No student posts yet)", "N/A", "Unknown", "User#None", "") for r in flat_rows)

            if not flat_rows or not has_real_data:
                task_progress[task_id] = f"No discussions found in {sis_id}. Skipping..."
                time.sleep(1)
                continue

            task_progress[task_id] = "Compiling Data into Excel Structure..."
            da.build_excel_sheet(wb, flat_rows, sis_id, course_name)
            try:
                task_progress[task_id] = "Syncing Report to Google Sheet Database..."
                sheets_helper.push_to_google_sheet(flat_rows, sis_id, course_name)
                # Bust the dashboard cache so next fetch reflects new audit data
                sheets_helper._dashboard_cache["data"] = None
            except Exception as e:
                original_print(f"Error pushing to Google Sheets: {e}")
            processed_count += 1

        if processed_count == 0:
            task_progress[task_id] = "ERROR: No student discussions found in any course."
            return

        if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
            del wb["Sheet"]

        task_progress[task_id] = "Finalizing Output File..."
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        task_results[task_id] = {
            "file": output,
            "filename": f"MultiCourse_Discussion_Audit_{sis_ids[0].replace(' ', '_')}.xlsx"
        }
        
        time.sleep(0.5) # Slight pause to let UI catch up
        task_progress[task_id] = "COMPLETE"
    except Exception as e:
        task_progress[task_id] = f"ERROR: {str(e)}"
    finally:
        current_task_id = None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/start', methods=['POST'])
def start_audit():
    data = request.json
    course_codes_input = data.get('course_codes', '')
    if not course_codes_input:
        return jsonify({"error": "Please enter at least one course code."}), 400
        
    task_id = str(uuid.uuid4())
    task_progress[task_id] = "Warming up Audit Engine..."
    
    # Run the scraping in a background thread
    thread = threading.Thread(target=run_audit, args=(task_id, course_codes_input))
    thread.start()
    
    return jsonify({"task_id": task_id})

@app.route('/status/<task_id>')
def status(task_id):
    state = task_progress.get(task_id, "Unknown Task")
    return jsonify({"status": state})

@app.route('/download/<task_id>')
def download(task_id):
    if task_id in task_results:
        res = task_results[task_id]
        return send_file(
            res["file"],
            as_attachment=True,
            download_name=res["filename"],
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return "File not found or expired.", 404

@app.route('/api/dashboard')
def api_dashboard():
    try:
        force_refresh = request.args.get('refresh', 'false').lower() == 'true'
        data = sheets_helper.get_dashboard_data(force_refresh=force_refresh)
        return jsonify(data)
    except Exception as e:
        original_print(f"Dashboard data fetch failed: {e}")
        return jsonify({"error": str(e)}), 500

import socket
import smtplib
import ssl

class IPv4SMTP_SSL(smtplib.SMTP_SSL):
    """SMTP_SSL client forcing IPv4 (AF_INET) to prevent IPv6 ENETUNREACH [Errno 101] errors in cloud containers."""
    def _get_socket(self, host, port, timeout):
        infos = socket.getaddrinfo(host, port, socket.AF_INET, socket.SOCK_STREAM)
        if not infos:
            raise OSError(f"Could not resolve IPv4 address for {host}")
        last_err = None
        for res in infos:
            af, socktype, proto, canonname, sa = res
            s = socket.socket(af, socktype, proto)
            s.settimeout(timeout)
            try:
                s.connect(sa)
                return self.context.wrap_socket(s, server_hostname=self._host)
            except OSError as e:
                last_err = e
                s.close()
        if last_err:
            raise last_err
        raise OSError(f"Failed to connect to {host}:{port}")

class IPv4SMTP(smtplib.SMTP):
    """SMTP client forcing IPv4 (AF_INET) for port 587 STARTTLS."""
    def _get_socket(self, host, port, timeout):
        infos = socket.getaddrinfo(host, port, socket.AF_INET, socket.SOCK_STREAM)
        if not infos:
            raise OSError(f"Could not resolve IPv4 address for {host}")
        last_err = None
        for res in infos:
            af, socktype, proto, canonname, sa = res
            s = socket.socket(af, socktype, proto)
            s.settimeout(timeout)
            try:
                s.connect(sa)
                return s
            except OSError as e:
                last_err = e
                s.close()
        if last_err:
            raise last_err
        raise OSError(f"Failed to connect to {host}:{port}")

def get_google_smtp_session(sender_email, app_password, timeout=12):
    """
    Connects and logs into Google SMTP with multiple fallback strategies:
    1. Forces IPv4 (AF_INET) to bypass container IPv6 network unreachable bugs.
    2. Tries SSL on Port 465 first.
    3. If Port 465 fails, falls back to Port 587 with STARTTLS.
    4. Diagnoses cloud firewall blocks (e.g., Render Free Tier blocking outbound SMTP ports 465/587).
    """
    context = ssl.create_default_context()
    first_err = None

    # Strategy 1: Port 465 (SSL)
    try:
        server = IPv4SMTP_SSL("smtp.gmail.com", 465, context=context, timeout=timeout)
        server.login(sender_email, app_password)
        return server
    except (smtplib.SMTPAuthenticationError, smtplib.SMTPServerDisconnected):
        raise
    except Exception as e:
        first_err = e

    # Strategy 2: Port 587 (STARTTLS)
    try:
        server = IPv4SMTP("smtp.gmail.com", 587, timeout=timeout)
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(sender_email, app_password)
        return server
    except (smtplib.SMTPAuthenticationError, smtplib.SMTPServerDisconnected):
        raise
    except Exception as e:
        err_msg = str(e) or str(first_err) or ""
        # Check if error is network unreachable or timeout (typical Render free tier block)
        if any(keyword in err_msg.lower() for keyword in ["101", "unreachable", "timed out", "timeout", "refused", "connection reset"]):
            raise RuntimeError(
                "Render Free Tier blocks outbound SMTP ports (465/587) to prevent spam. "
                "To fix: Upgrade your Render service to 'Starter' ($7/mo) to unlock SMTP, "
                "or run the tracker locally on your PC (python app.py) where SMTP is fully supported."
            )
        raise RuntimeError(f"SMTP connection error: {err_msg}")

@app.route('/api/verify-credentials', methods=['POST'])
def verify_credentials():
    data = request.json or {}
    sender_email = (data.get('sender_email') or '').strip()
    app_password = (data.get('app_password') or '').strip().replace(' ', '')
    if not sender_email or not app_password:
        return jsonify({"valid": False, "error": "Email and App Password are required."}), 400

    try:
        server = get_google_smtp_session(sender_email, app_password, timeout=10)
        try:
            server.quit()
        except Exception:
            pass
        return jsonify({"valid": True, "message": "Credentials verified via Google SMTP."})
    except (smtplib.SMTPAuthenticationError, smtplib.SMTPServerDisconnected):
        return jsonify({"valid": False, "error": "Authentication failed. Check your email or 16-character Google App Password."}), 401
    except Exception as e:
        err_str = str(e)
        is_cloud_block = "Render" in err_str or "unreachable" in err_str.lower()
        return jsonify({
            "valid": False,
            "error": err_str,
            "is_cloud_block": is_cloud_block
        }), 500

@app.route('/api/scan-pending-ta-emails', methods=['POST'])
def scan_pending_ta_emails():
    data = request.json or {}
    course_shell = (data.get('course_shell') or '').strip()
    if not course_shell:
        return jsonify({"error": "Please provide a valid Course Shell ID."}), 400

    try:
        summary = sheets_helper.get_pending_ta_summary(course_shell)
        return jsonify(summary)
    except Exception as e:
        original_print(f"Error scanning pending TA emails: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/api/send-ta-email', methods=['POST'])
def send_ta_email():
    data = request.json or {}
    sender_email  = (data.get('sender_email') or '').strip()
    app_password  = (data.get('app_password') or '').strip().replace(' ', '')
    to_email      = (data.get('to_email') or '').strip()
    cc_emails_raw = (data.get('cc_emails') or '').strip()
    subject       = (data.get('subject') or '').strip()
    body_html     = data.get('body_html', '')
    body_text     = data.get('body_text', '')
    ta_name       = (data.get('ta_name') or 'TA').strip()
    is_test       = bool(data.get('is_test', False))

    if not sender_email or not app_password:
        return jsonify({"success": False, "error": "Sender email and App Password must be configured in Settings."}), 400
    if not to_email:
        return jsonify({"success": False, "error": "Recipient TA email is required."}), 400

    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.utils import formataddr

    # ── Test Mode: redirect to sender's own inbox without altering email body ──
    actual_to = sender_email if is_test else to_email
    actual_cc = []           if is_test else [c.strip() for c in cc_emails_raw.replace(';', ',').split(',') if c.strip()]
    course_shell  = (data.get('course_shell') or '').strip()
    default_sub   = "Gentle Reminder | Discussion Forum Queries"
    effective_subject = f"TEST_{subject or default_sub}" if is_test else (subject or default_sub)

    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = effective_subject
        sender_display = data.get('full_name') or "Edgewood Operations"
        msg['From'] = formataddr((sender_display, sender_email))
        msg['To'] = formataddr((ta_name, actual_to))
        if actual_cc:
            msg['Cc'] = ", ".join(actual_cc)

        if body_text:
            msg.attach(MIMEText(body_text, 'plain', 'utf-8'))
        if body_html:
            msg.attach(MIMEText(body_html, 'html', 'utf-8'))

        recipients = [actual_to] + actual_cc
        server = get_google_smtp_session(sender_email, app_password, timeout=15)
        server.sendmail(sender_email, recipients, msg.as_string())
        try:
            server.quit()
        except Exception:
            pass

        now_str = datetime.now().strftime("%I:%M:%S %p")
        return jsonify({
            "success": True,
            "is_test": is_test,
            "recipient": actual_to,
            "original_recipient": to_email,
            "cc": actual_cc,
            "sent_at": now_str,
            "message": f"[TEST] Email preview sent to {sender_email}" if is_test else f"Email successfully sent to {ta_name} ({to_email})"
        })
    except (smtplib.SMTPAuthenticationError, smtplib.SMTPServerDisconnected):
        return jsonify({"success": False, "error": "SMTP Authentication failed. Check your App Password."}), 401
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.getenv("PORT", 5000))
    app.run(debug=True, host='0.0.0.0', port=port, threaded=True)
