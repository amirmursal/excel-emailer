import os
import smtplib
import subprocess
import sys
import tempfile
import uuid
import shutil
from datetime import datetime
from io import BytesIO
from email.message import EmailMessage
from email.utils import parseaddr

import pandas as pd
from flask import Flask, flash, jsonify, redirect, render_template, request, url_for

try:
    import win32com.client as win32
except ImportError:
    win32 = None

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "email-sender-dev-secret")
MISC_TO_RECIPIENTS = ["mmemon@orthosynetics.com"] # later change to mmemon@orthosynetics.com
MISC_CC_RECIPIENTS = ["rafiks77@gmail.com"] # later change to rafiks77@gmail.com

WORKFLOWS = [
    ("bv-ortho-no-info", "BV – Ortho No Info Emailer"),
    ("bv-dental-no-info", "BV – Dental No Info Emailer"),
    ("ev-ortho-rstc", "EV – Ortho RSTC Emailer"),
    ("ev-dental-rstc", "EV – Dental RSTC Emailer"),
    ("bv-ortho-rstc", "BV – Ortho RSTC Emailer"),
    ("bv-dental-rstc", "BV – Dental RSTC Emailer"),
]
WORKFLOW_MAP = dict(WORKFLOWS)
DEFAULT_WORKFLOW = WORKFLOWS[0][0]
STATE_BY_WORKFLOW = {
    workflow_id: {
        "records": [],
        "summary": {},
        "uploaded": False,
    }
    for workflow_id, _ in WORKFLOWS
}


def get_workflow_or_default(workflow_id):
    if workflow_id in WORKFLOW_MAP:
        return workflow_id
    return DEFAULT_WORKFLOW


def reset_state(workflow_id):
    state = STATE_BY_WORKFLOW[workflow_id]
    state["records"] = []
    state["summary"] = {}
    state["uploaded"] = False


def safe_file_name(name):
    return "".join(c for c in name if c.isalnum() or c in (" ", "-", "_")).strip()


def normalize_column_name(name):
    return "".join(ch for ch in str(name).strip().lower() if ch.isalnum())


def find_column(df, expected_name):
    expected_normalized = normalize_column_name(expected_name)
    for column in df.columns:
        if normalize_column_name(column) == expected_normalized:
            return column
    raise ValueError(
        f"Required column '{expected_name}' not found. Available columns: {', '.join(map(str, df.columns))}"
    )


def clean_recipients(raw_value):
    if pd.isna(raw_value):
        return []
    recipients = []
    for item in str(raw_value).replace(";", ",").split(","):
        candidate = item.strip()
        if not candidate or candidate.lower() == "nan":
            continue

        _, parsed_email = parseaddr(candidate)
        email = parsed_email.strip() if parsed_email else candidate
        if "@" in email:
            recipients.append(email)
    return recipients


def send_via_outlook_mac(to_email_list, cc_email_list, subject, body, file_path):
    script = """
on run argv
	set toList to item 1 of argv
	set ccList to item 2 of argv
	set theSubject to item 3 of argv
	set theBody to item 4 of argv
	set attachmentPath to item 5 of argv
	
	tell application "Microsoft Outlook"
		set newMessage to make new outgoing message with properties {subject:theSubject, content:theBody}
		tell newMessage
			if toList is not "" then
				set AppleScript's text item delimiters to ","
				repeat with recipientAddress in text items of toList
					set trimmedTo to my trim_text(recipientAddress as text)
					if trimmedTo is not "" then
						make new to recipient at end of to recipients with properties {email address:{address:trimmedTo}}
					end if
				end repeat
			end if
			
			if ccList is not "" then
				set AppleScript's text item delimiters to ","
				repeat with recipientAddress in text items of ccList
					set trimmedCc to my trim_text(recipientAddress as text)
					if trimmedCc is not "" then
						make new cc recipient at end of cc recipients with properties {email address:{address:trimmedCc}}
					end if
				end repeat
			end if
			
			make new attachment with properties {file:(POSIX file attachmentPath)}
			send
		end tell
	end tell
end run

on trim_text(theText)
	set tid to AppleScript's text item delimiters
	set AppleScript's text item delimiters to {" ", tab, return, linefeed}
	set textItems to text items of theText
	set AppleScript's text item delimiters to ""
	set trimmedText to textItems as text
	set AppleScript's text item delimiters to tid
	return trimmedText
end trim_text
"""
    subprocess.run(
        [
            "osascript",
            "-e",
            script,
            ",".join(to_email_list),
            ",".join(cc_email_list),
            subject,
            body,
            file_path,
        ],
        check=True,
    )


def get_smtp_server():
    smtp_host = os.getenv("SMTP_HOST")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER")
    smtp_password = os.getenv("SMTP_PASSWORD")
    smtp_from = os.getenv("SMTP_FROM", smtp_user or "")
    use_starttls = os.getenv("SMTP_STARTTLS", "true").lower() in {"1", "true", "yes"}

    if not smtp_host or not smtp_user or not smtp_password:
        raise RuntimeError(
            "SMTP config missing. Set SMTP_HOST, SMTP_USER, SMTP_PASSWORD "
            "(optional: SMTP_PORT, SMTP_FROM, SMTP_STARTTLS)."
        )

    smtp_server = smtplib.SMTP(smtp_host, smtp_port)
    smtp_server.ehlo()
    if use_starttls:
        smtp_server.starttls()
        smtp_server.ehlo()
    smtp_server.login(smtp_user, smtp_password)
    return smtp_server, smtp_from


def to_excel_bytes(rows):
    output = BytesIO()
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


def sanitize_for_json(value):
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.isoformat()
    return value


def sanitize_rows_for_json(rows):
    return [{key: sanitize_for_json(value) for key, value in row.items()} for row in rows]


def create_named_attachment_tempfile(attachment_name, attachment_bytes):
    temp_dir = tempfile.mkdtemp(prefix="excel-emailer-")
    temp_path = os.path.join(temp_dir, attachment_name)
    with open(temp_path, "wb") as temp_file:
        temp_file.write(attachment_bytes)
    return temp_path, temp_dir


def send_email(record, use_outlook_windows, use_outlook_mac, outlook, smtp_server, smtp_from):
    subject = f"BV Report - {record['office_name']}"
    body = f"""Hi,

Please find attached BV Report for {record['office_name']}.

Regards,
Mushtaq Memon
"""
    attachment_bytes = to_excel_bytes(record["slice_rows"])
    office_name_for_file = safe_file_name(record["office_name"]) or record.get("safe_name") or "Report"
    us_date_for_file = datetime.now().strftime("%m-%d-%Y")
    attachment_name = f"{office_name_for_file} {us_date_for_file}.xlsx"

    if use_outlook_windows and outlook is not None:
        temp_path, temp_dir = create_named_attachment_tempfile(attachment_name, attachment_bytes)
        try:
            mail = outlook.CreateItem(0)
            mail.To = "; ".join(record["to"])
            mail.CC = "; ".join(record["cc"])
            mail.Subject = subject
            mail.Body = body
            mail.Attachments.Add(temp_path)
            mail.Send()
        finally:
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except OSError:
                pass
    elif use_outlook_mac:
        temp_path, temp_dir = create_named_attachment_tempfile(attachment_name, attachment_bytes)
        try:
            send_via_outlook_mac(record["to"], record["cc"], subject, body, temp_path)
        finally:
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except OSError:
                pass
    else:
        message = EmailMessage()
        message["From"] = smtp_from
        message["To"] = ", ".join(record["to"])
        if record["cc"]:
            message["Cc"] = ", ".join(record["cc"])
        message["Subject"] = subject
        message.set_content(body)
        message.add_attachment(
            attachment_bytes,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=attachment_name,
        )
        smtp_server.send_message(message)


def build_records(email_bytes, data_bytes):
    # Always use first sheet, and auto-detect if files were uploaded in swapped order.
    first_df = pd.read_excel(BytesIO(email_bytes), sheet_name=0)
    second_df = pd.read_excel(BytesIO(data_bytes), sheet_name=0)
    first_df.columns = first_df.columns.str.strip()
    second_df.columns = second_df.columns.str.strip()

    def has_required_columns(df, required_names):
        try:
            for name in required_names:
                find_column(df, name)
            return True
        except ValueError:
            return False

    email_required = ["Doctor Name", "BV Report Send To", "BV Report Send CC"]
    data_required = ["Office Name"]
    first_is_email = has_required_columns(first_df, email_required)
    second_is_data = has_required_columns(second_df, data_required)
    second_is_email = has_required_columns(second_df, email_required)
    first_is_data = has_required_columns(first_df, data_required)

    if first_is_email and second_is_data:
        email_df, data_df = first_df, second_df
    elif second_is_email and first_is_data:
        email_df, data_df = second_df, first_df
    else:
        raise ValueError(
            "Could not identify Email and Data files. "
            "Email file must contain Doctor Name, BV Report Send To, BV Report Send CC; "
            "Data file must contain Office Name."
        )

    doctor_col = find_column(email_df, "Doctor Name")
    to_col = find_column(email_df, "BV Report Send To")
    cc_col = find_column(email_df, "BV Report Send CC")
    office_col = find_column(data_df, "Office Name")
    office_names = data_df[office_col].fillna("").astype(str).str.strip()
    doctor_names = email_df[doctor_col].fillna("").astype(str).str.strip()
    doctor_name_set = {name for name in doctor_names if name and name.lower() != "nan"}

    records = []
    skipped_no_doctor = 0
    skipped_no_to = 0
    skipped_no_data = 0

    for _, row in email_df.iterrows():
        doctor = str(row[doctor_col]).strip()
        to_email_list = clean_recipients(row[to_col])
        cc_email_list = clean_recipients(row[cc_col])
        record = {
            "id": str(uuid.uuid4()),
            "office_name": doctor,
            "safe_name": "",
            "to": to_email_list,
            "cc": cc_email_list,
            "status": "pending",
            "message": "",
            "row_count": 0,
            "slice_rows": [],
            "can_send": False,
        }

        if not doctor or doctor.lower() == "nan":
            skipped_no_doctor += 1
            continue

        if not to_email_list:
            skipped_no_to += 1
            continue

        filtered_df = data_df[office_names == doctor]
        if filtered_df.empty:
            skipped_no_data += 1
            continue

        clean_name = safe_file_name(doctor.replace(".", "").replace("&", "and")) or "Report"
        record["can_send"] = True
        record["safe_name"] = clean_name
        record["row_count"] = len(filtered_df)
        record["slice_rows"] = filtered_df.fillna("").to_dict(orient="records")
        records.append(record)

    unmatched_data_df = data_df[(~office_names.isin(doctor_name_set)) & (office_names != "")]
    unmatched_rows = unmatched_data_df.fillna("").to_dict(orient="records")
    if unmatched_rows:
        records.append(
            {
                "id": str(uuid.uuid4()),
                "office_name": "Miscellaneous (Non Matching)",
                "safe_name": "Misc-Non-Matching-Report",
                "to": MISC_TO_RECIPIENTS.copy(),
                "cc": MISC_CC_RECIPIENTS.copy(),
                "status": "pending",
                "message": "All non-matching records grouped here.",
                "row_count": len(unmatched_rows),
                "slice_rows": unmatched_rows,
                "can_send": True,
            }
        )

    summary = {
        "rows": len(email_df),
        "sent": 0,
        "failed": 0,
        "skipped_no_doctor": skipped_no_doctor,
        "skipped_no_to": skipped_no_to,
        "skipped_no_data": skipped_no_data,
    }
    return records, summary


def send_records(workflow_id, record_ids=None):
    state = STATE_BY_WORKFLOW[workflow_id]
    use_outlook_windows = os.name == "nt" and win32 is not None
    use_outlook_mac = sys.platform == "darwin"

    outlook = win32.Dispatch("outlook.application") if use_outlook_windows else None
    smtp_server = None
    smtp_from = ""

    if not use_outlook_windows and not use_outlook_mac:
        smtp_server, smtp_from = get_smtp_server()

    try:
        sent_count = 0
        failed_count = 0
        for record in state["records"]:
            if record_ids and record["id"] not in record_ids:
                continue

            if not record["can_send"]:
                continue

            try:
                send_email(
                    record, use_outlook_windows, use_outlook_mac, outlook, smtp_server, smtp_from
                )
                record["status"] = "sent"
                record["message"] = "Sent successfully"
                sent_count += 1
                print(f"[SENT] {record['office_name']} -> {', '.join(record['to'])}")
            except Exception as e:
                record["status"] = "failed"
                record["message"] = str(e)
                failed_count += 1
                print(f"[FAIL] {record['office_name']}: {e}")

        state["summary"]["sent"] += sent_count
        state["summary"]["failed"] += failed_count
        return state["summary"]
    finally:
        if smtp_server is not None:
            smtp_server.quit()


@app.get("/")
def home():
    workflow_id = get_workflow_or_default(request.args.get("workflow", DEFAULT_WORKFLOW))
    state = STATE_BY_WORKFLOW[workflow_id]
    return render_template(
        "index.html",
        records=state["records"],
        summary=state["summary"],
        uploaded=state["uploaded"],
        workflows=WORKFLOWS,
        current_workflow=workflow_id,
        current_workflow_label=WORKFLOW_MAP[workflow_id],
    )


@app.post("/upload/<workflow_id>")
def upload_files(workflow_id):
    workflow_id = get_workflow_or_default(workflow_id)
    state = STATE_BY_WORKFLOW[workflow_id]
    try:
        email_upload = request.files.get("email_file")
        data_upload = request.files.get("data_file")
        if not email_upload or not data_upload:
            flash("Please upload both files.", "error")
            return redirect(url_for("home"))

        email_bytes = email_upload.read()
        data_bytes = data_upload.read()
        records, summary = build_records(email_bytes, data_bytes)
        state["records"] = records
        state["summary"] = summary
        state["uploaded"] = True

        flash("Files uploaded and sliced reports generated.", "success")
        return redirect(url_for("home", workflow=workflow_id))
    except Exception as e:
        flash(f"Upload failed: {e}", "error")
        return redirect(url_for("home", workflow=workflow_id))


@app.post("/send-all/<workflow_id>")
def send_all(workflow_id):
    workflow_id = get_workflow_or_default(workflow_id)
    state = STATE_BY_WORKFLOW[workflow_id]
    if not state["uploaded"]:
        flash("Upload both files first.", "error")
        return redirect(url_for("home", workflow=workflow_id))
    try:
        send_records(workflow_id)
        flash("Send All completed.", "success")
    except Exception as e:
        flash(f"Send All failed: {e}", "error")
    return redirect(url_for("home", workflow=workflow_id))


@app.post("/reset/<workflow_id>")
def reset_app(workflow_id):
    workflow_id = get_workflow_or_default(workflow_id)
    try:
        reset_state(workflow_id)
        flash("App reset completed.", "success")
    except Exception as e:
        flash(f"Reset failed: {e}", "error")
    return redirect(url_for("home", workflow=workflow_id))


@app.post("/send/<workflow_id>/<record_id>")
def send_one(workflow_id, record_id):
    workflow_id = get_workflow_or_default(workflow_id)
    state = STATE_BY_WORKFLOW[workflow_id]
    if not state["uploaded"]:
        flash("Upload both files first.", "error")
        return redirect(url_for("home", workflow=workflow_id))

    record = next((r for r in state["records"] if r["id"] == record_id), None)
    if not record:
        flash("Record not found.", "error")
        return redirect(url_for("home", workflow=workflow_id))
    if not record["can_send"]:
        flash("This record cannot be sent.", "error")
        return redirect(url_for("home", workflow=workflow_id))

    try:
        send_records(workflow_id, record_ids={record_id})
        flash(f"Sent: {record['office_name']}", "success")
    except Exception as e:
        flash(f"Failed to send {record['office_name']}: {e}", "error")
    return redirect(url_for("home", workflow=workflow_id))


@app.get("/preview/<workflow_id>/<record_id>")
def preview_slice(workflow_id, record_id):
    workflow_id = get_workflow_or_default(workflow_id)
    state = STATE_BY_WORKFLOW[workflow_id]
    record = next((r for r in state["records"] if r["id"] == record_id), None)
    if not record or not record.get("slice_rows"):
        return jsonify({"ok": False, "error": "Sliced file not found."}), 404

    try:
        df = pd.DataFrame(record["slice_rows"])
        rows = sanitize_rows_for_json(df.head(200).to_dict(orient="records"))
        return jsonify(
            {
                "ok": True,
                "office_name": record["office_name"],
                "columns": list(df.columns),
                "rows": rows,
            }
        )
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5004, debug=True)