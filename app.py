import os
import smtplib
import subprocess
import sys
import tempfile
import uuid
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

STATE = {
    "records": [],
    "summary": {},
    "uploaded": False,
}


def reset_state():
    STATE["records"] = []
    STATE["summary"] = {}
    STATE["uploaded"] = False


def safe_file_name(name):
    return "".join(c for c in name if c.isalnum() or c in (" ", "-", "_")).strip()


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


def send_email(record, use_outlook_windows, use_outlook_mac, outlook, smtp_server, smtp_from):
    subject = f"BV Report - {record['office_name']}"
    body = f"""Hi,

Please find attached BV Report for {record['office_name']}.

Regards,
Mushtaq Memon
"""
    attachment_bytes = to_excel_bytes(record["slice_rows"])
    attachment_name = f"{record['safe_name']}.xlsx"

    if use_outlook_windows and outlook is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(attachment_bytes)
            temp_path = tmp.name
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
                os.remove(temp_path)
            except OSError:
                pass
    elif use_outlook_mac:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(attachment_bytes)
            temp_path = tmp.name
        try:
            send_via_outlook_mac(record["to"], record["cc"], subject, body, temp_path)
        finally:
            try:
                os.remove(temp_path)
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
    email_df = pd.read_excel(BytesIO(email_bytes))
    data_df = pd.read_excel(BytesIO(data_bytes), sheet_name="All Agent Data")
    data_df.columns = data_df.columns.str.strip()
    email_df.columns = email_df.columns.str.strip()

    records = []
    skipped_no_doctor = 0
    skipped_no_to = 0
    skipped_no_data = 0

    for _, row in email_df.iterrows():
        doctor = str(row["Doctor Name"]).strip()
        to_email_list = clean_recipients(row["BV Report Send To"])
        cc_email_list = clean_recipients(row["BV Report Send CC"])
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
            record["office_name"] = "(blank)"
            record["status"] = "skipped"
            record["message"] = "No doctor name"
            records.append(record)
            continue

        if not to_email_list:
            skipped_no_to += 1
            record["status"] = "skipped"
            record["message"] = "No valid To email"
            records.append(record)
            continue

        filtered_df = data_df[data_df["Office Name"].str.strip() == doctor]
        if filtered_df.empty:
            skipped_no_data += 1
            record["status"] = "skipped"
            record["message"] = "No matching data in Office Name"
            records.append(record)
            continue

        clean_name = safe_file_name(doctor.replace(".", "").replace("&", "and")) or "Report"
        record["can_send"] = True
        record["safe_name"] = clean_name
        record["row_count"] = len(filtered_df)
        record["slice_rows"] = filtered_df.fillna("").to_dict(orient="records")
        records.append(record)

    summary = {
        "rows": len(email_df),
        "sent": 0,
        "failed": 0,
        "skipped_no_doctor": skipped_no_doctor,
        "skipped_no_to": skipped_no_to,
        "skipped_no_data": skipped_no_data,
    }
    return records, summary


def send_records(record_ids=None):
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
        for record in STATE["records"]:
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

        STATE["summary"]["sent"] += sent_count
        STATE["summary"]["failed"] += failed_count
        return STATE["summary"]
    finally:
        if smtp_server is not None:
            smtp_server.quit()


@app.get("/")
def home():
    return render_template(
        "index.html",
        records=STATE["records"],
        summary=STATE["summary"],
        uploaded=STATE["uploaded"],
    )


@app.post("/upload")
def upload_files():
    try:
        email_upload = request.files.get("email_file")
        data_upload = request.files.get("data_file")
        if not email_upload or not data_upload:
            flash("Please upload both files.", "error")
            return redirect(url_for("home"))

        email_bytes = email_upload.read()
        data_bytes = data_upload.read()
        records, summary = build_records(email_bytes, data_bytes)
        STATE["records"] = records
        STATE["summary"] = summary
        STATE["uploaded"] = True

        flash("Files uploaded and sliced reports generated.", "success")
        return redirect(url_for("home"))
    except Exception as e:
        flash(f"Upload failed: {e}", "error")
        return redirect(url_for("home"))


@app.post("/send-all")
def send_all():
    if not STATE["uploaded"]:
        flash("Upload both files first.", "error")
        return redirect(url_for("home"))
    try:
        send_records()
        flash("Send All completed.", "success")
    except Exception as e:
        flash(f"Send All failed: {e}", "error")
    return redirect(url_for("home"))


@app.post("/reset")
def reset_app():
    try:
        reset_state()
        flash("App reset completed.", "success")
    except Exception as e:
        flash(f"Reset failed: {e}", "error")
    return redirect(url_for("home"))


@app.post("/send/<record_id>")
def send_one(record_id):
    if not STATE["uploaded"]:
        flash("Upload both files first.", "error")
        return redirect(url_for("home"))

    record = next((r for r in STATE["records"] if r["id"] == record_id), None)
    if not record:
        flash("Record not found.", "error")
        return redirect(url_for("home"))
    if not record["can_send"]:
        flash("This record cannot be sent.", "error")
        return redirect(url_for("home"))

    try:
        send_records(record_ids={record_id})
        flash(f"Sent: {record['office_name']}", "success")
    except Exception as e:
        flash(f"Failed to send {record['office_name']}: {e}", "error")
    return redirect(url_for("home"))


@app.get("/preview/<record_id>")
def preview_slice(record_id):
    record = next((r for r in STATE["records"] if r["id"] == record_id), None)
    if not record or not record.get("slice_rows"):
        return jsonify({"ok": False, "error": "Sliced file not found."}), 404

    try:
        df = pd.DataFrame(record["slice_rows"])
        rows = df.head(200).fillna("").to_dict(orient="records")
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