@app.route('/')
def home():
    return '✅ Supervisor AI Webhook is live!'

@app.route('/run-script', methods=['POST'])
def run_script():
    data = request.json
    print("Received data:", data)
    try:
        success = run_pipeline()
        if success:
            return jsonify({"status": "success", "message": "Webhook triggered and script executed!"})
        else:
            return jsonify({"status": "error", "message": "Failed to process the latest data."}), 500
    except Exception as e:
        print(f"Error during pipeline execution: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)

import os
import certifi
os.environ['REQUESTS_CA_BUNDLE'] = certifi.where()

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from transformers import pipeline
import time
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime
import smtplib
from email.message import EmailMessage
from flask import Flask, request, jsonify

# --- Setup constants ---
SERVICE_ACCOUNT_FILE = r"C:\Users\wfhq_lpham\Downloads\comment-analyzer-463511-51737bb4e537.json"
SHEET_NAME = "Automated Supervisor Report"
MODEL_PATH = r"C:\Users\wfhq_lpham\OneDrive - Mestek, Inc\jsonfiles"

SENDER_EMAIL = "lunachpham@gmail.com"
SENDER_APP_PASSWORD = "dcrnytbtcvjzntju"

app = Flask(__name__)

def run_pipeline():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
    gc = gspread.authorize(creds)
    sheet = gc.open(SHEET_NAME).sheet1

    headers = [h.strip().replace('\n', ' ') for h in sheet.row_values(1)]

    def find_header_index(headers, keyword):
        for i, h in enumerate(headers):
            if keyword.lower() in h.lower():
                return i
        raise ValueError(f"Header matching '{keyword}' not found")

    comment_keywords = [
        "How does this employee typically respond to feedback",
        "How effectively does this employee communicate with others",
        "How reliable is this employee in terms of attendance and use of time",
        "When your team encounters workflow disruptions",
        "In what ways does this employee demonstrate commitment to safety",
        "How effectively does this employee use technical documentation"
    ]

    score_headers = [
        "Score - Feedback & Conflict Resolution",
        "Score - Communication & Team Support",
        "Score - Reliability & Productivity",
        "Score - Adaptability & Quality Focus",
        "Score - Safety Commitment",
        "Score - Documentation & Procedures"
    ]

    # Add missing columns if needed
    for sh in score_headers + ["AI Score", "AI Feedback"]:
        if sh not in headers:
            sheet.add_cols(1)
            sheet.update_cell(1, len(headers) + 1, sh)
            headers.append(sh)

    comment_cols_idx = [find_header_index(headers, kw) + 1 for kw in comment_keywords]
    score_cols_idx = [find_header_index(headers, sh) + 1 for sh in score_headers]
    score_col = find_header_index(headers, "AI Score") + 1
    feedback_col = find_header_index(headers, "AI Feedback") + 1

    print("Loading zero-shot classifier...")
    classifier = pipeline("zero-shot-classification", model=MODEL_PATH, tokenizer=MODEL_PATH, local_files_only=True)

    label_to_score = {
        "excellent": 5,
        "good": 4,
        "average": 3,
        "poor": 2,
        "unacceptable": 1
    }
    labels = list(label_to_score.keys())

    def generate_feedback(score_summaries, avg_score):
        strengths, neutrals, weaknesses = [], [], []
        category_map = {
            "Feedback & Conflict Resolution": "responding to feedback and resolving conflict",
            "Communication & Team Support": "communication and team collaboration",
            "Reliability & Productivity": "reliability and productivity",
            "Adaptability & Quality Focus": "adaptability and quality assurance",
            "Safety Commitment": "safety and workplace organization",
            "Documentation & Procedures": "technical documentation and procedures"
        }
        for s in score_summaries:
            cat, val = s.split(":")
            score = int(val.split("(")[1][0])
            phrase = category_map.get(cat.strip(), cat.strip())
            if score >= 4:
                strengths.append(phrase)
            elif score == 3:
                neutrals.append(phrase)
            else:
                weaknesses.append(phrase)
        parts = []
        if strengths: parts.append("Strengths include " + ", ".join(strengths) + ".")
        if neutrals: parts.append("Satisfactory areas: " + ", ".join(neutrals) + ".")
        if weaknesses: parts.append("Needs improvement in " + ", ".join(weaknesses) + ".")
        parts.append(f"Overall performance score: {avg_score}/5.")
        return " ".join(parts)

    data = sheet.get_all_records()
    print(f"Total rows fetched: {len(data)}")

    latest_row_idx = None
    latest_time = None

    for i, row in enumerate(data):
        ts_str = row.get("Timestamp", "")
        for fmt in ["%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %I:%M:%S %p"]:
            try:
                ts = datetime.strptime(ts_str, fmt)
                break
            except ValueError:
                ts = None
        if not ts:
            continue
        if latest_time is None or ts > latest_time:
            latest_time = ts
            latest_row_idx = i + 2

    if not latest_row_idx:
        print("❌ No valid timestamps found in the 'Timestamp' column.")
        return False

    row = sheet.row_values(latest_row_idx)
    row_dict = dict(zip(headers, row))

    employee_name = row_dict.get("Employee Name", f"Employee {latest_row_idx}")
    department = row_dict.get("Department", "N/A")
    supervisor_name = row_dict.get("Supervisor Name", "N/A")
    date_of_review = time.strftime("%Y-%m-%d")
    submitter_email = row_dict.get("Email", "").strip()

    if not submitter_email:
        print("❌ No submitter email found in the latest row.")
        return False

    comment_scores = []
    score_summaries = []

    for j, col_idx in enumerate(comment_cols_idx):
        val = sheet.cell(latest_row_idx, col_idx).value
        if val:
            result = classifier(val, labels)
            top_label = result['labels'][0]
            score = label_to_score[top_label]
            comment_scores.append(score)
            sheet.update_cell(latest_row_idx, score_cols_idx[j], score)
            section = score_headers[j].replace("Score - ", "")
            score_summaries.append(f"{section}: {top_label} ({score}/5)")
        else:
            comment_scores.append(None)

    valid_scores = [s for s in comment_scores if s is not None]
    avg_score = round(sum(valid_scores) / len(valid_scores), 2) if valid_scores else 0
    feedback = generate_feedback(score_summaries, avg_score)
    sheet.update_cell(latest_row_idx, score_col, avg_score)
    sheet.update_cell(latest_row_idx, feedback_col, feedback)

    document = Document()
    for section in document.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    def add_line(text):
        p = document.add_paragraph(text)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        return p

    document.add_heading('MESTEK – Hourly Performance Appraisal', level=1)
    document.add_heading('Employee Information', level=2)
    add_line(f"\u2022 Employee Name: {employee_name}")
    add_line(f"\u2022 Department: {department}")
    add_line(f"\u2022 Supervisor Name: {supervisor_name}")
    add_line(f"\u2022 Date of Review: {date_of_review}")

    document.add_heading('Core Performance Categories', level=2).paragraph_format.space_after = Pt(0)
    scale = document.add_paragraph("1 – Poor | 2 – Needs Improvement | 3 – Meets Expectations | 4 – Exceeds Expectations | 5 – Outstanding")
    scale.paragraph_format.space_before = Pt(0)
    scale.paragraph_format.space_after = Pt(3)
    scale.paragraph_format.line_spacing = 1.0
    scale.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    scale.runs[0].font.size = Pt(9)

    table = document.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Category'
    hdr_cells[1].text = 'Rating (1–5)'
    hdr_cells[2].text = 'Supervisor Comments'
    for cell in hdr_cells:
        for p in cell.paragraphs:
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p.runs[0].bold = True

    for cat, idx, cidx in zip(score_headers, score_cols_idx, comment_cols_idx):
        row_cells = table.add_row().cells
        row_cells[0].text = cat.replace("Score - ", "")
        row_cells[1].text = str(sheet.cell(latest_row_idx, idx).value or "")
        row_cells[2].text = str(sheet.cell(latest_row_idx, cidx).value or "")
        for cell in row_cells:
            for p in cell.paragraphs:
                p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    document.add_heading('Performance Summary', level=2)
    add_line(feedback)

    document.add_heading('Goals for Next Review Period', level=2)
    for n in range(1, 4):
        p = document.add_paragraph()
        run = p.add_run(f"{n}. " + "_" * 60)
        run.font.size = Pt(12)

    document.add_heading('Sign-Offs', level=2)
    sig1 = document.add_paragraph()
    sig1.add_run('Employee Signature: ' + "_" * 50).bold = True
    sig1.add_run('\t\tDate: ' + "_" * 15)
    sig2 = document.add_paragraph()
    sig2.add_run('Supervisor Signature: ' + "_" * 50).bold = True
    sig2.add_run('\t\tDate: ' + "_" * 15)

    safe_name = employee_name.replace(" ", "_").replace("/", "_")
    output_path = os.path.abspath(f"{safe_name}_performance_report.docx")
    document.save(output_path)
    print(f"\n✅ Report saved to: {output_path}")

    def send_email_with_attachment(to_email, subject, body, attachment_path):
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg.set_content(body)

        with open(attachment_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(attachment_path)

        msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.wordprocessingml.document', filename=file_name)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(SENDER_EMAIL, SENDER_APP_PASSWORD)
            smtp.send_message(msg)

        print(f"📧 Email sent to {to_email} with attachment {file_name}")

    email_subject = "Your Performance Report"
    email_body = f"Attached is the latest performance report for {employee_name}.\n\nBest regards,\nMestek"
    send_email_with_attachment(submitter_email, email_subject, email_body, output_path)

    return True

