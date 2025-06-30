import os
import json
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
LOCAL_SERVICE_ACCOUNT_FILE = r"C:\Users\wfhq_lpham\Downloads\comment-analyzer-463511-51737bb4e537.json"
SHEET_NAME = "Automated Supervisor Report"
MODEL_PATH = r"C:\Users\wfhq_lpham\OneDrive - Mestek, Inc\jsonfiles"

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "lunachpham@gmail.com")
SENDER_APP_PASSWORD = os.environ.get("SENDER_APP_PASSWORD", "dcrnytbtcvjzntju")

app = Flask(__name__)

def get_gspread_creds():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds_json_str = os.environ.get("GOOGLE_CREDS_JSON")
    if creds_json_str:
        creds_dict = json.loads(creds_json_str)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name(LOCAL_SERVICE_ACCOUNT_FILE, scope)
    return creds

def run_pipeline(payload):
    try:
        creds = get_gspread_creds()
        gc = gspread.authorize(creds)
        sheet = gc.open(SHEET_NAME).sheet1

        headers = [h.strip().replace('\n', ' ') for h in sheet.row_values(1)]

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

        row = [
            payload.get("timestamp", ""),
            payload.get("email", ""),
            payload.get("employeeName", ""),
            payload.get("department", ""),
            payload.get("supervisorName", "")
        ] + payload.get("comments", [])

        sheet.append_row(row)

        return True
    except Exception as e:
        print("❌ Error in run_pipeline:", e)
        return False

@app.route('/', methods=['GET'])
def home():
    return '✅ Supervisor AI Webhook is live!'

@app.route('/submit', methods=['POST'])
def submit():
    data = request.json
    print("📨 Received POST /submit with payload:", json.dumps(data, indent=2))
    try:
        success = run_pipeline(data)
        if success:
            return jsonify({"status": "success", "message": "Processed successfully"})
        else:
            return jsonify({"status": "error", "message": "Processing failed"}), 500
    except Exception as e:
        print("❌ Exception during /submit:", e)
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)


@app.route('/submit', methods=['POST'])
def submit():
    data = request.get_json()
    print("📥 Received submission:", json.dumps(data, indent=2))

    try:
        # Process your data here or trigger your pipeline
        run_pipeline()  # Optionally, pass data to this function if needed
        return jsonify({"status": "success", "message": "Data received and processed!"}), 200
    except Exception as e:
        print("❌ Error processing submission:", str(e))
        return jsonify({"status": "error", "message": str(e)}), 500

