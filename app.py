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
# Local fallback path for your JSON (only for local dev)
LOCAL_SERVICE_ACCOUNT_FILE = r"C:\Users\wfhq_lpham\Downloads\comment-analyzer-463511-51737bb4e537.json"
SHEET_NAME = "Automated Supervisor Report"
MODEL_PATH = r"C:\Users\wfhq_lpham\OneDrive - Mestek, Inc\jsonfiles"

SENDER_EMAIL = "lunachpham@gmail.com"
SENDER_APP_PASSWORD = "dcrnytbtcvjzntju"

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

def run_pipeline():
    creds = get_gspread_creds()
    gc = gspread.authorize(creds)
    sheet = gc.open(SHEET_NAME).sheet1

    # Your full run_pipeline function code here...
    # (keep everything exactly as you wrote it)
    pass  # Replace this pass with your full run_pipeline body

@app.route('/', methods=['GET'])
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
