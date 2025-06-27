from flask import Flask, request, jsonify
from run_pipeline import run_pipeline

app = Flask(__name__)

@app.route('/')
def home():
    return '✅ Supervisor AI Webhook is live!'

@app.route('/run-script', methods=['POST'])
def run_script():
    data = request.json
    print("Received data:", data)
    try:
        run_pipeline()
    except Exception as e:
        print(f"Error running pipeline: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

    return jsonify({"status": "success", "message": "Webhook triggered and script executed successfully!"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
