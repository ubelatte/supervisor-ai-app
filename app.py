from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route('/')
def home():
    return '✅ Supervisor AI Webhook is live!'

@app.route('/run-script', methods=['POST'])
def run_script():
    data = request.json
    print("Received data:", data)
    return jsonify({"status": "success", "message": "Webhook triggered successfully!"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
