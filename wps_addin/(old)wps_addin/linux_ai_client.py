# linux_ai_client.py
import datetime
import traceback
import sys
import threading
import requests
from flask import Flask, request, jsonify

# --- Configuration ---
# This is the address of your main AI backend server
BACKEND_URL = "http://127.0.0.1:8000"

# This is the port the local client server will listen on.
# The WPS macro will send its requests here.
CLIENT_PORT = 5001

# --- App Initialization ---
app = Flask(__name__)

# --- Terminal Logging for Debugging ---
def log_message(message):
    """Writes a message with a timestamp to the terminal (stdout)."""
    print(f"[{datetime.datetime.now()}] {message}")

# --- Forwarding Requests to the Backend Server ---
def forward_request_to_backend(endpoint: str, data: dict):
    """Forwards a request from the WPS macro to the main backend server."""
    log_message(f"Forwarding request to backend endpoint: {endpoint}")
    try:
        response = requests.post(f"{BACKEND_URL}{endpoint}", json=data, timeout=300)
        response.raise_for_status()
        result = response.json().get("result", "")
        log_message("Successfully received response from AI backend.")
        return jsonify({"result": result})
    except requests.exceptions.ConnectionError as e:
        log_message(f"ConnectionError calling {endpoint}: {e}")
        return jsonify({"result": "ERROR: Could not connect to the AI backend server."})
    except Exception as e:
        log_message(f"Unexpected Exception calling {endpoint}: {e}\n{traceback.format_exc()}")
        return jsonify({"result": f"An unexpected error occurred: {e}"})

# --- Flask Routes (Endpoints for WPS Macro) ---
@app.route('/run_prompt', methods=['POST'])
def run_prompt():
    data = request.json
    prompt = data.get("prompt", "")
    return forward_request_to_backend("/process", {"prompt": prompt})

@app.route('/analyze_document', methods=['POST'])
def analyze_document():
    data = request.json
    content = data.get("content", "")
    prompt = data.get("prompt", "")
    return forward_request_to_backend("/analyze", {"content": content, "prompt": prompt})

@app.route('/summarize_document', methods=['POST'])
def summarize_document():
    data = request.json
    content = data.get("content", "")
    return forward_request_to_backend("/summarize", {"content": content})

# Add routes for other endpoints like memo, report, etc.
@app.route('/create_report', methods=['POST'])
def create_report():
    data = request.json
    prompt = data.get("prompt", "")
    return forward_request_to_backend("/create_report", {"prompt": prompt})

@app.route('/create_memo', methods=['POST'])
def create_memo():
    data = request.json
    return forward_request_to_backend("/create_memo", data)

@app.route('/create_cover_letter', methods=['POST'])
def create_cover_letter():
    data = request.json
    return forward_request_to_backend("/create_cover_letter", data)

@app.route('/create_minutes', methods=['POST'])
def create_minutes():
    data = request.json
    return forward_request_to_backend("/create_minutes", data)

# --- Server Start-up ---
if __name__ == '__main__':
    log_message(f"Starting WPS AI client server on port {CLIENT_PORT}...")
    try:
        app.run(host='127.0.0.1', port=CLIENT_PORT)
    except Exception as e:
        log_message(f"Failed to start server: {e}")
        sys.exit(1)