"""
Minimal Flask server for the Georgia Economic Data site.

- Serves static files (index.html, data/)
- Provides /admin page for uploading country×product Excel files
- Triggers data rebuild after upload

Run: python server.py [--port 5000]
"""

import subprocess
import sys
from pathlib import Path

from flask import Flask, jsonify, redirect, request, send_from_directory

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "geostat_data"
UPLOAD_NAME = "export_by_country_product.xlsx"

app = Flask(__name__, static_folder=str(BASE_DIR), static_url_path="")


@app.route("/")
def index():
    return send_from_directory(str(BASE_DIR), "index.html")


@app.route("/admin")
def admin():
    return send_from_directory(str(BASE_DIR), "admin.html")


@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("file")
    if not f or not f.filename.endswith(".xlsx"):
        return jsonify(ok=False, error="Please upload an .xlsx file"), 400

    DATA_DIR.mkdir(exist_ok=True)
    dest = DATA_DIR / UPLOAD_NAME
    f.save(str(dest))

    # Rebuild JSON
    try:
        result = subprocess.run(
            [sys.executable, "build_data.py"],
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            timeout=120,
        )
        if result.returncode != 0:
            return jsonify(ok=False, error=result.stderr[-500:] if result.stderr else "Build failed"), 500
        return jsonify(ok=True, output=result.stdout[-500:])
    except subprocess.TimeoutExpired:
        return jsonify(ok=False, error="Build timed out"), 500


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=5000)
    parser.add_argument("--host", default="0.0.0.0")
    args = parser.parse_args()
    app.run(host=args.host, port=args.port, debug=True)
