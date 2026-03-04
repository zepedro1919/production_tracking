"""
============================================================
PRINT AGENT — Automatic Label Printer Service
============================================================

A lightweight HTTP server that runs on the PC connected to
the label printer. Receives print requests from Google Apps
Script and prints PDFs automatically.

Requirements:
    pip install flask requests

Usage:
    python print_agent.py

The agent listens on port 5555 by default.
Google Apps Script sends a POST with the PDF content,
the agent saves it temporarily and sends it to the printer.

On Windows, it uses SumatraPDF (free) for silent printing.
On Linux/Mac, it uses lp/lpr commands.
"""

import os
import sys
import time
import json
import tempfile
import subprocess
import platform
import logging
from pathlib import Path
from datetime import datetime

from flask import Flask, request, jsonify

# ============================================================
# CONFIGURATION — Edit these to match your setup
# ============================================================

# Port the agent listens on
PORT = 5555

# Printer name (as it appears in your OS printer list)
# Set to None to use the default printer
PRINTER_NAME = "EPSON ET-M1170 Series"

# Path to SumatraPDF (Windows only — for silent printing)
# Download free from: https://www.sumatrapdfreader.org/
SUMATRA_PATH = r"C:\Users\ASUS-JPB\AppData\Local\SumatraPDF\SumatraPDF.exe"

# Alternatively, if you prefer Foxit Reader:
# FOXIT_PATH = r"C:\Program Files (x86)\Foxit Software\Foxit Reader\FoxitReader.exe"

# Temp folder for downloaded PDFs
TEMP_DIR = os.path.join(tempfile.gettempdir(), "print_agent_labels")

# Secret token to prevent unauthorized printing (set the same in Apps Script)
AUTH_TOKEN = "producao2026"

# ============================================================

app = Flask(__name__)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("print_agent.log", encoding="utf-8")
    ]
)
log = logging.getLogger("PrintAgent")

# Ensure temp dir exists
os.makedirs(TEMP_DIR, exist_ok=True)


@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint — Apps Script can ping this."""
    return jsonify({
        "status": "ok",
        "printer": PRINTER_NAME or "(default)",
        "platform": platform.system(),
        "time": datetime.now().isoformat()
    })


@app.route("/print", methods=["POST"])
def print_label():
    """
    Receives a PDF as binary data in the request body and prints it.
    
    Headers:
        Authorization: Bearer <AUTH_TOKEN>
        Content-Type: application/pdf
    
    Query params (optional):
        copies: number of copies (default 1)
    """
    # Auth check
    auth = request.headers.get("Authorization", "")
    expected = f"Bearer {AUTH_TOKEN}"
    if auth != expected:
        log.warning("Unauthorized print attempt")
        return jsonify({"error": "Unauthorized"}), 401

    # Get PDF data
    pdf_data = request.get_data()
    if not pdf_data or len(pdf_data) < 100:
        return jsonify({"error": "No PDF data received"}), 400

    copies = int(request.args.get("copies", 1))

    # Save to temp file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    pdf_path = os.path.join(TEMP_DIR, f"label_{timestamp}.pdf")

    try:
        with open(pdf_path, "wb") as f:
            f.write(pdf_data)
        log.info(f"PDF saved: {pdf_path} ({len(pdf_data)} bytes)")
        
        # Save a debug copy for inspection
        debug_path = os.path.join(TEMP_DIR, "last_label_debug.pdf")
        with open(debug_path, "wb") as f:
            f.write(pdf_data)
        log.info(f"Debug copy saved: {debug_path}")
    except Exception as e:
        log.error(f"Failed to save PDF: {e}")
        return jsonify({"error": f"Failed to save PDF: {e}"}), 500

    # Print it
    try:
        success = send_to_printer(pdf_path, copies)
        if success:
            log.info(f"✅ Printed successfully: {pdf_path}")
            # Clean up temp file after a short delay
            cleanup_later(pdf_path)
            return jsonify({"status": "printed", "file": pdf_path})
        else:
            log.error(f"❌ Print failed: {pdf_path}")
            return jsonify({"error": "Print command failed"}), 500
    except Exception as e:
        log.error(f"❌ Print error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/print-url", methods=["POST"])
def print_from_url():
    """
    Alternative: receives a JSON body with a URL to a PDF,
    downloads it, and prints it.
    
    Body: { "url": "https://...", "copies": 1 }
    """
    import requests as req

    auth = request.headers.get("Authorization", "")
    if auth != f"Bearer {AUTH_TOKEN}":
        return jsonify({"error": "Unauthorized"}), 401

    data = request.get_json()
    if not data or "url" not in data:
        return jsonify({"error": "Missing 'url' in body"}), 400

    pdf_url = data["url"]
    copies = int(data.get("copies", 1))

    try:
        resp = req.get(pdf_url, timeout=30)
        resp.raise_for_status()
        pdf_data = resp.content
    except Exception as e:
        log.error(f"Failed to download PDF: {e}")
        return jsonify({"error": f"Download failed: {e}"}), 500

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    pdf_path = os.path.join(TEMP_DIR, f"label_{timestamp}.pdf")

    with open(pdf_path, "wb") as f:
        f.write(pdf_data)

    try:
        success = send_to_printer(pdf_path, copies)
        cleanup_later(pdf_path)
        if success:
            return jsonify({"status": "printed"})
        else:
            return jsonify({"error": "Print failed"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def send_to_printer(pdf_path, copies=1):
    """
    Send a PDF file to the printer.
    Uses SumatraPDF on Windows, lp/lpr on Linux/Mac.
    """
    system = platform.system()

    if system == "Windows":
        return print_windows(pdf_path, copies)
    elif system == "Linux":
        return print_linux(pdf_path, copies)
    elif system == "Darwin":  # macOS
        return print_mac(pdf_path, copies)
    else:
        log.error(f"Unsupported platform: {system}")
        return False


def print_windows(pdf_path, copies=1):
    """
    Silent print on Windows using SumatraPDF.
    SumatraPDF supports: -print-to <printer> -print-to-default
    """
    if os.path.exists(SUMATRA_PATH):
        # SumatraPDF silent print (best option)
        for _ in range(copies):
            cmd = [SUMATRA_PATH, "-print-to-default", "-silent", pdf_path]
            if PRINTER_NAME:
                cmd = [SUMATRA_PATH, "-print-to", PRINTER_NAME, "-silent", pdf_path]
            
            log.info(f"Running: {' '.join(cmd)}")
            result = subprocess.run(cmd, capture_output=True, timeout=30)
            if result.returncode != 0:
                log.error(f"SumatraPDF error: {result.stderr.decode('utf-8', errors='replace')}")
                return False
        return True
    else:
        # Fallback: use Windows ShellExecute "print" verb
        # This may show a dialog depending on the default PDF reader
        log.warning(f"SumatraPDF not found at {SUMATRA_PATH}. Using ShellExecute fallback.")
        try:
            for _ in range(copies):
                os.startfile(pdf_path, "print")
                time.sleep(2)  # Give it time to spool
            return True
        except Exception as e:
            log.error(f"ShellExecute print failed: {e}")
            return False


def print_linux(pdf_path, copies=1):
    """Silent print on Linux using lp command."""
    cmd = ["lp", "-n", str(copies)]
    if PRINTER_NAME:
        cmd.extend(["-d", PRINTER_NAME])
    cmd.append(pdf_path)
    
    log.info(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, timeout=30)
    return result.returncode == 0


def print_mac(pdf_path, copies=1):
    """Silent print on macOS using lpr command."""
    cmd = ["lpr", "-#", str(copies)]
    if PRINTER_NAME:
        cmd.extend(["-P", PRINTER_NAME])
    cmd.append(pdf_path)
    
    log.info(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, timeout=30)
    return result.returncode == 0


def cleanup_later(pdf_path, delay_seconds=60):
    """Delete a file after a delay (in a background thread)."""
    import threading
    def _delete():
        time.sleep(delay_seconds)
        try:
            os.remove(pdf_path)
            log.info(f"Cleaned up: {pdf_path}")
        except:
            pass
    t = threading.Thread(target=_delete, daemon=True)
    t.start()


def list_printers():
    """List available printers on the system."""
    system = platform.system()
    
    if system == "Windows":
        try:
            import subprocess
            result = subprocess.run(
                ["wmic", "printer", "get", "name"],
                capture_output=True, text=True, timeout=10
            )
            printers = [line.strip() for line in result.stdout.split("\n") if line.strip() and line.strip() != "Name"]
            return printers
        except:
            return ["(could not list printers)"]
    elif system == "Linux":
        result = subprocess.run(["lpstat", "-p"], capture_output=True, text=True, timeout=10)
        return result.stdout.strip().split("\n")
    else:
        return ["(use system preferences to find printer names)"]


if __name__ == "__main__":
    print("=" * 60)
    print("  🖨️  PRINT AGENT — Production Label Printer")
    print("=" * 60)
    print(f"  Platform:  {platform.system()}")
    print(f"  Port:      {PORT}")
    print(f"  Printer:   {PRINTER_NAME or '(default printer)'}")
    
    if platform.system() == "Windows":
        if os.path.exists(SUMATRA_PATH):
            print(f"  SumatraPDF: ✅ Found")
        else:
            print(f"  SumatraPDF: ❌ NOT FOUND at {SUMATRA_PATH}")
            print(f"               Download from: https://www.sumatrapdfreader.org/")
            print(f"               (will fallback to ShellExecute)")
    
    print()
    print("  Available printers:")
    for p in list_printers():
        print(f"    - {p}")
    
    print()
    print(f"  Endpoints:")
    print(f"    GET  http://localhost:{PORT}/health")
    print(f"    POST http://localhost:{PORT}/print")
    print(f"    POST http://localhost:{PORT}/print-url")
    print()
    print("  Waiting for print jobs...")
    print("=" * 60)
    
    app.run(host="0.0.0.0", port=PORT, debug=False)
