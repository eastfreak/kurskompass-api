"""
KursKompass API – Flask Backend (für Render.com)
Stellt die Scraping-Funktionalität als REST-API bereit.
"""
import os
import json
import threading
from datetime import datetime
from io import BytesIO
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from scraper import (
    QISScraper, tree_to_dict, dict_to_tree
)
app = Flask(__name__)
CORS(app)
# Einfache API-Key Authentifizierung
API_KEY = os.environ.get("API_KEY", "loni-kurskompass-2026-secret")
# Globaler Scraper
scraper_instance = QISScraper()
cached_tree = None
cached_veranstaltungen = None
scraper_lock = threading.Lock()
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
os.makedirs(DATA_DIR, exist_ok=True)
def check_auth():
    """Prüft API-Key im Header oder Query-Parameter."""
    key = request.headers.get("X-API-Key") or request.args.get("key")
    if key != API_KEY:
        return jsonify({"error": "Unauthorized"}), 401
    return None
@app.route("/")
def health():
    return jsonify({"status": "ok", "app": "KursKompass API", "version": "1.0"})
@app.route("/api/ping")
def ping():
    """Wakeup-Endpoint (kein Auth nötig)."""
    return jsonify({"status": "awake"})
@app.route("/api/scan-tree", methods=["POST"])
def api_scan_tree():
    auth_error = check_auth()
    if auth_error:
        return auth_error
    global cached_tree
    if scraper_lock.locked():
        return jsonify({"error": "Scraping läuft bereits"}), 409
    def do_scan():
        global cached_tree
        with scraper_lock:
            tree = scraper_instance.scan_tree()
            cached_tree = tree
            tree_data = tree_to_dict(tree)
            with open(os.path.join(DATA_DIR, "tree.json"), "w", encoding="utf-8") as f:
                json.dump(tree_data, f, ensure_ascii=False, indent=2)
    thread = threading.Thread(target=do_scan)
    thread.start()
    return jsonify({"status": "started"})
@app.route("/api/tree")
def api_get_tree():
    """Gibt gecachten Baum zurück."""
    auth_error = check_auth()
    if auth_error:
        return auth_error
    global cached_tree
    tree_file = os.path.join(DATA_DIR, "tree.json")
    if cached_tree:
        return jsonify(tree_to_dict(cached_tree))
    elif os.path.exists(tree_file):
        with open(tree_file, "r", encoding="utf-8") as f:
            data = json.load(f)
            cached_tree = dict_to_tree(data)
            return jsonify(data)
    return jsonify(None)
@app.route("/api/scrape", methods=["POST"])
def api_scrape():
    auth_error = check_auth()
    if auth_error:
        return auth_error
    global cached_tree, cached_veranstaltungen
    if scraper_lock.locked():
        return jsonify({"error": "Scraping läuft bereits"}), 409
    selected = request.json.get("selected", [])
    if not selected:
        return jsonify({"error": "Keine Bereiche ausgewählt"}), 400
    if not cached_tree:
        tree_file = os.path.join(DATA_DIR, "tree.json")
        if os.path.exists(tree_file):
            with open(tree_file, "r", encoding="utf-8") as f:
                cached_tree = dict_to_tree(json.load(f))
        else:
            return jsonify({"error": "Bitte zuerst Struktur laden"}), 400
    def do_scrape():
        global cached_veranstaltungen
        with scraper_lock:
            veranstaltungen = scraper_instance.scrape_selected(cached_tree, set(selected))
            from dataclasses import asdict
            ver_data = [asdict(v) for v in veranstaltungen]
            cached_veranstaltungen = ver_data
            with open(os.path.join(DATA_DIR, "veranstaltungen.json"), "w", encoding="utf-8") as f:
                json.dump(ver_data, f, ensure_ascii=False, indent=2)
    thread = threading.Thread(target=do_scrape)
    thread.start()
    return jsonify({"status": "started"})
@app.route("/api/progress")
def api_progress():
    auth_error = check_auth()
    if auth_error:
        return auth_error
    return jsonify(scraper_instance.progress)
@app.route("/api/veranstaltungen")
def api_veranstaltungen():
    """Gibt gecachte Veranstaltungen zurück."""
    auth_error = check_auth()
    if auth_error:
        return auth_error
    global cached_veranstaltungen
    ver_file = os.path.join(DATA_DIR, "veranstaltungen.json")
    if cached_veranstaltungen:
        return jsonify({"data": cached_veranstaltungen, "count": len(cached_veranstaltungen)})
    elif os.path.exists(ver_file):
        with open(ver_file, "r", encoding="utf-8") as f:
            cached_veranstaltungen = json.load(f)
            return jsonify({"data": cached_veranstaltungen, "count": len(cached_veranstaltungen)})
    return jsonify({"data": None, "count": 0})
@app.route("/api/download-excel")
def api_download_excel():
    auth_error = check_auth()
    if auth_error:
        return auth_error
    global cached_veranstaltungen
    if not cached_veranstaltungen:
        ver_file = os.path.join(DATA_DIR, "veranstaltungen.json")
        if os.path.exists(ver_file):
            with open(ver_file, "r", encoding="utf-8") as f:
                cached_veranstaltungen = json.load(f)
    if not cached_veranstaltungen:
        return jsonify({"error": "Keine Daten vorhanden"}), 404
    wb = Workbook()
    ws = wb.active
    ws.title = "Stundenplan"
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="D4467E", end_color="D4467E", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    headers = [
        "Bereich", "Kennung", "Titel", "Art", "Dozent",
        "Tag", "Zeit", "Rhythmus", "Raum", "SWS",
        "Max. TN", "Belegung", "Semester", "Studiengänge"
    ]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    for row_idx, v in enumerate(cached_veranstaltungen, 2):
        values = [
            v.get("pfad", ""), v.get("kennung", ""), v.get("titel", ""),
            v.get("veranstaltungsart", ""), v.get("dozent", ""),
            v.get("tag", ""), v.get("zeit", ""), v.get("rhythmus", ""),
            v.get("raum", ""), v.get("sws", ""), v.get("max_teilnehmer", ""),
            v.get("belegung", ""), v.get("semester", ""), v.get("studiengaenge", ""),
        ]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    widths = [30, 12, 40, 20, 25, 6, 16, 8, 30, 5, 8, 12, 12, 40]
    for i, w in enumerate(widths):
        col_letter = chr(65 + i)
        ws.column_dimensions[col_letter].width = w
    ws.auto_filter.ref = f"A1:N{len(cached_veranstaltungen) + 1}"
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"KursKompass_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
