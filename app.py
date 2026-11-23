from flask import Flask, jsonify, request, render_template, send_file
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from pathlib import Path
from datetime import datetime, date
from threading import RLock
import os, re, shutil, sys, threading, webbrowser

# PyInstaller bundles templates/static as read-only resources, while workbook data
# must live in a writable folder outside the temporary extraction directory.
BASE_DIR = Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
if getattr(sys, "frozen", False):
    DATA_DIR = Path.home() / "Documents" / "SheetFlow"
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    bundled_sample = RESOURCE_DIR / "sample_data.xlsx"
    user_sample = DATA_DIR / "sample_data.xlsx"
    if not user_sample.exists() and bundled_sample.exists():
        shutil.copy2(bundled_sample, user_sample)
else:
    DATA_DIR = BASE_DIR

app = Flask(__name__, template_folder=str(RESOURCE_DIR / "templates"), static_folder=str(RESOURCE_DIR / "static"))
EXCEL_FILE = Path(os.environ.get("EXCEL_FILE", DATA_DIR / "sample_data.xlsx"))
SHEET_NAME = os.environ.get("SHEET_NAME", "Employee Records")
LOCK = RLock()

def clean_value(value):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    return value

def workbook_data():
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE}")
    wb = load_workbook(EXCEL_FILE)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    headers = [str(c.value or f"Column {i}") for i, c in enumerate(ws[1], 1)]
    records = []
    for excel_row in range(2, ws.max_row + 1):
        vals = [clean_value(ws.cell(excel_row, col).value) for col in range(1, len(headers)+1)]
        if any(v != "" for v in vals):
            records.append({"_row": excel_row, "values": vals})
    return headers, records

def style_workbook(ws):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 30
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(name="Aptos Display", bold=True, color="FFFFFF", size=11)
        cell.alignment = Alignment(vertical="center", horizontal="left")
        cell.border = Border(bottom=Side(style="medium", color="2F75B5"))
    for row in range(2, ws.max_row + 1):
        ws.row_dimensions[row].height = 23
        for cell in ws[row]:
            cell.font = Font(name="Aptos", size=10, color="243247")
            cell.alignment = Alignment(vertical="center")
            if row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F2F6FC")
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        max_len = max([len(str(ws.cell(r, col).value or "")) for r in range(1, ws.max_row+1)] + [12])
        ws.column_dimensions[letter].width = min(max(max_len + 3, 14), 36)
    # Rebuild the Excel table so added/deleted columns and rows remain included.
    for name in list(ws.tables.keys()):
        del ws.tables[name]
    if ws.max_row >= 1 and ws.max_column >= 1:
        ref = f"A1:{get_column_letter(ws.max_column)}{max(ws.max_row, 1)}"
        table = Table(displayName="ManagedRecords", ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        ws.add_table(table)
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(ws.max_row, 1)}"

def save_wb(wb):
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    style_workbook(ws)
    temp = EXCEL_FILE.with_suffix(".tmp.xlsx")
    wb.save(temp)
    temp.replace(EXCEL_FILE)

@app.get("/")
def index():
    return render_template("index.html")

@app.get("/api/data")
def get_data():
    with LOCK:
        headers, records = workbook_data()
    return jsonify({"headers": headers, "records": records, "file": EXCEL_FILE.name})

@app.post("/api/rows")
def add_row():
    data = request.get_json(force=True)
    with LOCK:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        vals = data.get("values", [])
        vals = (vals + [""] * ws.max_column)[:ws.max_column]
        ws.append(vals)
        save_wb(wb)
        row_id = ws.max_row
    return jsonify({"ok": True, "row": row_id})

@app.put("/api/cells")
def update_cell():
    data = request.get_json(force=True)
    row, col, value = int(data["row"]), int(data["col"]), data.get("value", "")
    with LOCK:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        if row < 2 or row > ws.max_row or col < 1 or col > ws.max_column:
            return jsonify({"error": "Cell is outside the data range."}), 400
        # Keep numeric values numeric when the input looks like a number.
        if isinstance(value, str):
            v = value.strip()
            if v and re.fullmatch(r"-?\d+(\.\d+)?", v):
                value = float(v) if "." in v else int(v)
        ws.cell(row, col).value = value
        save_wb(wb)
    return jsonify({"ok": True})

@app.delete("/api/rows/<int:row>")
def delete_row(row):
    with LOCK:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        if row < 2 or row > ws.max_row:
            return jsonify({"error": "Record not found."}), 404
        ws.delete_rows(row, 1)
        save_wb(wb)
    return jsonify({"ok": True})

@app.post("/api/columns")
def add_column():
    data = request.get_json(force=True)
    name = str(data.get("name", "")).strip()
    if not name:
        return jsonify({"error": "Column name is required."}), 400
    with LOCK:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        existing = [str(c.value).strip().lower() for c in ws[1] if c.value]
        if name.lower() in existing:
            return jsonify({"error": "A column with that name already exists."}), 400
        new_col = ws.max_column + 1
        ws.cell(1, new_col).value = name
        for r in range(2, ws.max_row + 1):
            ws.cell(r, new_col).value = ""
        save_wb(wb)
    return jsonify({"ok": True})

@app.put("/api/columns/<int:col>")
def rename_column(col):
    name = str(request.get_json(force=True).get("name", "")).strip()
    if not name:
        return jsonify({"error": "Column name is required."}), 400
    with LOCK:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        if col < 1 or col > ws.max_column:
            return jsonify({"error": "Column not found."}), 404
        existing = [str(ws.cell(1, c).value or "").strip().lower() for c in range(1, ws.max_column+1) if c != col]
        if name.lower() in existing:
            return jsonify({"error": "A column with that name already exists."}), 400
        ws.cell(1, col).value = name
        save_wb(wb)
    return jsonify({"ok": True})

@app.delete("/api/columns/<int:col>")
def delete_column(col):
    with LOCK:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        if col < 1 or col > ws.max_column:
            return jsonify({"error": "Column not found."}), 404
        if ws.max_column <= 1:
            return jsonify({"error": "The worksheet must keep at least one column."}), 400
        ws.delete_cols(col, 1)
        save_wb(wb)
    return jsonify({"ok": True})

@app.get("/download")
def download():
    return send_file(EXCEL_FILE, as_attachment=True, download_name=EXCEL_FILE.name)

if __name__ == "__main__":
    # Open the local web interface automatically when running the packaged app.
    threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)
