import os
import json
import tempfile
import subprocess
from datetime import datetime
from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
from docx.table import Table
from PyPDF2 import PdfMerger
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "BOL INPUT.docx")
NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# In-memory BOL store (persisted to a JSON file)
BOL_STORE_PATH = os.path.join(os.path.dirname(__file__), "bol_store.json")


def load_store():
    if os.path.exists(BOL_STORE_PATH):
        with open(BOL_STORE_PATH, "r") as f:
            return json.load(f)
    return []


def save_store(bols):
    with open(BOL_STORE_PATH, "w") as f:
        json.dump(bols, f, indent=2)


def get_next_id(bols):
    if not bols:
        return 1
    return max(b["id"] for b in bols) + 1


def get_nested_table(cell):
    nested_elems = cell._tc.findall(f"./{NS}tbl")
    if nested_elems:
        return Table(nested_elems[0], cell._tc)
    return None


def set_run_text(paragraph, text):
    if paragraph.runs:
        paragraph.runs[0].text = text
    else:
        paragraph.add_run(text)


def fill_bol(data):
    doc = Document(TEMPLATE_PATH)
    table = doc.tables[0]

    bol_number = data.get("bol_number", "")
    cell = table.rows[2].cells[9]
    set_run_text(cell.paragraphs[0], bol_number)

    address = data.get("address", "")
    cell = table.rows[4].cells[0]
    nested = get_nested_table(cell)
    if nested:
        nested_cell = nested.rows[0].cells[0]
        if len(nested_cell.paragraphs) >= 3:
            set_run_text(nested_cell.paragraphs[2], address)

    carrier = data.get("carrier", "")
    cell = table.rows[4].cells[9]
    if len(cell.paragraphs) >= 2:
        set_run_text(cell.paragraphs[1], carrier)

    pro_number = data.get("pro_number", "")
    cell = table.rows[6].cells[9]
    set_run_text(cell.paragraphs[0], pro_number)

    shipment = data.get("shipment", "")
    cell = table.rows[7].cells[0]
    if len(cell.paragraphs) >= 2:
        set_run_text(cell.paragraphs[1], shipment)

    po_numbers = data.get("po_numbers", [])
    pallets_per_po = data.get("pallets_per_po", [])

    for i, po in enumerate(po_numbers):
        if i > 7:
            break
        row_idx = 11 + i
        if i == 0:
            cell = table.rows[11].cells[0]
            nested = get_nested_table(cell)
            if nested:
                nested_cell = nested.rows[0].cells[0]
                set_run_text(nested_cell.paragraphs[0], po)
        else:
            cell = table.rows[row_idx].cells[0]
            if cell.paragraphs:
                set_run_text(cell.paragraphs[0], po)

        if i < len(pallets_per_po):
            pallet_val = pallets_per_po[i]
            if i == 0:
                cell7 = table.rows[11].cells[7]
                set_run_text(cell7.paragraphs[0], pallet_val)
            else:
                cell7 = table.rows[row_idx].cells[7]
                if cell7.paragraphs:
                    set_run_text(cell7.paragraphs[0], pallet_val)

    total_pallets = data.get("total_pallets", "")
    cell = table.rows[20].cells[0]
    set_run_text(cell.paragraphs[0], total_pallets)

    total_weight = data.get("total_weight", "")
    cell = table.rows[20].cells[4]
    set_run_text(cell.paragraphs[0], total_weight)

    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    doc.save(tmp.name)
    tmp.close()
    return tmp.name


def find_libreoffice():
    """Find the LibreOffice binary across Mac, Linux, and Windows."""
    candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # Mac standard
        "/usr/bin/libreoffice",                                    # Linux
        "/usr/bin/soffice",                                        # Linux alt
        "libreoffice",                                             # Homebrew / PATH
        "soffice",                                                 # Windows PATH
        r"C:\Program Files\LibreOffice\program\soffice.exe",      # Windows default
    ]
    for path in candidates:
        try:
            subprocess.run([path, "--version"], capture_output=True, timeout=10)
            return path
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
    raise RuntimeError(
        "LibreOffice not found. Please install it from https://www.libreoffice.org/download/"
    )


def docx_to_pdf(docx_path):
    soffice = find_libreoffice()
    out_dir = os.path.dirname(docx_path)
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
        check=True, capture_output=True, timeout=60,
    )
    return docx_path.rsplit(".", 1)[0] + ".pdf"


def merge_pdfs(pdf_paths, output_path):
    merger = PdfMerger()
    for pdf in pdf_paths:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()


def clean_bol_data(bol_data):
    cleaned = {}
    for key, val in bol_data.items():
        if isinstance(val, str):
            cleaned[key] = "" if val.strip().upper() == "N/A" else val
        elif isinstance(val, list):
            cleaned[key] = ["" if v.strip().upper() == "N/A" else v for v in val]
        else:
            cleaned[key] = val
    return cleaned


def build_excel_shortage(bols):
    """Build an Excel workbook from BOL store data."""
    wb = openpyxl.Workbook()

    # ── Sheet 1: All Orders ──────────────────────────────────────────────
    ws_all = wb.active
    ws_all.title = "All Orders"

    hdr_fill = PatternFill("solid", fgColor="1A365D")
    hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    alt_fill = PatternFill("solid", fgColor="EBF2FA")
    border_side = Side(style="thin", color="CCCCCC")
    cell_border = Border(
        left=border_side, right=border_side,
        top=border_side, bottom=border_side
    )
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    headers = [
        "Order #", "BOL Number", "Ship To Address", "Carrier",
        "PRO Number", "TMS / Shipment", "PO Numbers", "Pallets per PO",
        "Total Pallets", "Total Weight (lbs)", "Date Added"
    ]
    ws_all.append(headers)
    for col, h in enumerate(headers, 1):
        c = ws_all.cell(1, col)
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = center
        c.border = cell_border

    col_widths = [8, 14, 30, 18, 14, 16, 20, 14, 12, 16, 18]
    for i, w in enumerate(col_widths, 1):
        ws_all.column_dimensions[get_column_letter(i)].width = w
    ws_all.row_dimensions[1].height = 28
    ws_all.freeze_panes = "A2"

    for row_i, bol in enumerate(bols, 2):
        po_str = "\n".join(bol.get("po_numbers", []))
        pal_str = "\n".join(bol.get("pallets_per_po", []))
        row_data = [
            bol.get("id", ""),
            bol.get("bol_number", ""),
            bol.get("address", ""),
            bol.get("carrier", ""),
            bol.get("pro_number", ""),
            bol.get("shipment", ""),
            po_str,
            pal_str,
            bol.get("total_pallets", ""),
            bol.get("total_weight", ""),
            bol.get("date_added", ""),
        ]
        ws_all.append(row_data)
        fill = alt_fill if row_i % 2 == 0 else None
        for col in range(1, len(headers) + 1):
            c = ws_all.cell(row_i, col)
            if fill:
                c.fill = fill
            c.font = Font(name="Arial", size=10)
            c.alignment = left if col == 3 else center
            c.border = cell_border

    # ── Sheet 2: Collated by PO ──────────────────────────────────────────
    ws_po = wb.create_sheet("By PO Number")
    po_headers = ["PO Number", "Order # (BOL ID)", "BOL Number", "Carrier", "Pallets for this PO", "Total Pallets", "Weight", "Ship To"]
    ws_po.append(po_headers)
    for col, h in enumerate(po_headers, 1):
        c = ws_po.cell(1, col)
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = center
        c.border = cell_border

    po_col_widths = [18, 12, 14, 18, 16, 12, 14, 28]
    for i, w in enumerate(po_col_widths, 1):
        ws_po.column_dimensions[get_column_letter(i)].width = w
    ws_po.row_dimensions[1].height = 28
    ws_po.freeze_panes = "A2"

    # Collate: group rows by PO number
    po_map = {}
    for bol in bols:
        po_numbers = bol.get("po_numbers", [])
        pallets_per_po = bol.get("pallets_per_po", [])
        for j, po in enumerate(po_numbers):
            if not po.strip():
                continue
            if po not in po_map:
                po_map[po] = []
            po_map[po].append({
                "order_id": bol.get("id", ""),
                "bol_number": bol.get("bol_number", ""),
                "carrier": bol.get("carrier", ""),
                "pallets_this_po": pallets_per_po[j] if j < len(pallets_per_po) else "",
                "total_pallets": bol.get("total_pallets", ""),
                "weight": bol.get("total_weight", ""),
                "address": bol.get("address", ""),
            })

    # Sort POs and write — highlight repeated POs
    repeat_fill = PatternFill("solid", fgColor="FFF3CD")
    row_i = 2
    for po in sorted(po_map.keys()):
        entries = po_map[po]
        is_repeat = len(entries) > 1
        for entry in entries:
            ws_po.append([
                po,
                entry["order_id"],
                entry["bol_number"],
                entry["carrier"],
                entry["pallets_this_po"],
                entry["total_pallets"],
                entry["weight"],
                entry["address"],
            ])
            for col in range(1, len(po_headers) + 1):
                c = ws_po.cell(row_i, col)
                c.fill = repeat_fill if is_repeat else (alt_fill if row_i % 2 == 0 else PatternFill())
                c.font = Font(name="Arial", size=10, bold=(col == 1 and is_repeat))
                c.alignment = center
                c.border = cell_border
            row_i += 1

    # ── Sheet 3: Shortage Summary ────────────────────────────────────────
    ws_sh = wb.create_sheet("Shortage Sheet")
    sh_headers = ["PO Number", "Times Ordered", "BOL Numbers", "Carriers", "Total Pallets Across Orders"]
    ws_sh.append(sh_headers)
    for col, h in enumerate(sh_headers, 1):
        c = ws_sh.cell(1, col)
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = center
        c.border = cell_border

    sh_col_widths = [18, 14, 30, 30, 22]
    for i, w in enumerate(sh_col_widths, 1):
        ws_sh.column_dimensions[get_column_letter(i)].width = w
    ws_sh.row_dimensions[1].height = 28
    ws_sh.freeze_panes = "A2"

    high_fill = PatternFill("solid", fgColor="F8D7DA")  # red tint for repeated
    row_i = 2
    for po in sorted(po_map.keys()):
        entries = po_map[po]
        count = len(entries)
        bol_numbers = ", ".join(str(e["bol_number"]) for e in entries if e["bol_number"])
        carriers = ", ".join(set(str(e["carrier"]) for e in entries if e["carrier"]))
        # Sum up pallets where numeric
        pallet_total = 0
        for e in entries:
            try:
                pallet_total += int(str(e["pallets_this_po"]).strip())
            except (ValueError, AttributeError):
                pass

        ws_sh.append([po, count, bol_numbers, carriers, pallet_total if pallet_total else ""])
        for col in range(1, len(sh_headers) + 1):
            c = ws_sh.cell(row_i, col)
            c.fill = high_fill if count > 1 else (alt_fill if row_i % 2 == 0 else PatternFill())
            c.font = Font(name="Arial", size=10, bold=(count > 1))
            c.alignment = center
            c.border = cell_border
        row_i += 1

    # Add a legend note
    ws_sh.cell(row_i + 1, 1).value = "🟥 Yellow = PO ordered on multiple BOLs (potential shortage)"
    ws_sh.cell(row_i + 1, 1).font = Font(name="Arial", size=9, italic=True, color="856404")

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    wb.save(tmp.name)
    tmp.close()
    return tmp.name


# ── Routes ────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/bols", methods=["GET"])
def list_bols():
    return jsonify(load_store())


@app.route("/api/bols", methods=["POST"])
def add_bols():
    data = request.get_json()
    bols_input = data.get("bols", [])
    store = load_store()
    added = []
    for bol_data in bols_input:
        cleaned = clean_bol_data(bol_data)
        cleaned["id"] = get_next_id(store)
        cleaned["date_added"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        store.append(cleaned)
        added.append(cleaned)
    save_store(store)
    return jsonify({"added": len(added), "bols": added})


@app.route("/api/bols/<int:bol_id>", methods=["DELETE"])
def delete_bol(bol_id):
    store = load_store()
    store = [b for b in store if b["id"] != bol_id]
    save_store(store)
    return jsonify({"ok": True})


@app.route("/api/bols/clear", methods=["POST"])
def clear_bols():
    save_store([])
    return jsonify({"ok": True})


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    bols = data.get("bols", [])
    if not bols:
        return jsonify({"error": "No BOL data provided"}), 400

    pdf_paths = []
    docx_paths = []
    try:
        for bol_data in bols:
            cleaned = clean_bol_data(bol_data)
            docx_path = fill_bol(cleaned)
            docx_paths.append(docx_path)
            pdf_path = docx_to_pdf(docx_path)
            pdf_paths.append(pdf_path)

        output_path = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name
        if len(pdf_paths) == 1:
            output_path = pdf_paths[0]
        else:
            merge_pdfs(pdf_paths, output_path)

        return send_file(output_path, mimetype="application/pdf",
                         as_attachment=True, download_name="BOLs.pdf")
    finally:
        for p in docx_paths:
            try:
                os.unlink(p)
            except Exception:
                pass


@app.route("/generate/store", methods=["POST"])
def generate_from_store():
    """Generate PDFs from selected BOL IDs in the store."""
    data = request.get_json()
    ids = set(data.get("ids", []))
    store = load_store()
    bols = [b for b in store if b["id"] in ids] if ids else store

    if not bols:
        return jsonify({"error": "No BOLs found"}), 400

    pdf_paths = []
    docx_paths = []
    try:
        for bol_data in bols:
            docx_path = fill_bol(bol_data)
            docx_paths.append(docx_path)
            pdf_path = docx_to_pdf(docx_path)
            pdf_paths.append(pdf_path)

        output_path = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name
        if len(pdf_paths) == 1:
            output_path = pdf_paths[0]
        else:
            merge_pdfs(pdf_paths, output_path)

        return send_file(output_path, mimetype="application/pdf",
                         as_attachment=True, download_name="BOLs.pdf")
    finally:
        for p in docx_paths:
            try:
                os.unlink(p)
            except Exception:
                pass


@app.route("/export/excel", methods=["POST"])
def export_excel():
    data = request.get_json()
    ids = set(data.get("ids", [])) if data and data.get("ids") else None
    store = load_store()
    bols = [b for b in store if b["id"] in ids] if ids else store

    if not bols:
        return jsonify({"error": "No BOLs to export"}), 400

    xlsx_path = build_excel_shortage(bols)
    return send_file(xlsx_path, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name="BOL_Shortage_Sheet.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)