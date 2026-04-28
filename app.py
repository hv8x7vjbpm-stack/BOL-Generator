import os
import re
import json
import tempfile
from datetime import datetime
from flask import Flask, render_template, request, send_file, jsonify
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pypdf import PdfWriter
import pdfplumber

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

app = Flask(__name__)
BOL_STORE_PATH = os.path.join(os.path.dirname(__file__), "bol_store.json")

NAVY  = colors.HexColor("#1a365d")
LGREY = colors.HexColor("#f0f2f5")
MGREY = colors.HexColor("#d1d5db")

def load_store():
    if os.path.exists(BOL_STORE_PATH):
        with open(BOL_STORE_PATH, "r") as f:
            return json.load(f)
    return []

def save_store(bols):
    with open(BOL_STORE_PATH, "w") as f:
        json.dump(bols, f, indent=2)

def get_next_id(bols):
    return 1 if not bols else max(b["id"] for b in bols) + 1

def clean_bol_data(d):
    out = {}
    for k, v in d.items():
        if isinstance(v, str):
            out[k] = "" if v.strip().upper() == "N/A" else v
        elif isinstance(v, list):
            out[k] = ["" if x.strip().upper() == "N/A" else x for x in v]
        else:
            out[k] = v
    return out

def _sty(name, **kw):
    base = getSampleStyleSheet()["Normal"]
    return ParagraphStyle(name, parent=base, **kw)

S_TITLE = _sty("t", fontSize=9,  leading=11, textColor=colors.white,  fontName="Helvetica-Bold", alignment=TA_CENTER)
S_LABEL = _sty("l", fontSize=7,  leading=9,  textColor=NAVY,          fontName="Helvetica-Bold")
S_VALUE = _sty("v", fontSize=8,  leading=10, textColor=colors.black,  fontName="Helvetica")
S_SMALL = _sty("s", fontSize=6,  leading=8,  textColor=colors.black,  fontName="Helvetica")
S_SMALLB= _sty("sb",fontSize=6,  leading=8,  textColor=colors.black,  fontName="Helvetica-Bold")
S_TINY  = _sty("ti",fontSize=5.5,leading=7,  textColor=colors.black,  fontName="Helvetica")

def P(txt, sty=None):
    sty = sty or S_VALUE
    return Paragraph(str(txt).replace("&","&amp;").replace("<","&lt;"), sty)

TS_HDR = TableStyle([
    ("BACKGROUND",(0,0),(-1,-1),NAVY),("TEXTCOLOR",(0,0),(-1,-1),colors.white),
    ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),
    ("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
])
TS_BOX = TableStyle([
    ("BOX",(0,0),(-1,-1),0.5,MGREY),("INNERGRID",(0,0),(-1,-1),0.5,MGREY),
    ("VALIGN",(0,0),(-1,-1),"TOP"),
    ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
    ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
])

def bt(rows, widths, extra=None):
    t = Table(rows, colWidths=widths)
    t.setStyle(TS_BOX)
    if extra:
        t.setStyle(TableStyle(extra))
    return t

def generate_bol_pdf(data, out_path):
    doc = SimpleDocTemplate(out_path, pagesize=letter,
        leftMargin=0.35*inch, rightMargin=0.35*inch,
        topMargin=0.35*inch,  bottomMargin=0.35*inch)
    W = letter[0] - 0.70*inch
    story = []

    # Header bar
    t = Table([[P("",S_TITLE), P("Bill of Lading – Short Form – Non-Negotiable",S_TITLE), P("Page 1 of 1",S_TITLE)]],
              colWidths=[W*.15, W*.70, W*.15])
    t.setStyle(TS_HDR); story.append(t)

    # Ship From / BOL Number
    story.append(bt([[
        [P("Ship From",S_LABEL), P("JACKSON POTTERY INC",S_VALUE), P("2146 EMPIRE CENTRAL",S_VALUE),
         P("DALLAS, TX 75235",S_VALUE), P("214-974-0679",S_VALUE)],
        [P("Bill of Lading Number:",S_LABEL), P(data.get("bol_number",""),S_VALUE)]
    ]], [W*.60, W*.40]))

    # Ship To / Carrier
    addr_lines = data.get("address","").replace("\r","").split("\n")
    story.append(bt([[
        [P("Ship To",S_LABEL)] + [P(l,S_VALUE) for l in addr_lines if l.strip()],
        [P("Carrier Name:",S_LABEL), P(data.get("carrier",""),S_VALUE)]
    ]], [W*.60, W*.40]))

    # Third Party / PRO / Trailer
    story.append(bt([[
        [P("Third Party Freight Charges Bill to",S_LABEL),
         P("HOMEDEPOT.COM/ATTN: FREIGHT PAYABLES",S_VALUE),
         P("2455 PACES FERRY RD  ATLANTA, GA 30339",S_VALUE)],
        [P("PRO NUMBER",S_LABEL), P(data.get("pro_number",""),S_VALUE)],
        [P("TRAILER / SEAL NUMBER",S_LABEL), P("",S_VALUE)]
    ]], [W*.50, W*.25, W*.25]))

    # TMS / Freight Terms
    story.append(bt([[
        [P("TMS ID NUMBER",S_LABEL), P(data.get("shipment",""),S_VALUE), P("DO NOT STACK PALLETS",S_SMALLB)],
        [P("Freight Charge Terms:",S_LABEL), P("Prepaid \u2610    Collect \u2610    3rd Party \u2612",S_SMALL)]
    ]], [W*.50, W*.50]))

    # Customer Order header
    t = Table([[P("Customer Order Information",S_TITLE)]], colWidths=[W])
    t.setStyle(TS_HDR); story.append(t)

    # PO header
    pw = [W*.40, W*.15, W*.15, W*.15, W*.15]
    story.append(bt([[
        P("SPECIAL INSTRUCTIONS / PO NUMBERS",S_SMALLB),
        P("# of Pallets",S_SMALLB), P("",S_SMALLB),
        P("Pallet/Slip",S_SMALLB), P("Additional Shipper Info",S_SMALLB)
    ]], pw, extra=[("BACKGROUND",(0,0),(-1,-1),LGREY)]))

    # PO rows (always 8)
    po_numbers     = data.get("po_numbers", [])
    pallets_per_po = data.get("pallets_per_po", [])
    po_rows = []
    for i in range(8):
        po_rows.append([
            P(po_numbers[i] if i < len(po_numbers) else ""),
            P(pallets_per_po[i] if i < len(pallets_per_po) else ""),
            P(""), P(""), P("")
        ])
    story.append(bt(po_rows, pw))

    # Commodity header
    cw = [W*.06,W*.10,W*.06,W*.10,W*.09,W*.05,W*.33,W*.11,W*.10]
    story.append(bt([[
        P("Qty",S_SMALLB),P("Type",S_SMALLB),P("Qty",S_SMALLB),P("Type",S_SMALLB),
        P("Weight",S_SMALLB),P("HM",S_SMALLB),P("Commodity Description",S_SMALLB),
        P("NMFC No.",S_SMALLB),P("Class",S_SMALLB)
    ]], cw, extra=[("BACKGROUND",(0,0),(-1,-1),LGREY)]))

    # Commodity data
    story.append(bt([[
        P(data.get("total_pallets","")), P("PALLETS"), P(""), P(""),
        P(data.get("total_weight","")), P(""),
        P("CERAMIC, CHINA, EARTHENWARE, PORCELAIN OR STONEWARE / POTTERY"),
        P("47500-12"), P("55")
    ]], cw))

    # COD / value row
    story.append(bt([[
        P('Where the rate is dependent on value, shippers are required to state specifically in writing the agreed or declared value of the property.',S_TINY),
        P("COD Amount: $\nFee terms: Collect \u2610  Prepaid \u2610",S_TINY)
    ]], [W*.70, W*.30]))

    # Liability note
    story.append(bt([[P("Note: Liability limitation for loss or damage in this shipment may be applicable. See 49 USC \u00a7 14706(c)(1)(A) and (B).",S_TINY)]], [W]))

    # Signatures
    story.append(bt([[
        P("Received, subject to individually determined rates or contracts agreed upon in writing between the carrier and shipper, if applicable, otherwise to the rates, classifications, and rules established by the carrier and available to the shipper on request, and to all applicable state and federal regulations.",S_TINY),
        [P("Trailer Loaded:",S_SMALLB), P("\u2612 By shipper   \u2610 By driver",S_SMALL),
         P("Freight Counted:",S_SMALLB), P("\u2612 By shipper   \u2610 By driver",S_SMALL)],
        P("Carrier Signature / Pickup Date\n\n\nCarrier acknowledges receipt of packages and required placards.",S_TINY)
    ]], [W*.38, W*.30, W*.32]))

    story.append(bt([[
        [P("Shipper Signature / Date",S_SMALLB), P("",S_VALUE),
         P("This is to certify that the above-named materials are properly classified, packaged, marked, and labeled, and are in proper condition for transportation according to the applicable regulations of the DOT.",S_TINY)],
        P("The carrier shall not make delivery of this shipment without payment of charges and all other lawful fees.\n\nShipper Signature: ___________",S_TINY)
    ]], [W*.60, W*.40]))

    doc.build(story)

def generate_bols_pdf(bols_data):
    if len(bols_data) == 1:
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp.close()
        generate_bol_pdf(bols_data[0], tmp.name)
        return tmp.name
    writer = PdfWriter()
    tmps = []
    try:
        for bol in bols_data:
            t = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            t.close(); tmps.append(t.name)
            generate_bol_pdf(bol, t.name)
            writer.append(t.name)
        out = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        out.close()
        with open(out.name, "wb") as f:
            writer.write(f)
        return out.name
    finally:
        for p in tmps:
            try: os.unlink(p)
            except: pass

def extract_bol_from_pdf(pdf_path):
    result = {"bol_number":"","address":"","carrier":"","pro_number":"",
              "shipment":"","po_numbers":[],"pallets_per_po":[],"total_pallets":"","total_weight":""}
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        full_text = page.extract_text() or ""
        tables = page.extract_tables()
    lines = full_text.split("\n")

    for i, line in enumerate(lines):
        if "bill of lading number" in line.lower():
            parts = line.split(":")
            val = parts[-1].strip() if len(parts)>1 else ""
            result["bol_number"] = val if val not in ("","INPUT") else (lines[i+1].strip() if i+1<len(lines) else "")
            break

    addr_lines, in_addr = [], False
    for line in lines:
        if "ship to" in line.lower() and not in_addr: in_addr=True; continue
        if in_addr:
            if any(x in line.lower() for x in ("third party","carrier name","pro number","freight")): break
            s = line.strip()
            if s and s not in ("INPUT","HOME DEPOT"): addr_lines.append(s)
    result["address"] = "\n".join(addr_lines[:4])

    for i, line in enumerate(lines):
        if "carrier name" in line.lower():
            parts = line.split(":")
            v = parts[-1].strip() if len(parts)>1 else ""
            if v and v not in ("","INPUT"): result["carrier"]=v
            else:
                for j in range(1,4):
                    if i+j<len(lines) and lines[i+j].strip() not in ("","INPUT"):
                        result["carrier"]=lines[i+j].strip(); break
            break

    for i, line in enumerate(lines):
        if "pro number" in line.lower():
            parts = re.split(r"pro number[:\s]*", line, flags=re.IGNORECASE)
            v = parts[-1].strip() if len(parts)>1 else ""
            result["pro_number"] = v if v not in ("","INPUT") else (lines[i+1].strip() if i+1<len(lines) else "")
            break

    for i, line in enumerate(lines):
        if "tms id" in line.lower():
            m = re.search(r"tms id number[:\s]+([\w\-]+)", line, re.IGNORECASE)
            if m: result["shipment"]=m.group(1)
            else:
                for j in range(1,4):
                    if i+j<len(lines):
                        v=lines[i+j].strip()
                        if v and v not in ("INPUT","DO NOT STACK PALLETS",""): result["shipment"]=v; break
            break

    po_numbers, pallets_per_po = [], []
    for table in tables:
        for row in (table or []):
            if not row: continue
            rt = [str(c).strip() if c else "" for c in row]
            for ci, cell in enumerate(rt):
                if re.match(r"^(PO[-\s]?\d+|\d{5,})$", cell, re.IGNORECASE):
                    po_numbers.append(cell)
                    pal = next((c2 for c2 in rt[ci+1:] if c2 and re.match(r"^\d+$",c2)),"")
                    pallets_per_po.append(pal)

    seen, upos, upals = set(), [], []
    for i, po in enumerate(po_numbers):
        if po not in seen:
            seen.add(po); upos.append(po)
            upals.append(pallets_per_po[i] if i<len(pallets_per_po) else "")
    result["po_numbers"]=upos; result["pallets_per_po"]=upals

    half = lines[len(lines)//2:]
    for line in half:
        if "pallet" in line.lower() and not result["total_pallets"]:
            m = re.search(r"\b(\d+)\b", line)
            if m: result["total_pallets"]=m.group(1)
    for line in half:
        if not result["total_weight"]:
            m = re.search(r"(\d{3,6})\s*(?:lbs?)?", line, re.IGNORECASE)
            if m and int(m.group(1))>100: result["total_weight"]=m.group(1)

    result["po_numbers"]=[p for p in result["po_numbers"] if p.strip()]
    result["pallets_per_po"]=result["pallets_per_po"][:len(result["po_numbers"])]
    return result

def build_excel_shortage(bols):
    wb = openpyxl.Workbook()
    ws_all = wb.active; ws_all.title="All Orders"
    hdr_fill=PatternFill("solid",fgColor="1A365D")
    hdr_font=Font(bold=True,color="FFFFFF",name="Arial",size=10)
    alt_fill=PatternFill("solid",fgColor="EBF2FA")
    bs=Side(style="thin",color="CCCCCC")
    cb=Border(left=bs,right=bs,top=bs,bottom=bs)
    ctr=Alignment(horizontal="center",vertical="center",wrap_text=True)
    lft=Alignment(horizontal="left",  vertical="center",wrap_text=True)

    def hdr_row(ws, hdrs, widths):
        ws.append(hdrs)
        for col in range(1,len(hdrs)+1):
            c=ws.cell(1,col); c.fill=hdr_fill; c.font=hdr_font; c.alignment=ctr; c.border=cb
        for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
        ws.row_dimensions[1].height=28; ws.freeze_panes="A2"

    def style_cells(ws, row_i, ncols, fill_override=None, bold=False):
        for col in range(1,ncols+1):
            c=ws.cell(row_i,col)
            c.fill=fill_override if fill_override else (alt_fill if row_i%2==0 else PatternFill())
            c.font=Font(name="Arial",size=10,bold=bold); c.alignment=ctr; c.border=cb

    hdr_row(ws_all,
        ["Order #","BOL Number","Ship To Address","Carrier","PRO Number",
         "TMS / Shipment","PO Numbers","Pallets per PO","Total Pallets","Total Weight (lbs)","Date Added"],
        [8,14,30,18,14,16,20,14,12,16,18])
    for ri,bol in enumerate(bols,2):
        ws_all.append([bol.get("id",""),bol.get("bol_number",""),bol.get("address",""),
            bol.get("carrier",""),bol.get("pro_number",""),bol.get("shipment",""),
            "\n".join(bol.get("po_numbers",[])),"\n".join(bol.get("pallets_per_po",[])),
            bol.get("total_pallets",""),bol.get("total_weight",""),bol.get("date_added","")])
        style_cells(ws_all,ri,11)
        ws_all.cell(ri,3).alignment=lft

    ws_po=wb.create_sheet("By PO Number")
    hdr_row(ws_po,["PO Number","Order # (BOL ID)","BOL Number","Carrier",
        "Pallets for this PO","Total Pallets","Weight","Ship To"],[18,12,14,18,16,12,14,28])

    po_map={}
    for bol in bols:
        for j,po in enumerate(bol.get("po_numbers",[])):
            if not po.strip(): continue
            po_map.setdefault(po,[]).append({
                "order_id":bol.get("id",""),"bol_number":bol.get("bol_number",""),
                "carrier":bol.get("carrier",""),
                "pallets_this_po":(bol.get("pallets_per_po") or [])[j] if j<len(bol.get("pallets_per_po",[])) else "",
                "total_pallets":bol.get("total_pallets",""),"weight":bol.get("total_weight",""),
                "address":bol.get("address","")})

    rep_fill=PatternFill("solid",fgColor="FFF3CD"); ri=2
    for po in sorted(po_map):
        entries=po_map[po]; is_rep=len(entries)>1
        for e in entries:
            ws_po.append([po,e["order_id"],e["bol_number"],e["carrier"],
                e["pallets_this_po"],e["total_pallets"],e["weight"],e["address"]])
            style_cells(ws_po,ri,8,rep_fill if is_rep else None,bold=(is_rep))
            ri+=1

    ws_sh=wb.create_sheet("Shortage Sheet")
    hdr_row(ws_sh,["PO Number","Times Ordered","BOL Numbers","Carriers",
        "Total Pallets Across Orders"],[18,14,30,30,22])
    high_fill=PatternFill("solid",fgColor="F8D7DA"); ri=2
    for po in sorted(po_map):
        entries=po_map[po]; count=len(entries)
        bols_str=", ".join(str(e["bol_number"]) for e in entries if e["bol_number"])
        cars_str=", ".join(set(str(e["carrier"]) for e in entries if e["carrier"]))
        ptot=sum(int(str(e["pallets_this_po"]).strip()) for e in entries
                 if str(e.get("pallets_this_po","")).strip().isdigit())
        ws_sh.append([po,count,bols_str,cars_str,ptot if ptot else ""])
        style_cells(ws_sh,ri,5,high_fill if count>1 else None,bold=(count>1))
        ri+=1
    ws_sh.cell(ri+1,1).value="Red rows = PO ordered on multiple BOLs (potential shortage)"
    ws_sh.cell(ri+1,1).font=Font(name="Arial",size=9,italic=True,color="856404")

    tmp=tempfile.NamedTemporaryFile(suffix=".xlsx",delete=False)
    wb.save(tmp.name); tmp.close(); return tmp.name

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/bols", methods=["GET"])
def list_bols():
    return jsonify(load_store())

@app.route("/api/bols", methods=["POST"])
def add_bols():
    data=request.get_json(); bols_input=data.get("bols",[]); store=load_store(); added=[]
    for bd in bols_input:
        c=clean_bol_data(bd); c["id"]=get_next_id(store)
        c["date_added"]=datetime.now().strftime("%Y-%m-%d %H:%M")
        store.append(c); added.append(c)
    save_store(store)
    return jsonify({"added":len(added),"bols":added})

@app.route("/api/bols/<int:bol_id>", methods=["DELETE"])
def delete_bol(bol_id):
    save_store([b for b in load_store() if b["id"]!=bol_id])
    return jsonify({"ok":True})

@app.route("/api/bols/clear", methods=["POST"])
def clear_bols():
    save_store([]); return jsonify({"ok":True})

@app.route("/generate", methods=["POST"])
def generate():
    data=request.get_json(); bols=data.get("bols",[])
    if not bols: return jsonify({"error":"No BOL data provided"}),400
    try:
        out=generate_bols_pdf([clean_bol_data(b) for b in bols])
        return send_file(out,mimetype="application/pdf",as_attachment=True,download_name="BOLs.pdf")
    except Exception as e:
        return jsonify({"error":str(e)}),500

@app.route("/generate/store", methods=["POST"])
def generate_from_store():
    data=request.get_json(); ids=set(data.get("ids",[])); store=load_store()
    bols=[b for b in store if b["id"] in ids] if ids else store
    if not bols: return jsonify({"error":"No BOLs found"}),400
    try:
        out=generate_bols_pdf(bols)
        return send_file(out,mimetype="application/pdf",as_attachment=True,download_name="BOLs.pdf")
    except Exception as e:
        return jsonify({"error":str(e)}),500

@app.route("/import/pdfs", methods=["POST"])
def import_pdfs():
    files=request.files.getlist("files")
    if not files: return jsonify({"error":"No files uploaded"}),400
    results,errors=[],[]
    for f in files:
        if not f.filename.lower().endswith(".pdf"):
            errors.append(f"{f.filename}: not a PDF"); continue
        tmp=tempfile.NamedTemporaryFile(suffix=".pdf",delete=False)
        try:
            f.save(tmp.name); tmp.close()
            ext=extract_bol_from_pdf(tmp.name); ext["source_filename"]=f.filename
            results.append(ext)
        except Exception as e:
            errors.append(f"{f.filename}: {str(e)}")
        finally:
            try: os.unlink(tmp.name)
            except: pass
    return jsonify({"extracted":results,"errors":errors})

@app.route("/export/excel", methods=["POST"])
def export_excel():
    data=request.get_json(); ids=set(data.get("ids",[])) if data and data.get("ids") else None
    store=load_store(); bols=[b for b in store if b["id"] in ids] if ids else store
    if not bols: return jsonify({"error":"No BOLs to export"}),400
    return send_file(build_excel_shortage(bols),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,download_name="BOL_Shortage_Sheet.xlsx")

if __name__=="__main__":
    app.run(host="0.0.0.0",port=5001,debug=True)