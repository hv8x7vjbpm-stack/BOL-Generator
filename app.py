import os, re, json, tempfile
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
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

app = Flask(__name__)
BOL_STORE_PATH = os.path.join(os.path.dirname(__file__), "bol_store.json")

# ── Exact colours from template ───────────────────────────────────────────────
GREY_HDR = colors.HexColor("#E6E6E6")
BLACK    = colors.black
WHITE    = colors.white

# ── Paragraph styles ──────────────────────────────────────────────────────────
def S(name, font="Helvetica", size=8, leading=None, color=BLACK, align=TA_LEFT):
    return ParagraphStyle(name, fontName=font, fontSize=size,
                          leading=leading or max(size+1.5, size*1.2),
                          textColor=color, alignment=align, wordWrap='CJK',
                          spaceBefore=0, spaceAfter=0)

sLabel  = S("lbl",  size=7.5)
sLabelC = S("lblc", size=7.5, align=TA_CENTER)
sNorm   = S("nrm",  size=8.5)
sNormC  = S("nrmc", size=8.5, align=TA_CENTER)
sNormB  = S("nrmb", size=8.5, font="Helvetica-Bold")
sBig    = S("big",  size=11)
sBigB   = S("bigb", size=11,  font="Helvetica-Bold")
sSmall  = S("sm",   size=6.5)
sSmallB = S("smb",  size=6.5, font="Helvetica-Bold")
sDontStack = S("ds", size=8,  font="Helvetica-Bold")  # Arial Black equiv - heavy not italic
sTiny   = S("ti",   size=5.5)
sTitle  = S("tt",   size=9,   font="Helvetica-Bold", color=WHITE, align=TA_CENTER)

def p(text, style=None):
    style = style or sNorm
    safe = str(text).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    return Paragraph(safe, style)

# ── Table styles ──────────────────────────────────────────────────────────────
GRID = [
    ("BOX",           (0,0),(-1,-1), 0.5, BLACK),
    ("INNERGRID",     (0,0),(-1,-1), 0.5, BLACK),
    ("VALIGN",        (0,0),(-1,-1), "TOP"),
    ("TOPPADDING",    (0,0),(-1,-1), 2),
    ("BOTTOMPADDING", (0,0),(-1,-1), 2),
    ("LEFTPADDING",   (0,0),(-1,-1), 3),
    ("RIGHTPADDING",  (0,0),(-1,-1), 3),
]
TITLE_STYLE = GRID + [
    ("BACKGROUND",    (0,0),(-1,-1), BLACK),
    ("ALIGN",         (0,0),(-1,-1), "CENTER"),
    ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
]
GREY_STYLE = GRID + [("BACKGROUND", (0,0),(-1,-1), GREY_HDR)]

def tbl(rows, widths, style=None, row_heights=None):
    t = Table(rows, colWidths=widths, rowHeights=row_heights)
    t.setStyle(TableStyle(style or GRID))
    return t

# ── Store helpers ─────────────────────────────────────────────────────────────
def load_store():
    if os.path.exists(BOL_STORE_PATH):
        with open(BOL_STORE_PATH,"r") as f: return json.load(f)
    return []

def save_store(bols):
    with open(BOL_STORE_PATH,"w") as f: json.dump(bols,f,indent=2)

def get_next_id(bols):
    return 1 if not bols else max(b["id"] for b in bols)+1

def clean_bol_data(d):
    out={}
    for k,v in d.items():
        if isinstance(v,str):   out[k]="" if v.strip().upper()=="N/A" else v
        elif isinstance(v,list):out[k]=["" if x.strip().upper()=="N/A" else x for x in v]
        else: out[k]=v
    return out

# ── PDF generation ────────────────────────────────────────────────────────────
def generate_bol_pdf(data, out_path):
    doc = SimpleDocTemplate(out_path, pagesize=letter,
        leftMargin=0.5*inch, rightMargin=0.5*inch,
        topMargin=0.5*inch,  bottomMargin=0.5*inch)

    W = 7.5 * inch  # usable width

    # ── Column width sets (scaled from exact EMU measurements) ────────────────
    # Main 2-col split: 3.800 / 3.828 scaled to W=7.5in
    SCALE2 = W / ((3.800 + 3.828) * inch)
    C_L = 3.800 * inch * SCALE2
    C_R = 3.828 * inch * SCALE2

    # PRO / Trailer: equal halves of C_R
    C_PRO = C_R / 2

    # PO columns: 2.925, 0.875, 0.5125, 0.800, 2.5153
    PO_RAW = [2.925, 0.875, 0.5125, 0.800, 2.5153]
    PO_SCALE = W / (sum(PO_RAW) * inch)
    PO_W = [x * inch * PO_SCALE for x in PO_RAW]

    # Commodity columns: exact from template
    COM_RAW = [0.3125, 0.6417, 0.3826, 0.5382, 0.5000, 0.4688, 3.4799, 0.6521, 0.6521]
    COM_SCALE = W / (sum(COM_RAW) * inch)
    COM_W = [x * inch * COM_SCALE for x in COM_RAW]

    # COD split: 4.1125 / 3.5153
    COD_SCALE = W / ((4.1125 + 3.5153) * inch)
    C_VAL = 4.1125 * inch * COD_SCALE
    C_COD = 3.5153 * inch * COD_SCALE

    # Received / sig split: 3.3625 / 4.2653
    REC_SCALE = W / ((3.3625 + 4.2653) * inch)
    C_REC = 3.3625 * inch * REC_SCALE
    C_SIG = 4.2653 * inch * REC_SCALE

    # Bottom sig 4-col: 2.3750, 0.9875, 1.9139, 2.3514
    SIG_RAW = [2.3750, 0.9875, 1.9139, 2.3514]
    SIG_SCALE = W / (sum(SIG_RAW) * inch)
    SIG_W = [x * inch * SIG_SCALE for x in SIG_RAW]

    # ── Exact row heights from template (in points, converted from EMU/12700) ─
    RH = {
        'title':    23.05,
        'lbl1':     14.40,
        'content1': 74.05,
        'lbl2':     14.40,
        'content2': None,   # auto — expands with address
        'lbl3':     14.40,
        'content3': 24.45,
        'tms':      13.70,
        'master':   14.40,
        'cust_hdr': 14.40,
        'po_hdr':   10.80,
        'po_row':   7.05,
        'com_hdr':  24.50,
        'com_data': 14.40,
        'cod':       8.80,
        'liab':     10.80,
        'recv':     36.00,
        'sig':      53.50,
    }

    def pt(key):
        v = RH[key]
        return v if v is None else v

    # ── Pull data ─────────────────────────────────────────────────────────────
    bol_number     = data.get("bol_number", "")
    address        = data.get("address", "")
    carrier        = data.get("carrier", "")
    pro_number     = data.get("pro_number", "")
    shipment       = data.get("shipment", "")
    po_numbers     = data.get("po_numbers", [])
    pallets_per_po = data.get("pallets_per_po", [])
    total_pallets  = data.get("total_pallets", "")
    total_weight   = data.get("total_weight", "")

    story = []

    # ── ROW 0: Title bar ──────────────────────────────────────────────────────
    story.append(tbl([[
        p("", sTitle),
        p("Bill of Lading \u2013 Short Form \u2013 Non-Negotiable", sTitle),
        p("Page 1 of 1", sTitle),
    ]], [W*1.3368/7.6278, W*4.5257/7.6278, W*1.7653/7.6278],
    style=TITLE_STYLE, row_heights=[pt('title')]))

    # ── ROW 1: Ship From label | Bill of Lading Number label ─────────────────
    story.append(tbl([[
        p("Ship From", sLabel),
        p("Bill of Lading Number:", sLabel),
    ]], [C_L, C_R],
    style=GREY_STYLE + [("BACKGROUND",(1,0),(1,0), WHITE)],
    row_heights=[pt('lbl1')]))

    # ── ROW 2: Ship From content | BOL Number value ───────────────────────────
    story.append(tbl([[
        [p("JACKSON POTTERY INC", sNorm),
         p("2146 EMPIRE CENTRAL", sNorm),
         p("DALLAS, TX 75235", sNorm),
         p("214-974-0679", sNorm)],
        p(bol_number, sBig),
    ]], [C_L, C_R], row_heights=[pt('content1')]))

    # ── ROW 3: Ship To label | Carrier Name label ─────────────────────────────
    story.append(tbl([[
        p("Ship To", sLabel),
        p("Carrier Name:", sLabel),
    ]], [C_L, C_R],
    style=GREY_STYLE + [("BACKGROUND",(1,0),(1,0), WHITE)],
    row_heights=[pt('lbl2')]))

    # ── ROW 4: Ship To content | Carrier value ────────────────────────────────
    # Address lines bold like the real BOL — fixed min height so empty fields
    # don't collapse the row
    addr_lines = address.replace("\r","").split("\n")
    ship_to_content = [p("HOME DEPOT \u2013 Store", sBigB)]
    for l in addr_lines:
        if l.strip():
            ship_to_content.append(p(l.strip(), sNormB))
    story.append(tbl([[
        ship_to_content,
        p(carrier, sBigB),
    ]], [C_L, C_R], row_heights=[pt('content2')]))

    # ── ROW 5: Third Party label | PRO label | Trailer label ─────────────────
    story.append(tbl([[
        p("Third Party Freight Charges Bill to", sLabel),
        p("PRO NUMBER", sLabel),
        p("TRAILER / SEAL NUMBER", sLabel),
    ]], [C_L, C_PRO, C_PRO],
    style=GREY_STYLE, row_heights=[pt('lbl3')]))

    # ── ROW 6: Third Party content | PRO value | Trailer value ───────────────
    story.append(tbl([[
        [p("HOMEDEPOT.COM/ATTN: FREIGHT PAYABLES", sNorm),
         p("2455 PACES FERRY RD", sNorm),
         p("ATLANTA, GA 30339", sNorm)],
        p(pro_number, sNorm),
        p("", sNorm),
    ]], [C_L, C_PRO, C_PRO], row_heights=[pt('content3')]))

    # ── ROW 7: TMS ID | Freight terms ────────────────────────────────────────
    story.append(tbl([[
        [p("TMS ID NUMBER", sLabel),
         p(shipment, sBig),
         p("DO NOT STACK PALLETS", sDontStack)],
        [p("Freight Charge Terms (Freight charges are prepaid unless marked otherwise):", sSmall),
         p("Prepaid \u2610  Collect \u2610  3rd Party  x", sSmall)],
    ]], [C_L, C_R], row_heights=[pt('tms')]))

    # ── ROW 8: Master bill row ────────────────────────────────────────────────
    story.append(tbl([[
        p("", sNorm),
        p("\u2610 Master bill of lading with attached underlying bills of lading.", sSmall),
    ]], [C_L, C_R], row_heights=[pt('master')]))

    # ── ROW 9: Customer Order Information header ──────────────────────────────
    story.append(tbl([[
        p("Customer Order Information", sTitle),
    ]], [W], style=TITLE_STYLE, row_heights=[pt('cust_hdr')]))

    # ── ROW 10: PO table header ───────────────────────────────────────────────
    story.append(tbl([[
        p("SPECIAL INSTRUCTIONS\nPO NUMBERS", sLabel),
        p("# of Pallets", sLabelC),
        p("", sLabel),
        p("Pallet/Slip\n(circle one)", sLabelC),
        p("Additional Shipper Information", sLabel),
    ]], PO_W, style=GREY_STYLE, row_heights=[pt('po_hdr')]))

    # ── ROWS 11-18: 8 PO data rows with fixed height ─────────────────────────
    # Fixed height means empty rows look identical to filled rows
    for i in range(8):
        pov = po_numbers[i]     if i < len(po_numbers)     else ""
        pav = pallets_per_po[i] if i < len(pallets_per_po) else ""
        story.append(tbl([[
            p(pov, sNorm), p(pav, sNormC), p(""), p(""), p(""),
        ]], PO_W, row_heights=[pt('po_row')]))

    # ── ROW 19: Commodity header ──────────────────────────────────────────────
    story.append(tbl([[
        p("Qty", sLabelC), p("Type", sLabel),
        p("Qty", sLabelC), p("Type", sLabel),
        p("Weight", sLabel), p("HM(X)", sLabelC),
        p("Commodity Description\nCommodities requiring special or additional care or "
          "attention in handling or stowing must be so marked and packaged as to ensure "
          "safe transportation with ordinary care. See Section 2(e) of NMFC item 360", sTiny),
        p("NMFC No.", sLabelC), p("Class", sLabelC),
    ]], COM_W, style=GREY_STYLE, row_heights=[pt('com_hdr')]))

    # ── ROW 20: Commodity data ────────────────────────────────────────────────
    story.append(tbl([[
        p(total_pallets, sNormC), p("PALLETS", sNorm),
        p(""), p(""),
        p(total_weight, sNormC), p(""),
        p("CERAMIC, CHINA, EARTHENWARE, PORCELAIN OR STONEWARE/ POTTERY", sNorm),
        p("47500-12", sNormC), p("55", sNormC),
    ]], COM_W, row_heights=[pt('com_data')]))

    # ── ROW 21: Value declaration | COD ──────────────────────────────────────
    story.append(tbl([[
        p('Where the rate is dependent on value, shippers are required to state specifically '
          'in writing the agreed or declared value of the property as follows: "The agreed or '
          'declared value of the property is specifically stated by the shipper to be not '
          'exceeding _______________ per _______________.', sTiny),
        [p("COD Amount: $", sTiny),
         p("Fee terms: Collect \u2610  Prepaid \u2610  Customer check acceptable \u2610", sTiny)],
    ]], [C_VAL, C_COD], row_heights=[pt('cod')]))

    # ── ROW 22: Liability note ────────────────────────────────────────────────
    story.append(tbl([[
        p("Note: Liability limitation for loss or damage in this shipment may be applicable. "
          "See 49 USC \u00a7 14706(c)(1)(A) and (B).", sTiny),
    ]], [W], row_heights=[pt('liab')]))

    # ── ROW 23: Received text | Carrier payment / Shipper sig ────────────────
    story.append(tbl([[
        p("Received, subject to individually determined rates or contracts that have been "
          "agreed upon in writing between the carrier and shipper, if applicable, otherwise "
          "to the rates, classifications, and rules that have been established by the carrier "
          "and are available to the shipper, on request, and to all applicable state and "
          "federal regulations.", sTiny),
        [p("The carrier shall not make delivery of this shipment without payment of charges "
           "and all other lawful fees.", sTiny),
         Spacer(1, 4),
         p("Shipper Signature: _________________________", sTiny)],
    ]], [C_REC, C_SIG], row_heights=[pt('recv')]))

    # ── ROW 24: Signature row ─────────────────────────────────────────────────
    story.append(tbl([[
        [p("Shipper Signature/Date", sLabel),
         Spacer(1, 6),
         p("This is to certify that the above-named materials are properly classified, "
           "packaged, marked, and labeled, and are in proper condition for transportation "
           "according to the applicable regulations of the DOT.", sTiny)],
        [p("Trailer Loaded:", sLabel),
         p("x By shipper", sSmall),
         p("\u2610 By driver", sSmall),
         Spacer(1, 3),
         p("Trailer Counted", sSmallB),
         p("x By shipper", sSmall),
         p("\u2610 By driver", sSmall)],
        [p("Freight Counted:", sLabel),
         p("x By shipper", sSmall),
         p("\u2610 By driver/pallets said to contain", sSmall),
         p("\u2610 By driver/pieces", sSmall)],
        [p("Carrier Signature/Pickup Date", sLabel),
         Spacer(1, 6),
         p("Carrier acknowledges receipt of packages and required placards. Carrier "
           "certifies emergency response information was made available and/or carrier "
           "has the DOT emergency response guidebook or equivalent documentation in the "
           "vehicle. Property described above is received in good order, except as noted.", sTiny)],
    ]], SIG_W, row_heights=[pt('sig')]))

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
        with open(out.name, "wb") as f: writer.write(f)
        return out.name
    finally:
        for p2 in tmps:
            try: os.unlink(p2)
            except: pass


def extract_bol_from_pdf(pdf_path):
    result = {"bol_number":"","address":"","carrier":"","pro_number":"",
              "shipment":"","po_numbers":[],"pallets_per_po":[],"total_pallets":"","total_weight":""}
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]; full_text = page.extract_text() or ""; tables = page.extract_tables()
    lines = full_text.split("\n")
    for i,line in enumerate(lines):
        if "bill of lading number" in line.lower():
            parts = line.split(":")
            v = parts[-1].strip() if len(parts)>1 else ""
            result["bol_number"] = v if v not in ("","INPUT") else (lines[i+1].strip() if i+1<len(lines) else "")
            break
    addr_lines, in_addr = [], False
    for line in lines:
        if "ship to" in line.lower() and not in_addr: in_addr=True; continue
        if in_addr:
            if any(x in line.lower() for x in ("third party","carrier name","pro number","freight")): break
            s = line.strip()
            if s and s not in ("INPUT","HOME DEPOT"): addr_lines.append(s)
    result["address"] = "\n".join(addr_lines[:4])
    for i,line in enumerate(lines):
        if "carrier name" in line.lower():
            parts = line.split(":")
            v = parts[-1].strip() if len(parts)>1 else ""
            if v and v not in ("","INPUT"): result["carrier"]=v
            else:
                for j in range(1,4):
                    if i+j<len(lines) and lines[i+j].strip() not in ("","INPUT"):
                        result["carrier"]=lines[i+j].strip(); break
            break
    for i,line in enumerate(lines):
        if "pro number" in line.lower():
            parts = re.split(r"pro number[:\s]*",line,flags=re.IGNORECASE)
            v = parts[-1].strip() if len(parts)>1 else ""
            result["pro_number"] = v if v not in ("","INPUT") else (lines[i+1].strip() if i+1<len(lines) else "")
            break
    for i,line in enumerate(lines):
        if "tms id" in line.lower():
            m = re.search(r"tms id number[:\s]+([\w\-]+)",line,re.IGNORECASE)
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
            for ci,cell in enumerate(rt):
                if re.match(r"^(PO[-\s]?\d+|\d{5,})$",cell,re.IGNORECASE):
                    po_numbers.append(cell)
                    pal = next((c2 for c2 in rt[ci+1:] if c2 and re.match(r"^\d+$",c2)),"")
                    pallets_per_po.append(pal)
    seen,upos,upals = set(),[],[]
    for i,po in enumerate(po_numbers):
        if po not in seen:
            seen.add(po); upos.append(po)
            upals.append(pallets_per_po[i] if i<len(pallets_per_po) else "")
    result["po_numbers"]=upos; result["pallets_per_po"]=upals
    half = lines[len(lines)//2:]
    for line in half:
        if "pallet" in line.lower() and not result["total_pallets"]:
            m = re.search(r"\b(\d+)\b",line)
            if m: result["total_pallets"]=m.group(1)
    for line in half:
        if not result["total_weight"]:
            m = re.search(r"(\d{3,6})\s*(?:lbs?)?",line,re.IGNORECASE)
            if m and int(m.group(1))>100: result["total_weight"]=m.group(1)
    result["po_numbers"]=[p2 for p2 in result["po_numbers"] if p2.strip()]
    result["pallets_per_po"]=result["pallets_per_po"][:len(result["po_numbers"])]
    return result


def build_excel_shortage(bols):
    wb=openpyxl.Workbook(); ws_all=wb.active; ws_all.title="All Orders"
    hf=PatternFill("solid",fgColor="1A365D"); hfont=Font(bold=True,color="FFFFFF",name="Arial",size=10)
    af=PatternFill("solid",fgColor="EBF2FA"); bs=Side(style="thin",color="CCCCCC")
    cb=Border(left=bs,right=bs,top=bs,bottom=bs)
    ctr=Alignment(horizontal="center",vertical="center",wrap_text=True)
    lft=Alignment(horizontal="left",  vertical="center",wrap_text=True)
    def make_hdr(ws,hdrs,widths):
        ws.append(hdrs)
        for col in range(1,len(hdrs)+1):
            c=ws.cell(1,col); c.fill=hf; c.font=hfont; c.alignment=ctr; c.border=cb
        for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
        ws.row_dimensions[1].height=28; ws.freeze_panes="A2"
    def style_cells(ws,ri,ncols,fill=None,bold=False):
        for col in range(1,ncols+1):
            c=ws.cell(ri,col)
            c.fill=fill if fill else (af if ri%2==0 else PatternFill())
            c.font=Font(name="Arial",size=10,bold=bold); c.alignment=ctr; c.border=cb
    make_hdr(ws_all,["Order #","BOL Number","Ship To Address","Carrier","PRO Number",
        "TMS / Shipment","PO Numbers","Pallets per PO","Total Pallets","Total Weight (lbs)","Date Added"],
        [8,14,30,18,14,16,20,14,12,16,18])
    for ri,bol in enumerate(bols,2):
        ws_all.append([bol.get("id",""),bol.get("bol_number",""),bol.get("address",""),
            bol.get("carrier",""),bol.get("pro_number",""),bol.get("shipment",""),
            "\n".join(bol.get("po_numbers",[])),"\n".join(bol.get("pallets_per_po",[])),
            bol.get("total_pallets",""),bol.get("total_weight",""),bol.get("date_added","")])
        style_cells(ws_all,ri,11); ws_all.cell(ri,3).alignment=lft
    ws_po=wb.create_sheet("By PO Number")
    make_hdr(ws_po,["PO Number","Order # (BOL ID)","BOL Number","Carrier",
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
    rf=PatternFill("solid",fgColor="FFF3CD"); ri=2
    for po in sorted(po_map):
        entries=po_map[po]; is_rep=len(entries)>1
        for e in entries:
            ws_po.append([po,e["order_id"],e["bol_number"],e["carrier"],
                e["pallets_this_po"],e["total_pallets"],e["weight"],e["address"]])
            style_cells(ws_po,ri,8,rf if is_rep else None,bold=is_rep); ri+=1
    ws_sh=wb.create_sheet("Shortage Sheet")
    make_hdr(ws_sh,["PO Number","Times Ordered","BOL Numbers","Carriers",
        "Total Pallets Across Orders"],[18,14,30,30,22])
    hif=PatternFill("solid",fgColor="F8D7DA"); ri=2
    for po in sorted(po_map):
        entries=po_map[po]; count=len(entries)
        bstr=", ".join(str(e["bol_number"]) for e in entries if e["bol_number"])
        cstr=", ".join(set(str(e["carrier"]) for e in entries if e["carrier"]))
        ptot=sum(int(str(e["pallets_this_po"]).strip()) for e in entries
                 if str(e.get("pallets_this_po","")).strip().isdigit())
        ws_sh.append([po,count,bstr,cstr,ptot if ptot else ""])
        style_cells(ws_sh,ri,5,hif if count>1 else None,bold=(count>1)); ri+=1
    ws_sh.cell(ri+1,1).value="Red rows = PO ordered on multiple BOLs (potential shortage)"
    ws_sh.cell(ri+1,1).font=Font(name="Arial",size=9,italic=True,color="856404")
    tmp=tempfile.NamedTemporaryFile(suffix=".xlsx",delete=False)
    wb.save(tmp.name); tmp.close(); return tmp.name


@app.route("/")
def index(): return render_template("index.html")

@app.route("/api/bols",methods=["GET"])
def list_bols(): return jsonify(load_store())

@app.route("/api/bols",methods=["POST"])
def add_bols():
    data=request.get_json(); bols_input=data.get("bols",[]); store=load_store(); added=[]
    for bd in bols_input:
        c=clean_bol_data(bd); c["id"]=get_next_id(store)
        c["date_added"]=datetime.now().strftime("%Y-%m-%d %H:%M")
        store.append(c); added.append(c)
    save_store(store); return jsonify({"added":len(added),"bols":added})

@app.route("/api/bols/<int:bol_id>",methods=["DELETE"])
def delete_bol(bol_id):
    save_store([b for b in load_store() if b["id"]!=bol_id]); return jsonify({"ok":True})

@app.route("/api/bols/clear",methods=["POST"])
def clear_bols(): save_store([]); return jsonify({"ok":True})

@app.route("/generate",methods=["POST"])
def generate():
    data=request.get_json(); bols=data.get("bols",[])
    if not bols: return jsonify({"error":"No BOL data provided"}),400
    try:
        out=generate_bols_pdf([clean_bol_data(b) for b in bols])
        return send_file(out,mimetype="application/pdf",as_attachment=True,download_name="BOLs.pdf")
    except Exception as e: return jsonify({"error":str(e)}),500

@app.route("/generate/store",methods=["POST"])
def generate_from_store():
    data=request.get_json(); ids=set(data.get("ids",[])); store=load_store()
    bols=[b for b in store if b["id"] in ids] if ids else store
    if not bols: return jsonify({"error":"No BOLs found"}),400
    try:
        out=generate_bols_pdf(bols)
        return send_file(out,mimetype="application/pdf",as_attachment=True,download_name="BOLs.pdf")
    except Exception as e: return jsonify({"error":str(e)}),500

@app.route("/import/pdfs",methods=["POST"])
def import_pdfs():
    files=request.files.getlist("files")
    if not files: return jsonify({"error":"No files uploaded"}),400
    results,errors=[],[]
    for f in files:
        if not f.filename.lower().endswith(".pdf"): errors.append(f"{f.filename}: not a PDF"); continue
        tmp=tempfile.NamedTemporaryFile(suffix=".pdf",delete=False)
        try:
            f.save(tmp.name); tmp.close()
            ext=extract_bol_from_pdf(tmp.name); ext["source_filename"]=f.filename; results.append(ext)
        except Exception as e: errors.append(f"{f.filename}: {str(e)}")
        finally:
            try: os.unlink(tmp.name)
            except: pass
    return jsonify({"extracted":results,"errors":errors})

@app.route("/export/excel",methods=["POST"])
def export_excel():
    data=request.get_json(); ids=set(data.get("ids",[])) if data and data.get("ids") else None
    store=load_store(); bols=[b for b in store if b["id"] in ids] if ids else store
    if not bols: return jsonify({"error":"No BOLs to export"}),400
    return send_file(build_excel_shortage(bols),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,download_name="BOL_Shortage_Sheet.xlsx")

if __name__=="__main__":
    app.run(host="0.0.0.0",port=5001,debug=True)
