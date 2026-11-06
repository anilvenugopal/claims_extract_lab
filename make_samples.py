"""Generate a coherent set of (fictional) insurance-claims documents for testing.

    python make_samples.py              # writes into ./samples
    python make_samples.py some/folder  # writes elsewhere

Everything revolves around the same eight claims for "Harbor Mutual Insurance (fictional)":
a claims register, a loss run, a first-notice-of-loss form, adjuster notes, a settlement letter,
a quarterly review deck, a PDF loss summary, and so on - in many formats and with deliberately
awkward layouts (offset tables, side-by-side tables, merged cells, formulas, text boxes, charts,
hidden sheets, scanned pages, encrypted files, mis-named files).

Formats that cannot be generated programmatically here (.xlsb, .msg, .doc, .ppt) are skipped;
drop your own copies into samples/ to try them.
"""
from __future__ import annotations

import base64
import io
import shutil
import sys
import zipfile
from datetime import date, datetime, time as _time


def _dt(d):
    return datetime.combine(d, _time())
from pathlib import Path
from typing import Dict, List

HERE = Path(__file__).resolve().parent
PASSWORD = "secret"
VALUATION = date(2026, 4, 30)

# --------------------------------------------------------------------------- shared data
# claim_id, policy, insured, date_of_loss, cause, status, paid, reserve, loss_ratio
CLAIMS: List[tuple] = [
    ("CL-10421", "PL-77310", "Alder Logistics, LLC", date(2026, 1, 14), "Water damage", "Closed", 12500.75, 0.00, 0.6234),
    ("CL-10422", "PL-77310", "Alder Logistics, LLC", date(2026, 2, 3), "Theft", "Open", 4200.00, 8500.00, 0.4517),
    ("CL-10435", "PL-81207", "Brightwater Dental", date(2026, 2, 19), "Fire", "Open", 36750.40, 61000.00, 1.2038),
    ("CL-10440", "PL-90015", "Cobalt Freight", date(2026, 3, 2), "Collision", "Closed", 9875.10, 0.00, 0.8051),
    ("CL-10447", "PL-90015", "Cobalt Freight", date(2026, 3, 11), "Cargo loss", "Reopened", 15300.00, 7200.50, 0.9376),
    ("CL-10452", "PL-66402", "Dunmore Bakery", date(2026, 3, 25), "Equipment breakdown", "Open", 0.00, 11250.00, 0.0000),
    ("CL-10458", "PL-81207", "Brightwater Dental", date(2026, 4, 8), "Liability", "Denied", 0.00, 0.00, 0.0000),
    ("CL-10463", "PL-55128", "Elm Street Clinic", date(2026, 4, 22), "Hail", "Open", 2310.25, 5000.00, 0.3312),
]
PAYMENTS_10421 = [
    ("Indemnity", "Alder Logistics, LLC", 10000.00, date(2026, 1, 30)),
    ("Expense", "Restoration vendor", 1800.25, date(2026, 2, 6)),
    ("Expense", "Independent adjuster fee", 700.50, date(2026, 2, 12)),
]
COVERAGES = [("Property", 500000, 2500), ("General liability", 1000000, 5000), ("Business interruption", 250000, 10000)]
TOTAL_PAID = round(sum(c[6] for c in CLAIMS), 2)        # 80936.50
TOTAL_RESERVE = round(sum(c[7] for c in CLAIMS), 2)     # 92950.50

PNG_FALLBACK = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="
)


def _png_bytes() -> bytes:
    """A small placeholder 'damage photo' (falls back to a 1x1 PNG without Pillow)."""
    try:
        from PIL import Image, ImageDraw

        img = Image.new("RGB", (480, 300), (226, 232, 238))
        d = ImageDraw.Draw(img)
        d.rectangle((40, 150, 440, 270), fill=(168, 150, 128), outline=(70, 60, 50), width=3)     # building
        d.polygon([(30, 150), (240, 60), (450, 150)], fill=(120, 70, 60), outline=(70, 40, 35))     # roof
        d.rectangle((190, 190, 290, 270), fill=(90, 120, 150))                                       # door
        d.ellipse((330, 170, 410, 250), fill=(60, 110, 190))                                         # "water damage"
        d.text((16, 14), "CL-10421  water damage  (placeholder photo)", fill=(30, 30, 30))
        buf = io.BytesIO()
        img.save(buf, "PNG")
        return buf.getvalue()
    except Exception:  # pragma: no cover
        return PNG_FALLBACK


def money(x: float) -> str:
    return f"${x:,.2f}"


# --------------------------------------------------------------------------- plain text family
def _text_files(out: Path, made: Dict[str, Path]) -> None:
    png_b64 = base64.b64encode(_png_bytes()).decode()
    files = {
        "claim_notes.txt": (
            "CLAIM FILE NOTES - CL-10421 (Alder Logistics, LLC)\n"
            "2026-01-14  FNOL received by phone. Burst pipe in warehouse unit 4; stock on pallets affected.\n"
            "2026-01-16  Site inspection completed. Moisture readings elevated in 3 of 5 bays.\n"
            "2026-01-30  Indemnity payment of $10,000.00 issued to insured.\n"
            "2026-02-06  Restoration vendor invoice approved: $1,800.25.\n"
            "2026-02-12  Independent adjuster fee posted: $700.50. Total paid to date: $12,500.75.\n"
            "2026-02-20  Claim closed. No subrogation potential identified.\n"
        ),
        "claim_report.html": (
            "<!DOCTYPE html><html><head><title>Claim CL-10435</title><style>td{padding:4px}</style></head><body>"
            "<h1>Claim CL-10435 - Fire loss</h1>"
            "<p>Insured: <b>Brightwater Dental</b> &middot; Policy PL-81207 &middot; Loss date 2026-02-19</p>"
            "<table border='1'><tr><th>Item</th><th>Amount (USD)</th></tr>"
            "<tr><td>Paid to date</td><td>36,750.40</td></tr><tr><td>Outstanding reserve</td><td>61,000.00</td></tr></table>"
            "<ul><li>Fire marshal report received</li><li>Equipment appraisal pending</li></ul>"
            "<!-- internal: reserve reviewed 2026-03-30 -->"
            "</body></html>"
        ),
        "claim_feed.xml": (
            "<?xml version='1.0' encoding='UTF-8'?>\n<claims valuation='2026-04-30'>\n"
            + "".join(
                f"  <claim id='{c[0]}' policy='{c[1]}'><insured>{c[2].replace(',', '')}</insured>"
                f"<status>{c[5]}</status><paid currency='USD'>{c[6]:.2f}</paid></claim>\n"
                for c in CLAIMS[:3]
            )
            + "</claims>\n"
        ),
        "adjuster_notes.md": (
            "# Adjuster notes - CL-10447\n\n"
            "**Insured:** Cobalt Freight &nbsp;|&nbsp; **Policy:** PL-90015\n\n"
            "<!-- internal reserve rationale:\n     supplemental cargo invoice expected, hold reserve at 7,200.50 -->\n\n"
            "## Timeline\n\n"
            "1. 2026-03-11 - Loss reported (cargo shifted in transit).\n"
            "2. 2026-03-28 - Claim closed after payment of $15,300.00.\n"
            "3. 2026-04-15 - **Reopened**: supplemental invoice received.\n\n"
            "## Payments\n\n| Date | Type | Amount |\n|---|---|---|\n"
            "| 2026-03-27 | Indemnity | 14,000.00 |\n| 2026-03-28 | Expense | 1,300.00 |\n\n"
            "Signature of adjuster: ![signature](data:image/png;base64," + png_b64 + ")\n\n"
            "Scene photo: ![scene](https://example.com/photos/CL-10447.png)\n"
        ),
        "settlement_letter.rtf": (
            r"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}"
            r"{\b Harbor Mutual Insurance}\par "
            r"Re: Claim CL-10421 \endash  final settlement\par\par "
            r"Dear Alder Logistics, LLC,\par "
            r"We confirm payment of {\b $12,500.75} in full and final settlement of the above claim "
            r"for water damage reported on 14 January 2026.\par "
            r"Sincerely,\par{\i Claims Department}\par}"
        ),
    }
    for name, body in files.items():
        p = out / name
        p.write_text(body, encoding="utf-8")
        made[name] = p

    # CSV with quoted commas, thousands separators stored as text, blanks and a literal "N/A"
    rows = ["Claim ID,Policy,Insured,Date of Loss,Cause,Status,Paid,Reserve,Adjuster"]
    adjusters = ["M. Okafor", "M. Okafor", "L. Petrov", "R. Singh", "R. Singh", "", "L. Petrov", "N/A"]
    for c, adj in zip(CLAIMS, adjusters):
        paid = f"\"{c[6]:,.2f}\"" if c[6] >= 10000 else f"{c[6]:.2f}"
        rows.append(f"{c[0]},{c[1]},\"{c[2]}\",{c[3].isoformat()},{c[4]},{c[5]},{paid},{c[7]:.2f},{adj}")
    p = out / "claims_register.csv"
    p.write_text("\n".join(rows) + "\n", encoding="utf-8")
    made[p.name] = p


# --------------------------------------------------------------------------- Word
def _docx(out: Path, made: Dict[str, Path]) -> None:
    import docx

    d = docx.Document()
    d.add_heading("First Notice of Loss", level=0)
    d.add_paragraph("Claim CL-10421 · Policy PL-77310 · Harbor Mutual Insurance (fictional)")

    d.add_heading("1. Insured and policy", level=2)
    t = d.add_table(rows=5, cols=2)
    t.style = "Table Grid"
    for r, (k, v) in enumerate([
        ("Insured", "Alder Logistics, LLC"), ("Policy number", "PL-77310"), ("Date of loss", "14 January 2026"),
        ("Loss location", "Unit 4, 220 Dockside Road"), ("Reported by", "Facilities manager (telephone)"),
    ]):
        t.cell(r, 0).text, t.cell(r, 1).text = k, v

    d.add_heading("2. Description of loss", level=2)
    p = d.add_paragraph("A supply pipe above bay 3 failed overnight. ")
    p.add_run("Approximately 40 pallets of packaged goods").bold = True
    p.add_run(" were wet before the water was shut off; the insured estimates stock losses of $18,000.")
    d.add_paragraph("The insured has taken photographs and retained the damaged stock for inspection.")

    d.add_heading("3. Coverage", level=2)
    cov = d.add_table(rows=5, cols=4)
    cov.style = "Table Grid"
    cov.cell(0, 0).merge(cov.cell(0, 3)).text = "Coverage summary - amounts in USD"          # horizontal merge
    for c, h in enumerate(("Coverage", "Limit", "Deductible", "Applies")):
        cov.cell(1, c).text = h
    for r, (name, limit, ded) in enumerate(COVERAGES, start=2):
        cov.cell(r, 0).text, cov.cell(r, 1).text, cov.cell(r, 2).text = name, f"{limit:,}", f"{ded:,}"
        cov.cell(r, 3).text = "Yes" if name == "Property" else "No"

    d.add_paragraph("Documents attached:")
    for item in ("Photographs (12)", "Stock list with purchase invoices", "Plumber's report"):
        d.add_paragraph(item, style="List Bullet")

    d.add_page_break()

    d.add_heading("4. Payment history", level=2)
    pay = d.add_table(rows=5, cols=5)
    pay.style = "Table Grid"
    for c, h in enumerate(("Claim", "Type", "Payee", "Amount", "Date")):
        pay.cell(0, c).text = h
    for r, (typ, payee, amt, dt) in enumerate(PAYMENTS_10421, start=1):
        pay.cell(r, 1).text, pay.cell(r, 2).text = typ, payee
        pay.cell(r, 3).text, pay.cell(r, 4).text = f"{amt:,.2f}", dt.isoformat()
    pay.cell(1, 0).merge(pay.cell(3, 0)).text = "CL-10421"                                   # vertical merge
    pay.cell(4, 1).text, pay.cell(4, 3).text = "Total", "12,500.75"

    d.add_heading("5. Adjuster certification", level=2)
    d.add_paragraph("I certify that the information above is accurate to the best of my knowledge.")
    d.add_paragraph("Signed: M. Okafor, Senior Adjuster        Date: 20 February 2026")

    p = out / "fnol_form.docx"
    d.save(p)
    made[p.name] = p

    # .docm = same package with the macro-enabled main content type
    docm = out / "fnol_form.docm"
    with zipfile.ZipFile(p) as zin, zipfile.ZipFile(docm, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    b"application/vnd.ms-word.document.macroEnabled.main+xml",
                )
            zout.writestr(item, data)
    made[docm.name] = docm

    noext = out / "fnol_no_extension"                       # signature-sniffing fixture
    shutil.copyfile(p, noext)
    made[noext.name] = noext


# --------------------------------------------------------------------------- Excel
def _excel(out: Path, made: Dict[str, Path]) -> None:
    import xlsxwriter

    # ---- loss_run.xlsx: ONE sheet, deliberately awkward layout ---------------------------------
    path = out / "loss_run.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Loss Run")
    title = wb.add_format({"bold": True, "font_size": 14})
    bold = wb.add_format({"bold": True})
    cur = wb.add_format({"num_format": "$#,##0.00"})
    cur_bold = wb.add_format({"num_format": "$#,##0.00", "bold": True})
    dt = wb.add_format({"num_format": "yyyy-mm-dd"})
    pct = wb.add_format({"num_format": "0.0%"})
    ital = wb.add_format({"italic": True, "font_color": "#555555"})
    ctr = wb.add_format({"align": "center", "valign": "vcenter", "border": 1})

    ws.merge_range("B2:I2", "Loss Run - Harbor Mutual Insurance (fictional) - amounts in USD", title)     # title offset from A1
    ws.write("B3", "Valuation date:", bold)
    ws.write_datetime("C3", _dt(VALUATION), dt)

    # Table 1: an explicit Excel table, offset from A1 (B5:I13)
    ws.add_table("B5:I13", {
        "name": "Claims",
        "style": "Table Style Light 9",
        "columns": [
            {"header": "Claim ID"}, {"header": "Insured"}, {"header": "Date of loss", "format": dt},
            {"header": "Cause"}, {"header": "Status"}, {"header": "Paid", "format": cur},
            {"header": "Reserve", "format": cur}, {"header": "Loss ratio", "format": pct},
        ],
        "data": [[c[0], c[2], c[3], c[4], c[5], c[6], c[7], c[8]] for c in CLAIMS],
    })
    ws.write("B14", "Total", bold)
    ws.write_formula("G14", "=SUM(G6:G13)", cur_bold, TOTAL_PAID)       # formula WITH cached result
    ws.write_formula("H14", "=SUM(H6:H13)", cur_bold, TOTAL_RESERVE)
    ws.write_comment("G7", "Includes salvage recovery of $1,200.00")

    # Table 2: side-by-side with table 1, not an Excel table (K4:M8)
    ws.write("K4", "Coverage details", bold)
    for c, h in enumerate(("Coverage", "Limit", "Deductible")):
        ws.write(4, 10 + c, h, bold)
    for r, (name, limit, ded) in enumerate(COVERAGES, start=5):
        ws.write(r, 10, name)
        ws.write(r, 11, limit, cur)
        ws.write(r, 12, ded, cur)

    # Table 3: stacked below, with a vertically merged "group" cell and a total (B17:F22)
    ws.write("B17", "Payment history - CL-10421", bold)
    for c, h in enumerate(("Claim", "Type", "Payee", "Amount", "Date")):
        ws.write(17, 1 + c, h, bold)
    ws.merge_range("B19:B21", "CL-10421", ctr)
    for r, (typ, payee, amt, d) in enumerate(PAYMENTS_10421, start=18):
        ws.write(r, 2, typ)
        ws.write(r, 3, payee)
        ws.write(r, 4, amt, cur)
        ws.write_datetime(r, 5, _dt(d), dt)
    ws.write("C22", "Total", bold)
    ws.write_formula("E22", "=SUM(E19:E21)", cur_bold, 12500.75)

    ws.write("B24", "* Reserve figures as of the valuation date. Source: claims system export (fictional data).", ital)
    ws.insert_textbox("K11", "Reviewer note: CL-10447 was reopened after a supplemental cargo invoice.", {
        "width": 280, "height": 70, "font": {"size": 11},
    })
    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "Paid", "categories": "='Loss Run'!$B$6:$B$13", "values": "='Loss Run'!$G$6:$G$13"})
    ch.set_title({"name": "Paid by claim (USD)"})
    ws.insert_chart("K17", ch)
    ws.set_column("B:B", 12)
    ws.set_column("C:C", 22)
    wb.close()
    made[path.name] = path

    # ---- claims_workbook.xlsx: several sheets incl. a hidden one --------------------------------
    path = out / "claims_workbook.xlsx"
    wb = xlsxwriter.Workbook(path)
    cur = wb.add_format({"num_format": "$#,##0.00"})
    dt = wb.add_format({"num_format": "yyyy-mm-dd"})
    pct = wb.add_format({"num_format": "0.0%"})
    bold = wb.add_format({"bold": True})

    s1 = wb.add_worksheet("Claims")
    for c, h in enumerate(("Claim ID", "Policy", "Insured", "Date of loss", "Cause", "Status", "Paid", "Reserve", "Loss ratio")):
        s1.write(0, c, h, bold)
    for r, c in enumerate(CLAIMS, start=1):
        s1.write(r, 0, c[0]); s1.write(r, 1, c[1]); s1.write(r, 2, c[2])
        s1.write_datetime(r, 3, _dt(c[3]), dt); s1.write(r, 4, c[4]); s1.write(r, 5, c[5])
        s1.write(r, 6, c[6], cur); s1.write(r, 7, c[7], cur); s1.write(r, 8, c[8], pct)
    s1.write(9, 5, "Total", bold)
    s1.write_formula(9, 6, "=SUM(G2:G9)", cur, TOTAL_PAID)

    s2 = wb.add_worksheet("Payments")
    for c, h in enumerate(("Claim ID", "Date", "Type", "Payee", "Amount")):
        s2.write(0, c, h, bold)
    pays = [("CL-10421", p[3], p[0], p[1], p[2]) for p in PAYMENTS_10421] + [
        ("CL-10422", date(2026, 3, 1), "Indemnity", "Alder Logistics, LLC", 4200.00),
        ("CL-10435", date(2026, 3, 15), "Indemnity", "Brightwater Dental", 30000.00),
        ("CL-10435", date(2026, 4, 10), "Expense", "Fire investigator", 6750.40),
        ("CL-10440", date(2026, 3, 20), "Indemnity", "Cobalt Freight", 9875.10),
        ("CL-10447", date(2026, 3, 27), "Indemnity", "Cobalt Freight", 14000.00),
        ("CL-10447", date(2026, 3, 28), "Expense", "Surveyor", 1300.00),
        ("CL-10463", date(2026, 4, 28), "Indemnity", "Elm Street Clinic", 2310.25),
    ]
    for r, (cid, d, typ, payee, amt) in enumerate(pays, start=1):
        s2.write(r, 0, cid); s2.write_datetime(r, 1, _dt(d), dt); s2.write(r, 2, typ); s2.write(r, 3, payee); s2.write(r, 4, amt, cur)
    s2.write(len(pays) + 2, 3, "Total", bold)
    s2.write_formula(len(pays) + 2, 4, f"=SUM(E2:E{len(pays) + 1})", cur, round(sum(p[4] for p in pays), 2))

    s3 = wb.add_worksheet("Reserves")
    s3.merge_range("A1:D1", "Reserve development by month", bold)
    for c, h in enumerate(("Month", "Opened", "Closed", "Net reserve")):
        s3.write(1, c, h, bold)
    for r, row in enumerate([("2026-01", 1, 0, 0.0), ("2026-02", 2, 1, 69500.0), ("2026-03", 3, 1, 81450.5), ("2026-04", 2, 0, 92950.5)], start=2):
        for c, v in enumerate(row):
            s3.write(r, c, v, cur if c == 3 else None)

    s4 = wb.add_worksheet("Internal Notes")
    s4.write("A1", "Reserve rationale (internal - not for distribution)", bold)
    s4.write("A2", "CL-10435: hold reserve at $61,000 pending equipment appraisal.")
    s4.hide()
    wb.close()
    made[path.name] = path

    # ---- loss_run_misnamed.xls: OOXML content with a legacy extension -----------------------------
    mis = out / "loss_run_misnamed.xls"
    shutil.copyfile(out / "loss_run.xlsx", mis)
    made[mis.name] = mis

    # ---- legacy_register.xls: a real BIFF workbook ----------------------------------------------
    try:
        import xlwt

        book = xlwt.Workbook()
        sh = book.add_sheet("Register")
        head = xlwt.easyxf("font: bold on")
        money_fmt = xlwt.easyxf(num_format_str="#,##0.00")
        date_fmt = xlwt.easyxf(num_format_str="YYYY-MM-DD")
        sh.write_merge(0, 0, 0, 5, "Legacy claims register (1997-style export)", xlwt.easyxf("font: bold on, height 280"))
        for c, h in enumerate(("Claim ID", "Insured", "Date of loss", "Status", "Paid", "Reserve")):
            sh.write(2, c, h, head)
        for r, c in enumerate(CLAIMS[:6], start=3):
            sh.write(r, 0, c[0]); sh.write(r, 1, c[2]); sh.write(r, 2, c[3], date_fmt)
            sh.write(r, 3, c[5]); sh.write(r, 4, c[6], money_fmt); sh.write(r, 5, c[7], money_fmt)
        legacy = out / "legacy_register.xls"
        book.save(str(legacy))
        made[legacy.name] = legacy
    except Exception as e:  # pragma: no cover
        print("skipped legacy .xls:", e)


# --------------------------------------------------------------------------- PowerPoint
def _pptx(out: Path, made: Dict[str, Path]) -> None:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches

    prs = Presentation()
    s1 = prs.slides.add_slide(prs.slide_layouts[0])
    s1.shapes.title.text = "Q1 2026 Claims Review"
    s1.placeholders[1].text = "Harbor Mutual Insurance (fictional) - prepared 2026-05-04"

    s2 = prs.slides.add_slide(prs.slide_layouts[5])
    s2.shapes.title.text = "Claims by status"
    table = s2.shapes.add_table(4, 3, Inches(0.5), Inches(1.6), Inches(4.5), Inches(1.6)).table
    for r, row in enumerate([("Status", "Count", "Paid (USD)"), ("Open", "4", "43,260.65"), ("Closed", "2", "22,375.85"), ("Denied", "1", "0.00")]):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    cd = CategoryChartData()
    cd.categories = ["Open", "Closed", "Denied"]
    cd.add_series("Paid", (43260.65, 22375.85, 0))
    s2.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(5.2), Inches(1.6), Inches(4.3), Inches(3.2), cd)
    s2.shapes.add_textbox(Inches(0.5), Inches(4.0), Inches(4.5), Inches(0.8)).text_frame.text = "Loss ratio trending above 90% on two accounts."
    s2.notes_slide.notes_text_frame.text = "Speaker notes: mention CL-10447 reopening."

    s3 = prs.slides.add_slide(prs.slide_layouts[1])
    s3.shapes.title.text = "Actions"
    tf = s3.placeholders[1].text_frame
    tf.text = "Complete CL-10435 equipment appraisal"
    for line in ("Re-reserve CL-10447 after supplemental invoice", "Review Cobalt Freight cargo wording"):
        tf.add_paragraph().text = line
    grp = s3.shapes.add_group_shape()                                                   # grouped shapes
    grp.shapes.add_textbox(Inches(0.5), Inches(4.8), Inches(4), Inches(0.6)).text_frame.text = "Key risk: reserve adequacy"
    grp.shapes.add_textbox(Inches(5), Inches(4.8), Inches(4), Inches(0.6)).text_frame.text = "Mitigation: monthly reserve review"

    p = out / "claims_review.pptx"
    prs.save(p)
    made[p.name] = p


# --------------------------------------------------------------------------- PDF
def _pdf(out: Path, made: Dict[str, Path]) -> None:
    import pymupdf

    HELV, BOLD, ITAL = "helv", "hebo", "heit"
    doc = pymupdf.open()                       # A4: 595 x 842 pt
    space = lambda size: pymupdf.get_text_length(" ", fontname=HELV, fontsize=size)

    def put(page, x, y, segments, size=11):
        """Place consecutive styled segments on one line (so formatting changes become separate spans)."""
        if isinstance(segments, str):
            segments = [(segments, HELV)]
        for text, font in segments:
            page.insert_text((x, y), text, fontsize=size, fontname=font)
            x += pymupdf.get_text_length(text, fontname=font, fontsize=size) + space(size)

    def footer(page, n):
        put(page, 72, 815, [(f"Confidential - Harbor Mutual (fictional) - page {n} of 3", ITAL)], size=8)

    def table(page, y, cols, rows, size=10):
        put_row = lambda yy, cells, font: [page.insert_text((x, yy), str(v), fontsize=size, fontname=font) for x, v in zip(cols, cells)]
        put_row(y, rows[0], BOLD)
        for i, row in enumerate(rows[1:], start=1):
            put_row(y + 15 * i, row, HELV)
        return y + 15 * len(rows)

    # ---- page 1 ----
    p1 = doc.new_page()
    put(p1, 72, 92, [("Q1 2026 Loss Summary", BOLD)], size=26)
    put(p1, 72, 114, [("Harbor Mutual Insurance (fictional) - valuation date 30 April 2026", ITAL)], size=12)
    put(p1, 72, 152, [("1. Executive summary", BOLD)], size=16)
    y = 176
    for segs in [
        [("Paid losses for the quarter total", HELV), (money(TOTAL_PAID), BOLD), ("across eight reported claims,", HELV)],
        [("with outstanding case reserves of", HELV), (money(TOTAL_RESERVE), BOLD), (". Two claims account for", HELV)],
        [("most of the movement: the Brightwater Dental fire (CL-10435) and the", HELV)],
        [("reopened Cobalt Freight cargo loss (CL-10447).", HELV)],
    ]:
        put(p1, 72, y, segs)
        y += 16
    put(p1, 72, y + 22, [("Claims at a glance", BOLD)], size=16)
    rows = [("Claim ID", "Insured", "Cause", "Status", "Paid (USD)")] + [
        (c[0], c[2].replace(",", ""), c[4], c[5], f"{c[6]:,.2f}") for c in CLAIMS
    ]
    y = table(p1, y + 46, (72, 140, 270, 390, 470), rows)
    put(p1, 72, y + 22, [("Note:", BOLD), ("amounts exclude unallocated expense and are shown before reinsurance.", ITAL)])
    footer(p1, 1)

    # ---- page 2 ----
    p2 = doc.new_page()
    put(p2, 72, 92, [("2. Large losses", BOLD)], size=16)
    put(p2, 72, 116, "Two claims exceeded the $10,000 large-loss threshold this quarter.")
    put(p2, 72, 152, [("2.1 CL-10435 - Fire (Brightwater Dental)", BOLD)], size=13)
    y = 174
    for line in ["A kitchen fire on 19 February 2026 damaged the treatment area and equipment.", "Key facts:"]:
        put(p2, 72, y, line)
        y += 16
    for line in ["- Paid to date: $36,750.40 (building repairs and emergency services)",
                 "- Reserve raised from $45,000 to $61,000 on 30 March 2026",
                 "- Fire marshal ruled the cause electrical; no subrogation target identified"]:
        put(p2, 96, y, line)
        y += 16
    put(p2, 72, y + 8, [("Equipment appraisal is pending and may increase the reserve further.", ITAL)])
    put(p2, 72, y + 48, [("2.2 CL-10447 - Cargo loss (Cobalt Freight)", BOLD)], size=13)
    y += 70
    for line in ["Closed on 28 March after payment of $15,300.00, then reopened on 15 April after a",
                 "supplemental invoice arrived. A reserve of $7,200.50 is held."]:
        put(p2, 72, y, line)
        y += 16
    put(p2, 72, y + 22, [("Reserve development", BOLD)], size=16)
    table(p2, y + 46, (72, 160, 250, 340),
          [("Month", "Opened", "Closed", "Net reserve (USD)"), ("2026-01", 1, 0, "0.00"), ("2026-02", 2, 1, "69,500.00"),
           ("2026-03", 3, 1, "81,450.50"), ("2026-04", 2, 0, "92,950.50")])
    footer(p2, 2)

    # ---- page 3 ----
    p3 = doc.new_page()
    put(p3, 72, 92, [("3. Recommendations", BOLD)], size=16)
    y = 120
    for n, (head, rest) in enumerate([
        ("Re-reserve", "CL-10447 once the supplemental invoice is validated."),
        ("Escalate", "CL-10435 to the property large-loss unit."),
        ("Review", "cargo wording for Cobalt Freight at renewal."),
    ], start=1):
        put(p3, 72, y, [(f"{n}.", HELV), (head, BOLD), (rest, HELV)])
        y += 18
    put(p3, 72, y + 24, [("4. Sign-off", BOLD)], size=16)
    put(p3, 72, y + 50, "Prepared by: Claims analytics (fictional)")
    put(p3, 72, y + 66, "Reviewed by: Head of Claims")
    footer(p3, 3)

    p = out / "loss_summary.pdf"
    doc.save(p)
    made[p.name] = p
    enc = out / "loss_summary_encrypted.pdf"
    doc.save(enc, encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw=PASSWORD, owner_pw="owner")
    made[enc.name] = enc
    doc.close()

    padded = out / "loss_summary_padded.pdf"          # bytes before the %PDF header (mail preamble)
    padded.write_bytes(b"MAIL PREAMBLE\r\n" + p.read_bytes())
    made[padded.name] = padded

    # scanned receipt: a page that is only an image - there is no text layer to extract
    scan = pymupdf.open()
    page = scan.new_page()
    page.insert_image(pymupdf.Rect(72, 72, 523, 360), stream=_png_bytes())
    sp = out / "scanned_receipt.pdf"
    scan.save(sp)
    scan.close()
    made[sp.name] = sp


def _encrypted_office(out: Path, made: Dict[str, Path]) -> None:
    try:
        import msoffcrypto

        for plain_name, enc_name in (("fnol_form.docx", "fnol_form_encrypted.docx"), ("loss_run.xlsx", "loss_run_encrypted.xlsx")):
            with open(out / plain_name, "rb") as fin, open(out / enc_name, "wb") as fout:
                msoffcrypto.OfficeFile(fin).encrypt(PASSWORD, fout)
            made[enc_name] = out / enc_name
    except Exception as e:  # pragma: no cover
        print("skipped encrypted office files:", e)


def _misc(out: Path, made: Dict[str, Path]) -> None:
    png = out / "damage_photo.png"
    png.write_bytes(_png_bytes())
    made[png.name] = png

    z = out / "claim_package.zip"                       # recognised by signature, but nothing unpacks it
    with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in ("fnol_form.docx", "claims_register.csv", "claim_notes.txt", "damage_photo.png"):
            zf.write(out / name, name)
    made[z.name] = z


def build_all(out_dir: str | Path | None = None) -> Dict[str, Path]:
    """Create every sample file; returns ``{filename: path}``."""
    out = Path(out_dir) if out_dir else HERE / "samples"
    out.mkdir(parents=True, exist_ok=True)
    made: Dict[str, Path] = {}
    _text_files(out, made)
    _docx(out, made)
    _excel(out, made)
    _pptx(out, made)
    _pdf(out, made)
    _encrypted_office(out, made)
    _misc(out, made)
    return made


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else None
    files = build_all(target)
    print(f"Wrote {len(files)} sample files:")
    for name in sorted(files):
        print("  ", name)
