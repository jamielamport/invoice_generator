import pandas as pd
import glob
import datetime
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

current_date = datetime.datetime.now()
current_date_string = current_date.strftime("%Y.%m.%d")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    print(invoice_nr)
    print(date)

    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=100, h=8, txt=f"Invoice Nr.{invoice_nr}", border=0, ln=1)

    pdf.cell(w=100, h=8, txt=f"Date {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    grand_total = 0

    # Add headers
    headers = list(df.columns)
    headers = [item.replace("_", " ").title() for item in headers]
    pdf.set_font(family="Arial", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=str(headers[0]), border=1)
    pdf.cell(w=70, h=8, txt=str(headers[1]), border=1)
    pdf.cell(w=30, h=8, txt=str(headers[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(headers[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(headers[4]), border=1, ln=1)

    # Add rows
    for index, row in df.iterrows():
        pdf.set_font(family="Arial", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=row["product_name"], border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)
        grand_total += row['total_price']

    # aAdd total row
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="Total", border=1)
    pdf.cell(w=30, h=8, txt=f"£ {grand_total}", border=1, ln=1)

    # Add company name and logo
    pdf.set_font(family="Arial", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"pdf_invoices/{invoice_nr}_invoice.pdf")

