import pandas as pd
import glob
import datetime
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

current_date = datetime.datetime.now()
current_date_string = current_date.strftime("%Y.%m.%d")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    print(invoice_nr)
    print(date)

    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=100, h=8, txt=f"Invoice Nr.{invoice_nr}", border=0, ln=1)

    pdf.cell(w=100, h=8, txt=f"Date {date}")

    pdf.output(f"pdf_invoices/{invoice_nr}_invoice.pdf")

