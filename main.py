import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)
print("\n")


for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # extracts the filename, minus extension and any path to file, such as
    # directories
    filename = Path(filepath).stem
    invoice_nr, invoice_date = filename.split("-")

    # Add invoice number as title
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", align="L", ln=1)

    # Adding date to invoice, extract date from filename and split
    invoice_date = invoice_date.split(".")
    invoice_date.reverse()

    # destructure to individual date elements, ie day, month, year
    day, month, year = invoice_date

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Date. {day}-{month}-{year}", align="L", ln=1)








    pdf.output(f"invoices/PDFs/{filename}.pdf")
