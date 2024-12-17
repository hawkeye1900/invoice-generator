import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)
print("\n")


for filepath in filepaths:

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

    # Read data from dataframe
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # get column headings
    columns = df.columns

    formatted_columns = [col.replace("_", " ").title() for col in columns]

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=formatted_columns[0], border=1,)
    pdf.cell(w=70, h=8, txt=formatted_columns[1], border=1,)
    pdf.cell(w=35, h=8, txt=formatted_columns[2], border=1,)
    pdf.cell(w=30, h=8, txt=f"{formatted_columns[3]} (£)", border=1,)
    pdf.cell(w=25, h=8, txt=f"{formatted_columns[4]} (£)", border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=25, h=8, txt=str(row["total_price"]), border=1, ln=1)

    total = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=25, h=8, txt=str(total), border=1, ln=1)

    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(255, 80, 80)
    pdf.cell(w=30, h=8, txt="")
    pdf.cell(w=70, h=8, txt="")
    pdf.cell(w=30, h=8, txt="")
    pdf.cell(w=35, h=8, align="C", txt=f"The total price is:")
    pdf.cell(w=25, h=8, txt=f"£{str(total)}")

    pdf.output(f"invoices/PDFs/{filename}.pdf")
