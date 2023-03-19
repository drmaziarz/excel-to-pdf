import glob
from pathlib import Path

import pandas as pd
from fpdf import FPDF

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)

    pdf.set_font(family="Arial", size=12, style="")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    pdf.set_font(family="Arial", size=10, style="B")

    # Add a header
    widths = [30, 70, 35, 30, 30]
    columns = df.columns
    columns = [column.replace("_", " ").title() for column in columns]
    for column, width in zip(columns, widths):
        if columns.index(column) == len(columns) - 1:
            pdf.cell(w=width, h=8, txt=str(column), border=1, ln=1)
        else:
            pdf.cell(w=width, h=8, txt=str(column), border=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Arial", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    sum_price = df["total_price"].sum()

    pdf.set_font(family="Arial", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(sum_price), border=1, ln=1)

    pdf.set_font(family="Arial", size=12, style="B")
    pdf.cell(w=30, h=12, txt=f"Total price is {sum_price} $", ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
