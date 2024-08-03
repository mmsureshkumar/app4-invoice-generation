import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice No.{invoice_nr}", align="L", ln=1)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Date:{date}", align="L", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header
    columns = df.columns  # it will return list of index objects
    col = [item.replace("_", " ").title() for item in columns]
    print(col)
    pdf.set_font(family="Times", style="B", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=col[0], border=1)
    pdf.cell(w=65, h=8, txt=col[1], border=1)
    pdf.cell(w=30, h=8, txt=col[2], border=1)
    pdf.cell(w=30, h=8, txt=col[3], border=1)
    pdf.cell(w=30, h=8, txt=col[4], border=1, ln=1)

    for index, row in df.iterrows():
        # Add a row
        pdf.set_font(family="Times", style="B", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=65, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
