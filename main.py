import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")
print(filepaths)
for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_no = filename.split("-")[0]
    date_on_bill = filename.split("-")[1]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice no.{invoice_no}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date->{date_on_bill}",ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    coloumns = list(df.columns)
    coloumns = [item.replace("_", " ").title() for item in coloumns]
    pdf.set_font(family="Times", size=9, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=coloumns[0], border=1)
    pdf.cell(w=70, h=8, txt=coloumns[1], border=1)
    pdf.cell(w=30, h=8, txt=coloumns[2], border=1)
    pdf.cell(w=70, h=8, txt=coloumns[3], border=1)
    pdf.cell(w=70, h=8, txt=coloumns[4], border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10, style="I")
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["price_per_unit"]), border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
