import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")
print(filepaths)
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_no = filename.split("-")[0]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice no.{invoice_no}")
    pdf.output(f"PDFs/{filename}.pdf")
