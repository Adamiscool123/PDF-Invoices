import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath)
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_number = filename.split("-")
    pdf.set_font(family="Times", size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Invoice nr. {invoice_number[0]}', ln=1)
    pdf.set_font(family="Times", size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Date {invoice_number[1]}', ln=1)

    columns = list(df.columns)
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=32, h=10, txt=columns[0], border=1)
    pdf.cell(w=65, h=10, txt=columns[1], border=1)
    pdf.cell(w=32, h=10, txt=columns[2], border=1)
    pdf.cell(w=32, h=10, txt=columns[3], border=1)
    pdf.cell(w=32, h=10, txt=columns[4], border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=32, h=10, txt=str(row["product_id"]).title(), border=1)
        pdf.cell(w=65, h=10, txt=str(row["product_name"]).title(), border=1)
        pdf.cell(w=32, h=10, txt=str(row["amount_purchased"]).title(), border=1)
        pdf.cell(w=32, h=10, txt=str(row["price_per_unit"]).title(), border=1)
        pdf.cell(w=32, h=10, txt=str(row["total_price"]).title(), border=1, ln=1)

    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80,80,80)
    pdf.cell(w=32, h=10, border=1)
    pdf.cell(w=65, h=10, border=1)
    pdf.cell(w=32, h=10, border=1)
    pdf.cell(w=32, h=10, border=1)
    pdf.cell(w=32, h=10, txt=str(total_sum) , border=1, ln=1)

    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f'The total due amount is {total_sum} Euros.', ln=1)

    pdf.set_font(family="Times", size=14, style='B')
    pdf.cell(w=25, h=8, txt="PythonHow")
    pdf.image('pythonhow.png', w=10)


    pdf.output(f"PDFS/{filename}.pdf")
