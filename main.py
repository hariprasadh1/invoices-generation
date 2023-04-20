import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*xlsx")

for filepath in filepaths:
    filename = Path(filepath).stem
    heading = f"Invoice nr.{filename.split('-')[0]}"
    date = f"Date {filename.split('-')[1]}"
    df = pd.read_excel(io=filepath, sheet_name='Sheet 1')
    pdf = FPDF(orientation="P", format="A4", unit="mm")
    pdf.set_auto_page_break(auto=False, margin=0)
    pdf.add_page()
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=0, h=12, txt=heading, ln=1, align='L')
    pdf.cell(w=0, h=12, txt=date, ln=1, align='L')

    for col in df.columns:
        col = ' '.join([word.capitalize() for word in col.split('_')])
        pdf.set_font(family='Times', size=10, style='B')
        if col == "Product Name":
            pdf.cell(70, 10, col, 1)
        else:
            pdf.cell(30, 10, col, 1)

    pdf.ln()
    for index, data in df.iterrows():
        pdf.set_font(family='Times', size=10, style='')
        pdf.cell(30, 10, str(data["product_id"]), 1)
        pdf.cell(70, 10, str(data["product_name"]), 1)
        pdf.cell(30, 10, str(data["amount_purchased"]), 1)
        pdf.cell(30, 10, str(data["price_per_unit"]), 1)
        pdf.cell(30, 10, str(data["total_price"]), 1)
        pdf.ln()

    for i in range(len(df.columns) - 1):
        if i == 1:
            pdf.cell(70, 10, "", 1)
        else:
            pdf.cell(30, 10, "", 1)

    total_price = sum(df["total_price"])

    pdf.cell(30, 10, str(total_price), 1)
    pdf.ln(20)

    pdf.set_font(family='Times', style='B', size=20)
    pdf.cell(0, 10, "Total price is "+str(total_price), 0)
    pdf.ln(20)

    pdf.set_font(family='Arial', style='I', size=22)
    pdf.cell(0, 20, "Analyzed by - Hariprasadh Thirumal", 0)
    pdf.ln(170)

    pdf.set_font(family='Arial', style='', size=12)
    pdf.cell(0, 20, "@Copyright TMAL Technology 2023", 0, align='C')
    pdf.output('reports/'+filepath[9:24]+'.pdf')


