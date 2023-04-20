import pandas as pd
import glob
from fpdf import FPDF

filepaths = glob.glob("invoices/*xlsx")

for filepath in filepaths:
    data = pd.read_excel(io=filepath, sheet_name='Sheet 1')
    print(sum(data['total_price']))


