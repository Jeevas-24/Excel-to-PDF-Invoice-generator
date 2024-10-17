import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    file_name = Path(filepath).stem
    invoice_nr,date = file_name.split('-')
    # invoice_nr = file_name.split('-')[0]
    # date = file_name.split('-')[1]

    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Invoices nr : {invoice_nr}',ln=1)
    pdf.cell(w=50, h=8, txt=f'Date : {date}', ln=1)

    pdf.output(f'PDFs/{file_name}.pdf')
