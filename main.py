import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    file_name = Path(filepath).stem
    invoice_nr,date = file_name.split('-')
    # invoice_nr = file_name.split('-')[0]
    # date = file_name.split('-')[1]

    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Invoices nr :{invoice_nr}',ln=1)
    pdf.cell(w=50, h=8, txt=f'Date : {date}', ln=1)
    pdf.ln(10)

    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    # total of a particular row
    total_sum = df['total_price'].sum()

    # Add header to the table
    columns = df.columns
    columns = [items.replace('_', ' ').title() for items in columns]
    pdf.set_font(family='Times', size=10, style='B')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt=str(columns[0]), border=1)
    pdf.cell(w=50, h=8, txt=str(columns[1]), border=1)
    pdf.cell(w=40, h=8, txt=str(columns[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[4]), border=1, ln=1)

    # Add rows to the table
    for index, rows in df.iterrows():
        pdf.set_font(family='Times',size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(rows['product_id']), border=1)
        pdf.cell(w=50, h=8, txt=str(rows['product_name']), border=1)
        pdf.cell(w=40, h=8, txt=str(rows['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(rows['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(rows['total_price']), border=1, ln=1)

    # Add total row
    pdf.set_font(family='Times', size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=50, h=8, txt='', border=1)
    pdf.cell(w=40, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt='', border=1, ln=1)

    pdf.ln(7)

    #Footer text
    pdf.set_font(family='Times', size=16, style='B')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=10, txt=f'The total due amount is {total_sum} Euros.',ln=1)
    pdf.cell(w=30, h=10, txt="Jeeva's File")
    pdf.image('pythonhow.png',w=10)


    pdf.output(f'PDFs/{file_name}.pdf')

