import pandas as pd
import glob
from fpdf import FPDF

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    file_name = filepath.title().split('\\')[-1].replace(".Xlsx", ".pdf")
    invoice_nr = file_name.split('-')[0]
    invoice_date = file_name.split('-')[-1].replace(".pdf", "")
    df = pd.read_excel(filepath)

    pdf = FPDF("landscape")
    pdf.add_page()
    pdf.set_font(family="Times", style='B', size=16)

    pdf.cell(txt=f"Invoice nr. {invoice_nr}", new_x="LEFT", new_y="NEXT")
    pdf.cell(txt=f"Invoice date. {invoice_date}", new_x="LEFT", new_y="NEXT")

    pdf.ln(10)
    pdf.set_font(family="Times", size=12, style='B')
    max_str = {}
    for data_top in df.head():
        data = data_top.replace('_', ' ').title()
        try:
            max_str[data_top] = df[data_top].str.len().max() * 4
        except AttributeError:
            max_str[data_top] = 45
        pdf.cell(w=max_str[data_top], h=10, border=True, align="C", txt=data, new_y="TOP")
    pdf.ln()

    total_price = df['total_price'].sum()
    df.loc['Total'] = pd.Series("TOTAL", index=['product_id'])

    for index, line in df.iterrows():
        for data_top in df.head():
            text = str(line[data_top]) if pd.notna(line[data_top]) else ""
            pdf.cell(w=max_str[data_top], h=10, border=True, align="C", txt=text,
                     new_y="TOP")
        pdf.ln()
    pdf.ln(20)

    final_message = f"Total price for this invoice is: {total_price} RON"
    pdf.cell(h=10, align="L", txt=final_message, new_y="TOP")
    pdf.output(fr"{file_name}")
