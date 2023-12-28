#importing all modules
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Set image as background
    pdf.image("image/img.png", x=80, y=120, w=50, h=70)

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
    Date = filename.split("-")[0]


    pdf.set_font(family="Times", size=18, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln = 1)


    pdf.set_font(family="Times", size=18, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln = 1)


    df = pd.read_excel(filepath, sheet_name="Sheet 1")

#Adding Header
    columns = list(df.columns)
    columns= [item.replace("_", " ").title()for item in columns]
    pdf.set_font(family="Times", size=10, style="BI")
    pdf.set_text_color(r=80, g=80, b=80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

#Adding Row
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=18)
        pdf.set_text_color(r=100, g=100, b=100)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border= 1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border= 1, ln =1)

#Calculating The total Bill
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=18)
    pdf.set_text_color(r=100, g=100, b=100)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border= 1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border= 1, ln =1)

#The total Bill
    pdf.set_font(family="Times", size=14, style= "BI")
    pdf.cell(w=30, h=8, txt=f"The Total Price is ${total_sum}")

#Adding Logo to the Top Right of the Invoice
    pdf.image("image/img.png", x=190, y=2, w=20, h=20)

    # Outputing the files
    pdf.output(f"Outputs/{filename}.pdf")



    #print(df)