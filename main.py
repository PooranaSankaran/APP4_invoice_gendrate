import pandas as pd
import glob #it is used to read multiple data which updated in pycharm as floder
from fpdf import FPDF
from pathlib import Path #P caps

filepaths = glob.glob('invoices/*.xlsx') # To load all the data from invoice floder


for filepath  in filepaths:
    df = pd.read_excel(filepath, sheet_name= 'Sheet 1')  # sheet 1 is the excel name when you open it
    pdf = FPDF(orientation = 'P', unit = 'mm',format='A4') # p = portate
    pdf.add_page() # add page is must while creating pdf
    filename =  Path(filepath).stem # filename is to extract 10001,10002 etc.. from the file name to show in the pdf
    #stem gives us the number 10001-2023.1.18 fom the file so, we used that method to extract it

    # we just need 10001 from 10001-2023.1.18
    invoice_nr = filename.split('-')[0]
    pdf.set_font(family = 'Times', size = 16 , style = 'B')# creating font in pdf
    pdf.cell(w = 50,h=0,txt = f'Invoice nr.{invoice_nr}',ln=1) # creating cell in pdf
    #ln represesnt if the cell is created and new cell income has to be in next line

    # we need to add date from the file name like we extacted 100001(10001-2023.1.18)
    date = invoice_nr = filename.split('-')[1]
    pdf.set_font(family = 'Times', size = 16 , style = 'B')
    pdf.cell(w = 50,h=8,txt = f'Date:{date}',ln=1)


    # adding datas from excel to make it as pdf
    # adding columns
    columns = list(df.columns)
    # removing underscore from columns
    columns = [item.replace('_',' ').title() for item in columns]
    pdf.set_font(family='Times', size=10, style = 'B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    for index,row in df.iterrows():
        pdf.set_font(family='Times',size=10)
        pdf.set_text_color(80,80,80)
        # we don't need ln because we need data in same line
        pdf.cell(w=30,h=8, txt=str(row['product_id']), border= 1)  # border will create each block after a cell enters
        pdf.cell(w=70, h=8, txt= str(row['product_name']), border= 1) #eg: 101 | ice_cream | ....etc.. | ('|') this are borders
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border= 1)
        pdf.cell(w=30, h=8, txt= str(row['price_per_unit']), border= 1)
        pdf.cell(w=30, h=8, txt= str(row['total_price']), border= 1,ln=1)
        # here in last cell we giving ln = 1 beacause after one product complete it show start next product in new line

    # adding total price for the each file
    total_sum = df['total_price'].sum()

    pdf.set_font(family='Times', size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt= "", border=1)
    pdf.cell(w=70, h=8, txt="",border=1)
    pdf.cell(w=30, h=8, txt= "", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt= str(total_sum) , border=1, ln=1)

    #add total sum sentences

    pdf.set_font(family='Times', size = 10, style='B')
    pdf.cell(w = 30,h=0,txt=f"The total price is :{total_sum}",ln=1)

    #adding company name and logo
    pdf.set_font(family='Times',size=14, style='B')
    pdf.cell(w=25, h =8 , txt=f"PythonHow")
    pdf.image("pythonhow.png",w=10)

    pdf.output(f"PDFs/{filename}.pdf") # we doesn't have the extention so we creted pdf as extention
    # output store in PDFs folder for each file gendrate..

