import tkinter as tk
from os import path
from glob import glob  
import pandas as pd
import pdfplumber


pd.reset_option('display.float_format')
pd.options.display.float_format = '{:.2f}'.format
#window creating
window = tk.Tk() 
window.title('pdf_extract')
window.geometry('400x150')  # 这里的乘是小x

var_path = tk.StringVar()
entry_path = tk.Entry(window, textvariable=var_path, font=('Arial', 14))
entry_path.place(x=100,y=90)

def reporting():
    path = var_path.get()
    # path2 = var_path2.get()
    pdf =  pdfplumber.open(path + '/Oil_GC.pdf') 

    df_alltables = pd.DataFrame(columns=['Deal', 'Group\nCompany\nCode', 'Trade', 'Counterparty Short\nName',
    'Purch', 'Product', 'Trade Date', 'Trader\nUser ID',
    'Security\nInserted By', 'Trade\nEstimated BL\nDate', 'Sign\noff',
    'Load Port', 'Discharge Port', 'Intention', 'Quantity'])

    for page in pdf.pages:
        tables = page.find_tables()
        len(tables)

        for i in range(len(tables)):
            t_content = tables[i].extract(x_tolerance = 5)
            df_table = pd.DataFrame(t_content[1:],columns=t_content[0])
            df_alltables = pd.concat([df_alltables, df_table])

    df_alltables['Trade'].replace(regex=True,inplace=True,to_replace=r'\n', value=r'')

    df_alltables.columns = ['Deal', 'Group Company Code', 'Trade', 'Counterparty Short\nName',
    'Purch', 'Product', 'Trade Date', 'Trader User ID',
    'Security\nInserted By', 'Trade Estimated BL Date', 'Sign off',
    'Load Port', 'Discharge Port', 'Intention', 'Quantity']

    df_alltables.to_excel(path + '/test_extract.xlsx')

# setting button
btn_print = tk.Button(window, text='Extract', command=reporting)
btn_print.place(x=50, y=50)

    

window.mainloop()