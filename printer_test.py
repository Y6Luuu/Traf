import tkinter as tk
import tkinter.ttk as ttk
import win32api
import win32print
import pandas as pd

#window creating
window = tk.Tk() 
window.title('LC Docs Printer')
window.geometry('500x300')  # 这里的乘是小x

def get_available_printers():
    return [printer[2] for printer in win32print.EnumPrinters(4)]

def update_default_printer_label():
    default_printer = win32print.GetDefaultPrinter()
    default_printer_text = 'Selected printer: {}'.format(default_printer)
    default_printer_label.config(text=default_printer_text)


def printer_file(filename):
    win32api.ShellExecute (
    0,
    "print",
    filename,
    None,
    ".",
    0
    )
    print(filename+'----打印成功')

# path and printer
tk.Label(window, text='Folder path:', font=('Arial', 14)).place(x=10, y=190)
tk.Label(window, text='Printer:', font=('Arial', 14)).place(x=10, y=150)

#setting printer
default_printer_label = tk.Label(window, bg='#626a77', fg='white')
default_printer_label.place(x=10, y=12)
update_default_printer_label()
refresh_button = tk.Button(window, text='Refresh', command=update_default_printer_label)
refresh_button.place(x=400, y=10)

selected_printer = tk.StringVar()
printer_choice_menu = ttk.Combobox(window, textvariable=selected_printer, values=get_available_printers(), width=35, state='readonly')
printer_choice_menu.place(x=120, y=150)

def set_default_printer():
    win32print.SetDefaultPrinter(selected_printer.get())
    update_default_printer_label()

set_default_printer_button = tk.Button(window, text='Set', command=set_default_printer)
set_default_printer_button.place(x=385, y=145, width=50)

# path and printer info entry
# path
var_path = tk.StringVar()
entry_path = tk.Entry(window, textvariable=var_path, font=('Arial', 14))
entry_path.place(x=120,y=190)
    
def printing():
    path = var_path.get()


    df = pd.read_excel(path + '/doc_amount.xlsx', engine = "openpyxl", sheet_name= "Sheet1")
    invoice_no = df.iat[0,1]
    assay_no = df.iat[1,1]
    weight_no = df.iat[2,1]
    coo_no = df.iat[3,1]
    pl_no = df.iat[4,1]

    # data_invoice = input("please enter the invoice amount: ")
    # data_assay = input("please enter the assay amount: ")
    # data_weight = input("please enter the weight amount: ")
    # data_origin = input("please enter the origin amount: ")
    # data_pl = input("please enter the pl amount: ")

    x1=int(invoice_no)
    x2=int(assay_no)
    x3=int(weight_no)
    x4=int(coo_no)
    x5=int(pl_no)

    #index definition
    i1 = 1
    sum1 = 0

    i2 = 1
    sum2 = 0

    i3 = 1
    sum3 = 0

    i4 = 1
    sum4 = 0

    i5 = 1
    sum5 = 0
    
    while i1 <= x1:
        print('while loop')
        print ('sum = %d, i = %d' % (sum1, i1))
        printer_file(path + '\INVOICE.docx')
        sum1 = sum1 + i1
        i1 = i1 + 1

    print ('出while循环.')
    

    while i2 <= x2:
        print('while loop')
        print ('sum = %d, i = %d' % (sum2, i2))
        printer_file(path + '\ASSAY.docx')
        sum2 = sum2 + i2
        i2 = i2 + 1

    print ('出while循环.')
    

    while i3 <= x3:
        print('while loop')
        print ('sum = %d, i = %d' % (sum3, i3))
        printer_file(path + '\WEIGHT.docx')
        sum3 = sum3 + i3
        i3 = i3 + 1

    print ('出while循环.')
    

    while i4 <= x4:
        print('while loop')
        print ('sum = %d, i = %d' % (sum4, i4))
        printer_file(path + '\COO.docx')
        sum4 = sum4 + i4
        i4 = i4 + 1

    print ('出while循环.')
    

    while i5 <= x5:
        print('while loop')
        print ('sum = %d, i = %d' % (sum5, i5))
        printer_file(path + '\PL.docx')
        sum5 = sum5 + i5
        i5 = i5 + 1

    print ('出while循环.')
    l = tk.Label(window, bg='pink', width=20, height=3, text='finish', font=('Arial', 14)).place(x=145, y= 40)

    # setting button
btn_print = tk.Button(window, text='Print', command=printing)
btn_print.place(x=385, y=190)

window.mainloop()