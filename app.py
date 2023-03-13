import tkinter as tk
from openpyxl import *
from datetime import *; from dateutil.relativedelta import *
import calendar
import os
from tkinter.ttk import Combobox

window = tk.Tk()

window.title("File Open")
window.geometry("600x450")

label1 = tk.Label(window, text = "Firma İsmi:")
label1.pack()

entry1=tk.Entry(window,bd=2)
entry1.pack()

label2 = tk.Label(window, text = "Proje Adı:")
label2.pack()

entry2=tk.Entry(window,bd=2)
entry2.pack()

label3 = tk.Label(window, text = "İş Emri No:")
label3.pack()

entry3=tk.Entry(window,bd=2)
entry3.pack()

label4 = tk.Label(window, text = "Mevcut Durumu veya Not:")
label4.pack()

entry4=tk.Entry(window,bd=2)
entry4.pack()

label5 = tk.Label(window, text = "Excel dosyanızın dizinini giriniz")
label5.pack()

entry5=tk.Entry(window,bd=2)
entry5.pack()

label6 = tk.Label(window, text = "Firmalar klasörünün dizinini giriniz")
label6.pack()

entry6=tk.Entry(window,bd=2)
entry6.pack()

variable_1=tk.StringVar()
variable_1.set("Çalışacak Kişiyi Seçin")
values_1=["Alper MUDURLU" ,"Bilal GELİNCİK"]
combobox_1=Combobox(master=window,textvariable=variable_1,values=values_1)
combobox_1.pack()


NOW = datetime.now()


def gonder():
    excel= load_workbook(r"C:\Users\t450\Desktop\5_dosya_acma\gelen_isler.xlsx") #Veri işlenecek excel dosya dizini
    sheet=excel.active
    sheet.append((entry1.get(),entry2.get(),entry3.get(),entry4.get(),combobox_1.get(),NOW))
    excel.save(r"C:\Users\t450\Desktop\5_dosya_acma\gelen_isler.xlsx")
    excel.close()

    os.makedirs(r"C:\Users\t450\Desktop\Firmalar" + "\\" + entry1.get() + "\\" + entry3.get() + "\\" + "3D")  #Firmalar klasörü dizini
    os.makedirs(r"C:\Users\t450\Desktop\Firmalar" + "\\" + entry1.get() + "\\" + entry3.get() + "\\" + "musteriden_gelen")
    os.makedirs(r"C:\Users\t450\Desktop\Firmalar" + "\\" + entry1.get() + "\\" + entry3.get() + "\\" + "pdf")
    os.makedirs(r"C:\Users\t450\Desktop\Firmalar" + "\\" + entry1.get() + "\\" + entry3.get() + "\\" + "setup")
    os.makedirs(r"C:\Users\t450\Desktop\Firmalar" + "\\" + entry1.get() + "\\" + entry3.get() + "\\" + "dıs_tedarik")
    os.startfile(r"C:\Users\t450\Desktop\Firmalar" + "\\" + entry1.get() + "\\" + entry3.get() + "\\" + "musteriden_gelen")

buton=tk.Button(window,text="Gönder",command=gonder)
buton.pack()

window.mainloop()