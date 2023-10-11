import tkinter as tk
from tkinter import ttk
from test import data
from tkinter import messagebox
import openpyxl
import os

File: set = {}
'''
data = {
#     "CSR001-1-1": ("Plastic Switch Cover", 90),
#     "CSR001-1-2": ("Metal Switch Cover", 45),
#     "CSR001-1-3": ("Decorative Switch Cover", 45),
#     "CSR001-2-1": ("Small Screws", 45),
#     "CSR001-2-2": ("Medium Screws", 45),
#     "CSR001-2-3": ("Large Screws", 45),
#     "CSR001-3-1": ("Copper Wire Terminal", 45),
#     "CSR001-3-2": ("Aluminum Wire Terminal", 45),
#     "CSR001-3-3": ("Brass Wire Terminal", 45),
#     "CSR002-1-1": ("E27 Bulb Holder", 45),
#     "CSR002-1-2": ("E14 Bulb Holder", 45),
#     "CSR002-2-1": ("5W LED Bulb", 100),
#     "CSR002-2-2": ("10W LED Bulb", 45),
#     "CSR002-2-3": ("15W LED Bulb", 45),
#     "CSR002-3-1": ("Cardboard Packaging", 45),
#     "CSR002-3-2": ("Plastic Packaging", 45),
#     "CSR003-1": ("Fan Blades", 80),
#     "CSR003-2": ("Mounting Brackets", 45),
#     "CSR003-3": ("Screws", 45),
#     "CSR004-1": ("Socket Cover", 45),
#     "CSR004-2": ("Socket Base", 45),
#     "CSR004-3": ("Screws", 45),
#     "CSR005-1": ("Breaker Switch", 45),
#     "CSR005-2": ("Mounting Rail", 45),
#     "CSR005-3": ("Wires", 45),
#     "CSR006-1": ("Copper Wire", 45),
#     "CSR006-2": ("Insulation", 45),
#     "CSR006-3": ("Packaging", 45),
#     "CSR007-1": ("Fixture Housing", 45),
#     "CSR007-3": ("Mounting Plate", 45),
#     "CSR008-1": ("Outlet Cover", 45),
#     "CSR008-2": ("Socket", 45),
#     "CSR008-3": ("Mounting Box", 45),
#     "CSR009-1": ("Panel Box", 45),
#     "CSR009-2": ("Circuit Breakers", 45),
#     "CSR009-3": ("Bus Bars", 45),
#     "CSR010-1": ("Plug", 45),
#     "CSR010-2": ("Socket", 45),
#     "CSR010-3": ("Cord Jacket", 45),
# }


# File = [("hall", 'CSR001-1-1', 45, 88), ("hall", "CSR002-2-1", 25, 36),
#         ("Kitchen", "CSR003-1", 55, 86), ("Kitchen", 'CSR002-2-1', 98, 63)]
'''

def showTree():

    for item in tree.get_children():
        tree.delete(item)

    for row in File:
        tree.insert("", 'end', values=row)


def addData():
    global File
    product_name: str = name_entry.get()
    product_qty: int = qty_entry.get()
    csr_entry: int | str = sel.get()
    if (product_name == "" or product_name == "Name") or (product_qty == "" or product_qty == "Qty") or (csr_entry == "" or csr_entry == "CSR no."):
        messagebox.showerror("App", "Please Enter Data Properly")
    else:
        csr_value = f"{data[csr_entry][0]}"
        rate = f"{data[csr_entry][2]}"
        row = (product_name, csr_value, product_qty, rate)
        File = set(File)
        File.add(row)
        print(File)
        showTree()


def DeleteItem():
    global File
    curItem = tree.focus()
    if curItem:
        item = tree.item(curItem)
        values = item['values']
        selected_tuple = (values[0], values[1], str(values[2]), str(values[3]))
        # print(selected_tuple)
        File.discard(selected_tuple)
        showTree()


def recSelected(event):
    global File
    curItem = tree.focus()
    if curItem:
        item = tree.item(curItem)
        values = item['values']
        selected_tuple = (values[0], values[1], str(values[2]), str(values[3]))
        name_entry.delete(0, tk.END)
        qty_entry.delete(0, tk.END)
        cb1.delete(0, tk.END)
        name_entry.insert(0, values[0])
        qty_entry.insert(0, values[2])
        for key, value in data.items():
            if value[0] == values[1]:
                csr = key
                cb1.insert(0, csr)


def EditItem():
    global File
    curItem = tree.focus()
    if curItem:
        item = tree.item(curItem)
        values = item['values']
        selected_tuple = (values[0], values[1], str(values[2]), str(values[3]))
        product_name: str = name_entry.get()
        product_qty: int = qty_entry.get()
        csr_entry: int | str = sel.get()
        csr_value = f"{data[csr_entry][0]}"
        rate = f"{data[csr_entry][1]}"
        new_tuple = (product_name, csr_value, product_qty, rate)
        File.discard(selected_tuple)
        File.add(new_tuple)
        print(selected_tuple, new_tuple)
        print(File)
        showTree()

    pass


def ExportItem():
    # os.remove('example.xlsx')
    global File,data   
    File = list(File)
    cols = ["E", 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O','P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    room = set()
    for i in File:
        room.add(i[0])
    room = list(room)

    # csr values
    csrValues = set()
    for i in File:
        csrValues.add(i[1])
    csrValues_list = list(csrValues)
    csrValues_list.sort()

    if 'example.xlsx' in os.listdir() :
        workbook = openpyxl.load_workbook('example.xlsx')
    else:
        workbook = openpyxl.Workbook()

    sheet = workbook.active
    sheet["A1"] = "Sr No"
    sheet["B1"] = "Description"
    sheet['C1'] = "CSR No"
    sheet['D1'] = "Rate"

    for i in range(len(csrValues_list)):
        sheet[f'A{i+2}'] = i + 1

    for i in range(len(csrValues)):
        sheet[f'B{i+2}'] = csrValues_list[i]
        print(csrValues_list[i])
        for key , value in data.items():
            if value[0] == csrValues_list[i]:
                csr = key

        sheet[f'C{i+2}'] = csr
        sheet[f'D{i+2}'] = f"{data.get(csr)[2]}/{data.get(csr)[1]}"

    for i in range(len(room)):
        sheet[f'{cols[i]}1'] = room[i]

    for data in File:
        for r in range(len(room)):
            if room[r] == data[0] :
                for i in range(len(csrValues_list)):
                    if csrValues_list[i] == data[1]:
                        sheet[f'{cols[r]}{csrValues_list.index(csrValues_list[i])+2}'] = data[2]
    
    workbook.save("example.xlsx")


win = tk.Tk()

style = ttk.Style(win)
win.tk.call("source", "forest-light.tcl")
win.tk.call("source", "forest-dark.tcl")
style.theme_use('forest-dark')

# Form


Form_Frame = ttk.LabelFrame(win, text="Execl Form")
Form_Frame.grid(row=0, column=0, padx=5, pady=10)

name_entry = ttk.Entry(Form_Frame,)
name_entry.insert(0, "Name")
name_entry.bind('<FocusIn>', lambda e: name_entry.delete(0, tk.END))
name_entry.grid(row=0, column=0, sticky="ew", padx=10, pady=5, columnspan=2)

qty_entry = ttk.Spinbox(Form_Frame, from_=1, to_=200)
qty_entry.insert(0, "Qty")
qty_entry.bind('<FocusIn>', lambda e: qty_entry.delete(0, tk.END)),
qty_entry.grid(row=1, column=0, sticky="ew", padx=10, pady=5, columnspan=2)


sel = tk.StringVar()  # string variable
cb1 = ttk.Combobox(Form_Frame, textvariable=sel)
cb1.insert(0, "CSR no.")
cb1['values'] = list(data.keys())
cb1.bind('<ButtonPress>', lambda e: cb1.delete(0, tk.END)),
cb1.grid(row=2, column=0, padx=5, pady=20, sticky="ew", columnspan=2)

# Buttons

Add_btn = ttk.Button(Form_Frame, text="Add", command=addData)
Add_btn.grid(row=3, column=0, pady=5)
Add_btn = ttk.Button(Form_Frame, text="Remove", command=DeleteItem)
Add_btn.grid(row=3, column=1, pady=5)

EditButton = ttk.Button(Form_Frame, text="Edit", command=EditItem)
EditButton.grid(row=4, column=0)

ExportBtn = ttk.Button(Form_Frame, text="Export", command=ExportItem)
ExportBtn.grid(row=4, column=1)

# Tree
treeFrame = ttk.Frame(win)
treeFrame.grid(row=0, column=1, pady=15, padx=5)

cols = ("Col A", "Col B", "Col C", "Col D")
tree = ttk.Treeview(treeFrame, show="headings", columns=cols, height=10)
tree.column(cols[1], anchor=tk.CENTER)
tree.column(cols[2], anchor=tk.CENTER)
tree.column(cols[3], anchor=tk.CENTER)
tree.heading(0, text="Name")
tree.heading(1, text="CSR")
tree.heading(2, text="Qty")
tree.heading(3, text="Rate")
tree.pack()
tree.bind('<<TreeviewSelect>>', recSelected)

showTree()
win.mainloop()
