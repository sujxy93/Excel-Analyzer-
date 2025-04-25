from openpyxl import workbook, load_workbook
import matplotlib.pyplot as plt
from tkinter import *
from tkinter import ttk, messagebox

root = Tk()
root.title("Excel Data Analysis")
root.geometry("470x500+450+150")

ID = {}

def ex():
    #ID[int(id)] = [fn, tit, xl, yl, op, col, rg1, rg2]
    for key, item in ID.items():
        book = load_workbook(''+item[0]+'.xlsx')
        sheet = book.active
        data1 = []
        for x in range(int(item[6]), int(item[7])):
            data1.append(sheet[str(item[5]) + str(x)].value)
        c = set(data1)
        a = list(c)
        if len(a) > 5:
            b = len(a)
            print(b)
            for x in range(6, b + 1):
                print(x)
                a.remove(str(a[5]))

        x_axis = []
        y_axis = []
        for x in a:
            print(f"{x}:{data1.count(x)}")
            x_axis.append(x)
            y_axis.append(data1.count(x))
        print(x_axis)
        print(y_axis)
        plt.title(item[1])
        plt.xlabel(item[2])
        plt.ylabel(item[3])
        plt.bar(x_axis, y_axis)
        plt.show()


def exit():
    root.destroy()


def insert():
    xl = xl_ent.get()
    yl = yl_ent.get()
    op = op_ent.get()
    col = col_ent.get()
    fn = fn_ent.get()
    rg1 = rg1_ent.get()
    rg2 = rg2_ent.get()
    tit = tit_ent.get()
    id = id_ent.get()
    if len(xl) == 0 or len(yl) == 0 or len(op) == 0 or len(col) == 0 or len(fn) == 0 or len(rg1) == 0 or len(tit) == 0 or len(id) == 0 or len(rg2) == 0:
        messagebox.showinfo("INFO", "PLease Enter all the required fields")
    elif ID.get(int(id)) is not None:
        messagebox.showinfo("ERROR", "ID already exists!")
    else:
        treeview.insert(parent='', index='end', iid=str(id), values=(str(id), fn, tit))
        ID[int(id)] = [fn, tit, xl, yl, op, col, rg1, rg2]
        print(ID)
        xl_ent.delete(0, END)
        yl_ent.delete(0, END)
        col_ent.delete(0, END)
        fn_ent.delete(0, END)
        rg1_ent.delete(0, END)
        rg2_ent.delete(0, END)
        tit_ent.delete(0, END)
        id_ent.delete(0, END)


def delete():
    try:
        selected_item = treeview.selection()[0]
        Values = treeview.item(selected_item)['values'][0]

        ID.pop(Values)
        treeview.delete(selected_item)
        print(ID)
    except:
        pass

frame = ttk.LabelFrame(root)
frame.pack(pady=5, padx=5)
treeview = ttk.Treeview(frame)


# columns
treeview['columns'] = ("ID", "FileName", "Title")
treeview.column("#0", width=0, minwidth=0)
treeview.column("ID", width=50, anchor=W)
treeview.column("FileName", width=150, anchor=W)
treeview.column("Title", width=150, anchor=W)

# Headings
treeview.heading("#0", text="", anchor=W)
treeview.heading("ID", text="ID", anchor=W)
treeview.heading("FileName", text="FileName", anchor=W)
treeview.heading("Title", text="Title", anchor=W)

treeview.pack(pady=15, padx=15)

frame2 = ttk.LabelFrame(root)
frame2.pack(pady=10, padx=10)

# context
# lables
xl_lab = Label(frame2, text="Xlable:")
yl_lab = Label(frame2, text="Ylable:")
op_lab = Label(frame2, text="Options:")
col_lab = Label(frame2, text="Column:")
rg_lab = Label(frame2, text="Range:")
fn_lab = Label(frame2, text="Filename:")
tit_lab = Label(frame2, text="Title:")
id_lab = Label(frame2, text="ID:")

xl_lab.grid(row=0, column=0, sticky=W, pady=5, padx=5)
yl_lab.grid(row=0, column=2, sticky=W, pady=5, padx=5)
op_lab.grid(row=1, column=0, sticky=W, pady=5, padx=5)
col_lab.grid(row=1, column=2, sticky=W, pady=5, padx=5)
fn_lab.grid(row=2, column=0, sticky=W, pady=5, padx=5)
rg_lab.grid(row=2, column=2, sticky=W, pady=5, padx=5)
tit_lab.grid(row=3, column=0, sticky=W, pady=5, padx=5)
id_lab.grid(row=3, column=2, sticky=W, pady=5, padx=5)

# entry
xl_ent = ttk.Entry(frame2, width=15)
yl_ent = ttk.Entry(frame2, width=15)
op_ent = ttk.Combobox(frame2, width=12)
col_ent = ttk.Entry(frame2, width=15)
fn_ent = ttk.Entry(frame2, width=15)
rg1_ent = ttk.Entry(frame2, width=8)
rg2_ent = ttk.Entry(frame2, width=8)
tit_ent = ttk.Entry(frame2, width=15)
id_ent = ttk.Entry(frame2, width=15)

data = ["Count", "Average"]
op_ent["values"] = data
op_ent.current(0)

xl_ent.grid(row=0, column=1)
yl_ent.grid(row=0, column=3)
op_ent.grid(row=1, column=1)
col_ent.grid(row=1, column=3)
fn_ent.grid(row=2, column=1)
rg1_ent.grid(row=2, column=3, sticky=W)
rg2_ent.grid(row=2, column=3, sticky=E)
tit_ent.grid(row=3, column=1)
id_ent.grid(row=3, column=3)

# Buttons
btn_exit = ttk.Button(frame2, text="Exit", width=15, command=exit)
btn_add = ttk.Button(frame2, text="Add", width=15, command=insert)
btn_del = ttk.Button(frame2, text="Delete", width=15, command=delete)
btn_exe = ttk.Button(frame2, text="Execute", width=15, command=ex)

btn_exit.grid(row=4, column=0, padx=5, pady=5)
btn_del.grid(row=4, column=1, padx=5, pady=5)
btn_add.grid(row=4, column=2, padx=5, pady=5)
btn_exe.grid(row=4, column=3, padx=5, pady=5)

root.mainloop()