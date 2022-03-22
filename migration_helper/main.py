from tkinter import Tk, PhotoImage, Label, Frame, Menu, StringVar, END
import tkinter as tk
from tkinter.filedialog import askopenfilename

import pandas as pd


class CFunctions_for_app():
    def __init__(self,window):
        self._sheets = ''
        self._sheetactual = ''
        self._path = ''
        self._mylistbox = ''
        self._key_columns_selection = []
        self._cols = []
        self.window = window

    @property
    def cols(self):
        return self._cols

    @cols.setter
    def cols(self, columns):
        self._cols = columns

    @property
    def key_columns_selection(self):
        return self._mylistbox

    @key_columns_selection.setter
    def key_columns_selection(self, keycolumnsselection):
        self._key_columns_selection = keycolumnsselection

    @property
    def mylistbox(self):
        return self._mylistbox

    @mylistbox.setter
    def mylistbox(self, my_listbox):
        self._mylistbox = my_listbox

    @property
    def sheetactual(self):
        return self._sheetactual

    @sheetactual.setter
    def sheetactual(self, sheet_actual):
        self._sheetactual = sheet_actual

    @property
    def path(self):
        return self._path

    @path.setter
    def path(self, new_path):
        self._path = new_path

    @property
    def sheets(self):
        return self._sheets

    @sheets.setter
    def sheets(self, new_sheets):
        self._sheets = new_sheets

    def onselect(self,evt):
        w = evt.widget
        i = int(w.curselection()[0])
        value = w.get(i)
        self.sheetactual = value
        return value


    def get_path(self):
        path = askopenfilename()
        self._path = path


    def get_sheets(self):
        file = pd.ExcelFile(self._path)
        self._sheets = file.sheet_names
        print(self._sheets)
        self._mylistbox = tk.Listbox(window)
        for item in self._sheets:
            self._mylistbox.insert(END, item)
        self._mylistbox.pack(pady=15)
        self._mylistbox.bind('<<ListboxSelect>>', self.onselect)

    def get_columns_actual(self):
        df = pd.read_excel(self._path,self.sheetactual)
        self._cols = df.columns
        self._mylistbox.delete(0, END)

        for item in self._cols:
            self._mylistbox.insert(END, item)
        self._mylistbox.pack(pady=15)
        self._mylistbox.bind('<<ListboxSelect>>', self.onselect)

    def define_key_columns_selection(self):

        selection = self._mylistbox.curselection()
        self._key_columns_selection.append(self.cols[selection[0]])
        print(self._key_columns_selection)






window = tk.Tk()
f = CFunctions_for_app(window)

window.title("Migration Helper v 0.1")
window.geometry("600x500+500+150")
bg = PhotoImage(file="images/login_background.png")
background = Label(window, image=bg)
background.place(x=0, y=0)


open_file_button = tk.Button(window,text="Open file",bg="#B4D2F3",fg="black",command=f.get_path,font='Times 15')
open_file_button.place(x=10, y=10)

define_sheets_button = tk.Button(window,text="Get sheets",bg="#B4D2F3",fg="black",command=f.get_sheets,font='Times 15')
define_sheets_button.place(x=10, y=50)

define_columns_button = tk.Button(window,text="Choose sheet",bg="#B4D2F3",fg="black",command=f.get_columns_actual,font='Times 15')
define_columns_button.place(x=10, y=90)

define_selected_columns_button = tk.Button(window,text="Choose columns",bg="#B4D2F3",fg="black",command=f.define_key_columns_selection,font='Times 15')
define_selected_columns_button.place(x=10, y=130)




window.mainloop()





