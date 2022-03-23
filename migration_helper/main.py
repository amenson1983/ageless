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
        self._information_label = ''
        self._show_df_button = tk.Button()
        self.window = window
    @property
    def show_df_button(self):
        return self._show_df_button

    @show_df_button.setter
    def show_df_button(self, showdfbutton):
        self._show_df_button = showdfbutton

    @property
    def information_label(self):
        return self._information_label

    @information_label.setter
    def information_label(self, informationlabel):
        self._information_label = informationlabel

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

    def onselect_sheet(self,evt):
        w = evt.widget
        i = int(w.curselection()[0])
        value = w.get(i)
        self._sheetactual = value

        return value


    def onselect_col(self,evt):
        w = evt.widget
        i = int(w.curselection()[0])
        value = w.get(i)

        return value

    def get_path(self):
        path = askopenfilename()
        self._path = path


    def get_sheets(self):
        file = pd.ExcelFile(self._path)
        self._sheets = file.sheet_names
        self._mylistbox = tk.Listbox(window)
        for item in self._sheets:
            self._mylistbox.insert(END, item)
        self._mylistbox.pack(pady=15)
        self._mylistbox.bind('<<ListboxSelect>>', self.onselect_sheet)



    def get_columns_actual(self):

        df = pd.read_excel(self._path,self._sheetactual)
        self._cols = df.columns
        self._mylistbox.delete(0, END)
        for item in self._cols:
            self._mylistbox.insert(END, item)
        self._mylistbox.pack(pady=15)
        self._mylistbox.bind('<<ListboxSelect>>', self.onselect_col)



    def erase_key_columns_selection(self):
        self._key_columns_selection = []
        self._information_label = tk.Label(background,text=self._key_columns_selection)
        self._information_label.place(x=120, y=450)

    def define_key_columns_selection(self):
        selection = self._mylistbox.curselection()
        if selection[0] not in self._key_columns_selection:
            self._key_columns_selection.append(self._cols[selection[0]])
        else: pass
        self._information_label = tk.Label(background,text=self._key_columns_selection)
        self._information_label.place(x=120, y=450)




    def destroy_listbox(self):
        self._mylistbox.destroy()
        self.erase_key_columns_selection()

    def change(self,df,event, row, col):
        # get value from Entry
        value = event.widget.get()
        # set value in dataframe
        df.iloc[row, col] = value
        print(df)

    def show_dataframe(self):
        print(self._sheetactual)
        df1 = pd.read_excel(self._path, self._sheetactual)

        df = pd.DataFrame()
        for col in self._key_columns_selection:
            df[col] = df1[col]

        rows, cols = df.shape
        window_dataframe = tk.Tk()
        window_dataframe.title(f"Sheet: {self._sheetactual}")
        window_dataframe.resizable(width=True, height=False)

        for r in range(rows):
            for c in range(cols):
                e = tk.Entry(window_dataframe,width=40)
                e.insert(0, df.iloc[r, c])
                e.grid(row=r, column=c)
                # ENTER
                e.bind('<Return>', lambda event, y=r, x=c: f.change(df,event, y, x))
                # ENTER on keypad
                e.bind('<KP_Enter>', lambda event, y=r, x=c: f.change(df,event, y, x))


        window_dataframe.mainloop()



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

show_df_button = tk.Button(window,text="Preview dataframe",bg="#B4D2F3",fg="black",command=f.show_dataframe,font='Times 15')
show_df_button.place(x=10, y=170)

erase_listbox_button = tk.Button(window,text="Destroy ListBox",bg="#B4D2F3",fg="black",command=f.destroy_listbox,font='Times 15')
erase_listbox_button.place(x=440, y=220)

window.mainloop()





