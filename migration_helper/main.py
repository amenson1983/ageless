from tkinter import Tk, PhotoImage, Label, Frame, Menu, StringVar, END
import tkinter as tk
from tkinter.filedialog import askopenfilename

import pandas as pd


class CFunctions_for_app():
    def __init__(self,window):
        self._sheets = ''
        self._path = ''
        self.window = window

    @property
    def path(self):
        return self._path

    @property
    def sheets(self):
        return self._sheets

    @path.setter
    def path(self, new_path):
        self._path = new_path

    @sheets.setter
    def sheets(self, new_sheets):
        self._sheets = new_sheets



    def get_path(self):
        path = askopenfilename()
        self._path = path


    def get_sheets(self):
        file = pd.ExcelFile(self._path)
        self._sheets = file.sheet_names
        print(self._sheets)
        my_listbox = tk.Listbox(window)
        for item in self._sheets:
            my_listbox.insert(END, item)
        my_listbox.pack(pady=15)









window = tk.Tk()
f = CFunctions_for_app(window)

window.title("Migration Helper v 0.1")
window.geometry("600x500+500+150")
bg = PhotoImage(file="images/login_background.png")
background = Label(window, image=bg)
background.place(x=0, y=0)


open_file_button = tk.Button(window,
                             text="Open file",
                             bg="#B4D2F3",
                             fg="black",
                             command=f.get_path,
                             font='Times 15'
                             )

open_file_button.place(x=10, y=10)
print_columns_button = tk.Button(window,
                             text="Get sheets",
                             bg="#B4D2F3",
                             fg="black",
                             command=f.get_sheets,
                             font='Times 15'
                             )

print_columns_button.place(x=10, y=50)





window.mainloop()





