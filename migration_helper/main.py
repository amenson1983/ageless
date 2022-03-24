import os
from tkinter import Tk, PhotoImage, Label, Frame, Menu, StringVar, END, VERTICAL, NS, RIGHT, Y
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter.ttk import Scrollbar

import pandas as pd
from fuzzywuzzy import process
from openpyxl import load_workbook


class CFunctions:

    def item_match_in_list_by_percent(self,item,list_values,percent):
        s = process.extractOne(item, list_values)
        matched_value = ''
        if s[1] >= percent:
            matched_value = s[0]
        return matched_value

    def intermediate_changed_list(self,string,unnecessary_symbols_list):
        mapping_item_dictionary = {}
        changed_string = []
        for symb in unnecessary_symbols_list:
            changed_string_ = string.lower()
            changed_string = changed_string_.split(symb)
            changed_string = ''.join(changed_string)
            changed_string = changed_string.translate({ord(symb):None})
        for symb in unnecessary_symbols_list:
            changed_string = changed_string.translate({ord(symb): None})

        mapping_item_dictionary.update({changed_string:string})
        return changed_string,mapping_item_dictionary

    def soft_add_sheet_to_existing_xlsx(self,full_path,df,sheet_name):
        sheet = df
        book = load_workbook(full_path)
        writer = pd.ExcelWriter(full_path, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        sheet.to_excel(writer, sheet_name, index=False)
        writer.save()

    def melt_df(self,df, hold_columns_list, feature_rename_to, value_rename_to):
        df = df.melt(hold_columns_list).rename(columns={'variable':feature_rename_to,'value':value_rename_to})
        print(df.columns)
        print(df)
        return df

    def loc_df_by_column_equals_to(self,df,column_to_compare,value):
        self.df = df.loc[df[column_to_compare] == value ]
        return self.df

    def key_field_two_columns_insertion_to_dataframe(self,df,columns_list,key_field_name):
        df[key_field_name] = df[columns_list[0]].map(str) + df[columns_list[1]].map(str)
        return df

    def key_field_three_columns_insertion_to_dataframe(self,df,columns_list,key_field_name):
        df[key_field_name] = df[columns_list[0]].map(str) + df[columns_list[1]].map(str) + df[columns_list[2]].map(str)
        return df

    def key_field_four_columns_insertion_to_dataframe(self, df, columns_list, key_field_name):
        df[key_field_name] = df[columns_list[0]].map(str) + df[columns_list[1]].map(str) + df[columns_list[2]].map(str) + df[columns_list[3]].map(str)
        return df

    def map_dataframe_column_via_dictionary_and_get_new_df(self,df,target_column,new_column_name,dictionary):
        self.df = df
        self.df[new_column_name] = self.df[f'{target_column}'].apply(lambda x: pd.Series(x).map(dictionary))
        return self.df

    def map_data_to_first_df_from_second_by_key(self, df1, df2, key_field, columns_to_map):
        for i in range(0,len(columns_to_map)):
            data_dict = dict(zip(df2[key_field].values,df2[columns_to_map[i]].values))
            df1 = self.map_dataframe_column_via_dictionary_and_get_new_df(df1,key_field,f"{columns_to_map[i]}",data_dict)
        return df1,df2

    def vlookup_column(self,df_current,df_source,key_field,columns_to_migrate,temp_file):
        for column in columns_to_migrate:
            transfer_keys = df_source[key_field].values
            transfer_values = df_source[column].values
            dictionary_transfer = dict(zip(transfer_keys,transfer_values))
            df_current = self.map_dataframe_column_via_dictionary_and_get_new_df(df_current,key_field,f"{column}_from_source",dictionary_transfer)
            df_source.to_excel(temp_file,engine='openpyxl',sheet_name='source_data',index=False)
            df_current[f"{column}_from_source"] = df_current[f"{column}_from_source"].fillna('not_found')
            f.soft_add_sheet_to_existing_xlsx(temp_file,df_current,'vlookup_result')
            os.startfile(temp_file)
        return df_current,temp_file

    def sumif_column(self,df_current,df_source,key_field,columns_to_sum,temp_file):

        df_current_list = df_current[key_field].values
        for column in columns_to_sum:
            list_temp = []
            for cur in df_current_list:
                sum_ = df_source[column].loc[df_source[key_field] == cur].sum()
                list_temp.append(sum_)
            df_current[f"{column}_sum_from_source"] = list_temp
        df_current.to_excel(temp_file, engine='openpyxl', sheet_name='sum_if_result', index=False)
        f.soft_add_sheet_to_existing_xlsx(temp_file, df_source, 'source_data')
        os.startfile(temp_file)
        return df_current,temp_file






class CFunctions_for_app():
    def __init__(self,window):
        self._sheets = ''
        self._sheetactual = ''
        self._path = ''
        self._mylistbox = ''
        self._mylistbox_two = ''
        self._slave_columns_selection = []
        self._slave_column_to_change = ''
        self._cols = []
        self._information_label = ''
        self._show_df_button = tk.Button()
        self._df_income_selected = pd.DataFrame()
        self.working_file = 'working_file.xlsx'
        self.raw_selected_sheet_name = 'raw_selected'
        self.ethalon_selected_sheet_name = 'ethalon_selected'
        self.window = window

    @property
    def df_income_selected(self):
        return self._df_income_selected

    @df_income_selected.setter
    def df_income_selected(self, dfincomeselected):
        self._df_income_selected = dfincomeselected

    @property
    def mylistbox_two(self):
        return self._mylistbox

    @mylistbox_two.setter
    def mylistbox_two(self, mylistboxtwo):
        self._mylistbox_two = mylistboxtwo

    @property
    def mylistbox(self):
        return self._mylistbox

    @mylistbox.setter
    def mylistbox(self, my_listbox):
        self._mylistbox = my_listbox

    @property
    def slave_column_to_change(self):
        return self._slave_column_to_change

    @slave_column_to_change.setter
    def slave_column_to_change(self, value):
        self._slave_column_to_change = value

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
    def slave_columns_selection(self):
        return self._slave_columns_selection

    @slave_columns_selection.setter
    def slave_columns_selection(self, slavecolumnsselection):
        self._slave_columns_selection = slavecolumnsselection

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
        print(w.curselection())
        i = int(w.curselection()[0])
        value = w.get(i)

        return value

    def onselect_col_to_change(self, evt):
        w = evt.widget
        print(w.curselection())
        i = int(w.curselection()[0])
        value = w.get(i)
        self._slave_column_to_change = value
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
        self._slave_columns_selection = []
        self._information_label = []
        self._information_label = tk.Label(background, text=self._slave_columns_selection)
        self._information_label.place(x=120, y=450)

    def define_key_columns_selection(self):
        selection = self._mylistbox.curselection()
        if selection[0] not in self._slave_columns_selection:
            self._slave_columns_selection.append(self._cols[selection[0]])
        else: pass
        self._information_label = tk.Label(background, text=self._slave_columns_selection)
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
        df1 = pd.read_excel(self._path, self._sheetactual)
        for col in self._slave_columns_selection:
            self._df_income_selected[col] = df1[col]
        rows, cols = self._df_income_selected.shape
        window_dataframe = tk.Tk()

        for r in range(rows):
            for c in range(cols):
                e = tk.Entry(window_dataframe,width=30)
                e.insert(0, self._df_income_selected.iloc[r, c])
                e.grid(row=r, column=c)
                # ENTER
                e.bind('<Return>', lambda event, y=r, x=c: ff.change(self._df_income_selected,event, y, x))
                # ENTER on keypad
                e.bind('<KP_Enter>', lambda event, y=r, x=c: ff.change(self._df_income_selected,event, y, x))

        window_dataframe.mainloop()


    def put_selected_income_data_to_temporary_xlsx(self):
        print(self._df_income_selected)
        self._df_income_selected.to_excel(self.working_file,sheet_name=self.raw_selected_sheet_name,index=False)
        #f.soft_add_sheet_to_existing_xlsx(self.working_file,self._df_income_selected,self.raw_selected_sheet_name)
        os.startfile(self.working_file)
        self._df_income_selected = pd.DataFrame()
        self._slave_columns_selection = []

    def add_selected_income_data_to_temporary_xlsx(self):
        print(self._df_income_selected)
        #self._df_income_selected.to_excel(self.working_file,sheet_name=self.raw_selected_sheet_name,index=False)
        f.soft_add_sheet_to_existing_xlsx(self.working_file,self._df_income_selected,self.raw_selected_sheet_name)
        os.startfile(self.working_file)
        self._df_income_selected = pd.DataFrame()
        self._slave_columns_selection = []

    def add_selected_ethalon_data_to_temporary_xlsx(self):
        print(self._df_income_selected)
        #self._df_income_selected.to_excel(self.working_file,sheet_name=self.raw_selected_sheet_name,index=False)
        f.soft_add_sheet_to_existing_xlsx(self.working_file,self._df_income_selected,self.ethalon_selected_sheet_name)
        os.startfile(self.working_file)
        self._df_income_selected = pd.DataFrame()
        self._slave_columns_selection = []

window = tk.Tk()
ff = CFunctions_for_app(window)
f = CFunctions()
window.title("Migration Helper v 0.1")
window.geometry("830x600+500+150")
bg = PhotoImage(file="C:\\ageless\\migration_helper\\images\\login_background.png")
background = Label(window, image=bg)
background.place(x=0, y=0)


open_file_button = tk.Button(window,text="Open file",bg="#B4D2F3",fg="black",command=ff.get_path,font='Times 13')
open_file_button.place(x=10, y=10)

define_sheets_button = tk.Button(window,text="Get sheets",bg="#B4D2F3",fg="black",command=ff.get_sheets,font='Times 13')
define_sheets_button.place(x=10, y=44)

define_columns_button = tk.Button(window,text="Choose sheet",bg="#B4D2F3",fg="black",command=ff.get_columns_actual,font='Times 13')
define_columns_button.place(x=10, y=78)

define_selected_columns_button = tk.Button(window,text="Choose columns",bg="#B4D2F3",fg="black",command=ff.define_key_columns_selection,font='Times 13')
define_selected_columns_button.place(x=10, y=112)

show_df_button = tk.Button(window,text="Preview dataframe",bg="#B4D2F3",fg="black",command=ff.show_dataframe,font='Times 13')
show_df_button.place(x=10, y=146)

to_xlsx_df_button = tk.Button(window,text="Selected raw dataframe to xlsx",bg="#B4D2F3",fg="black",command=ff.put_selected_income_data_to_temporary_xlsx,font='Times 13')
to_xlsx_df_button.place(x=10, y=180)

erase_listbox_button = tk.Button(window,text="Clear input data",bg="#B4D2F3",fg="black",command=ff.destroy_listbox,font='Times 13')
erase_listbox_button.place(x=10, y=214)

switch_to_ethalon_button = tk.Button(window,text="Selected ethalon dataframe to xlsx",bg="#FCA65E",fg="black",command=ff.add_selected_ethalon_data_to_temporary_xlsx,font='Times 13')
switch_to_ethalon_button.place(x=570, y=10)

window.mainloop()





