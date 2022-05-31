import os
from tkinter import Tk, PhotoImage, Label, Frame, Menu, StringVar, END, VERTICAL, NS, RIGHT, Y
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter.ttk import Scrollbar
import re
import pandera
from pandera import Column, Check
import numpy as np
import pandas as pd
from fuzzywuzzy import process
from openpyxl import load_workbook


class CFunctions:

    def item_match_in_list_by_percent(self,item,list_values,percent):
        s = process.extractOne(item, list_values)
        matched_value = ''
        if s[1] >= percent:
            matched_value = s[0]
        return matched_value, s[1]

    def intermediate_changed_list(self,string,unnecessary_symbols_list):
        mapping_item_dictionary = {}
        changed_string = []
        for symb in unnecessary_symbols_list:
            try:
                changed_string_ = string.lower()
            except Exception:
                changed_string_ = string
            changed_string = str(changed_string_).split(symb)
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

    def vlookup_column(self,df_current,df_source,key_field,columns_to_migrate):
        for column in columns_to_migrate:
            transfer_keys = df_source[key_field].values
            transfer_values = df_source[column].values
            dictionary_transfer = dict(zip(transfer_keys,transfer_values))

            df_current = self.map_dataframe_column_via_dictionary_and_get_new_df(df_current,key_field,column,dictionary_transfer)

        return df_current

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

    def list_correction_to_ethalon_naming_list(self,incoming_list,ethalon_naming_list,percent):
        corrected_list, problematic_items = [], {}

        for item in incoming_list:
            value = f.item_match_in_list_by_percent(item, ethalon_naming_list, percent)
            if value == '':
                s = process.extractOne(item, ethalon_naming_list)
                problematic_items.update({item:s[1]})
            corrected_list.append(value[0])
        if problematic_items != {}:
            print(f"Minimal percent: {min(problematic_items.values())}%")
        if problematic_items != {}:
            print(f"COGI: {problematic_items}")
        return corrected_list,problematic_items




class CFunctions_for_app():
    def __init__(self,window,background):
        self.pch_value_to_loc = ''
        self.pch_col_index = 0
        self.pch_col_to_replace_symbols_entry = ''
        self.pch_col_to_replace_for_symbols_entry = ''
        self.pch_col_to_replace_symbols = ''

        self.background = background
        self.accuracy = tk.IntVar()
        self.created_ethalon_column = ''
        self.pch_export_sheetname = 'frame'
        self.technical_sheet = 'ethalon_processed_sheet'
        self._sheets = ''
        self._sheetactual = ''
        self._path = ''
        self._mylistbox = ''
        self._mylistbox_two = ''
        self.unnecessary_symbols_list = ["№", "#","_","-", "%", "/", "|", ",", ".", ".", ",", "!", " ", "*",
                                         "(",")"]
        self.unnecessary_symbols_replace_dict = {"ß":"ss",
                                                 "ö":"oe",
                                                 "ü":"u","ē":"e",
                                                 "ä":"a",
                                                 "ā":"a",
                                                 "ī":"i",
                                                 "ū":"u"}
        self.full_words_to_replace = {"ГРИНДЕКС":""}
        self._slave_columns_selection = []
        self._slave_column_to_change = []
        self._remove_unnesc_symb_list = []
        self._cols = []
        self._information_label = tk.Label(window)
        self._information_label_path = tk.Label(window)
        self._information_label_sheet = tk.Label(window)
        self._information_label_columns = tk.Label(window)
        self._information_label_confirm_dataframe = tk.Label(window)
        self._information_label_export_raw_status = tk.Label(window)
        self._information_label_export_ethalon_status = tk.Label(window)
        self._information_label_working_path = tk.Label(window)
        self._information_label_working_sheet = tk.Label(window)
        self._information_label_working_columns_to_change = tk.Label(window)
        self._information_label_working_column_to_change = tk.Label(window)
        self._information_label_working_column_ethalon = tk.Label(window)
        self._information_label_working_column_result = tk.Label(window)
        self._show_df_button = tk.Button()
        self._df_income_selected = pd.DataFrame()
        self._df_active = pd.DataFrame()
        self.confirmed_col_to_change = ''
        self.confirmed_col_ethalon = ''
        self.key_field = tk.StringVar(window)
        self.key_field_label = ''
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
        self._information_label_path= tk.Label(window, text=self._path,background='white')
        self._information_label_path.place(x=100, y=15)
        return self._path

    def get_sheets(self):
        try:
            file = pd.ExcelFile(self._path)
            self._sheets = file.sheet_names
            self._mylistbox = tk.Listbox(window)
            for item in self._sheets:
                self._mylistbox.insert(END, item)
            self._mylistbox.place(x=600,y=500)
            self._mylistbox.bind('<<ListboxSelect>>', self.onselect_sheet)
        except IndexError:
            pass

    def get_columns_actual(self):
        try:
            df = pd.read_excel(self._path,self._sheetactual)

            self._information_label_sheet = tk.Label(window, text=self._sheetactual, background='white')
            self._information_label_sheet.place(x=120, y=85)
            self._cols = df.columns
            self._mylistbox.delete(0, END)
            for item in self._cols:
                self._mylistbox.insert(END, item)
            self._mylistbox.place(x=600,y=500)
            self._mylistbox.bind('<<ListboxSelect>>', self.onselect_col)
        except IndexError:
            pass



    def define_key_columns_selection(self):
        selection = self._mylistbox.curselection()
        if selection[0] not in self._slave_columns_selection:
            self._slave_columns_selection.append(self._cols[selection[0]])
        else: pass
        self._slave_columns_selection = pd.Series(self._slave_columns_selection).unique().tolist()
        self._information_label_columns.destroy()
        self._information_label_columns = tk.Label(window, text=self._slave_columns_selection, background='white')
        self._information_label_columns.place(x=150, y=120)


    def destroy_listbox(self):
        self._slave_columns_selection = []
        self._mylistbox.destroy()
        self._information_label.destroy()
        self._information_label_path.destroy()
        self._information_label_sheet.destroy()
        self._information_label_columns.destroy()
        self._information_label_export_raw_status.destroy()
        self._information_label_confirm_dataframe.destroy()



    def destroy_list_col_to_change(self):
        self._slave_column_to_change = []
        self._mylistbox.destroy()
        self._information_label_working_path.destroy()
        self._information_label_working_sheet.destroy()
        self._information_label_working_column_to_change.destroy()
        self._information_label_working_column_ethalon.destroy()
        self._information_label_working_column_result.destroy()
        self._information_label_export_ethalon_status.destroy()
        self._information_label_working_columns_to_change.destroy()


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
        self._information_label_confirm_dataframe = tk.Label(window, text="Successfully confirmed", background='white')
        self._information_label_confirm_dataframe.place(x=170, y=150)
        window_dataframe.mainloop()


    def put_selected_income_data_to_temporary_xlsx(self):
        df1 = pd.read_excel(self._path, self._sheetactual)
        for col in self._slave_columns_selection:
            self._df_income_selected[col] = df1[col]

        self._df_income_selected.to_excel(self.working_file,sheet_name=self.raw_selected_sheet_name,index=False)
        #f.soft_add_sheet_to_existing_xlsx(self.working_file,self._df_income_selected,self.raw_selected_sheet_name)
        os.startfile(self.working_file)
        self._df_income_selected = pd.DataFrame()
        self._slave_columns_selection = []
        self._information_label_export_raw_status = tk.Label(window, text="Export is successfully finished", background='white')
        self._information_label_export_raw_status.place(x=260, y=220)

    def add_selected_income_data_to_temporary_xlsx(self):
        df1 = pd.read_excel(self._path, self._sheetactual)
        for col in self._slave_columns_selection:
            self._df_income_selected[col] = df1[col]
        print(self._df_income_selected)
        self._df_income_selected.to_excel(self.working_file,sheet_name=self.raw_selected_sheet_name,index=False)
        #f.soft_add_sheet_to_existing_xlsx(self.working_file,self._df_income_selected,self.raw_selected_sheet_name)
        os.startfile(self.working_file)
        self._df_income_selected = pd.DataFrame()
        self._slave_columns_selection = []

    def add_selected_ethalon_data_to_temporary_xlsx(self):
        df1 = pd.read_excel(self._path, self._sheetactual)
        for col in self._slave_columns_selection:
            self._df_income_selected[col] = df1[col]
        print(self._df_income_selected)
        #self._df_income_selected.to_excel(self.working_file,sheet_name=self.raw_selected_sheet_name,index=False)
        f.soft_add_sheet_to_existing_xlsx(self.working_file,self._df_income_selected,self.ethalon_selected_sheet_name)
        os.startfile(self.working_file)
        self._df_income_selected = pd.DataFrame()
        self._slave_columns_selection = []
        self._information_label_export_ethalon_status = tk.Label(window, text="Export is successfully finished", background='white')
        self._information_label_export_ethalon_status.place(x=260, y=255)

    def get_sheets_in_working_file(self):
        try:
            file = pd.ExcelFile(self.working_file)
            self._sheets = file.sheet_names
            self._mylistbox = tk.Listbox(window)
            for item in self._sheets:
                self._mylistbox.insert(END, item)
            self._mylistbox.place(x=600,y=500)
            self._mylistbox.bind('<<ListboxSelect>>', self.onselect_sheet)
        except IndexError:
            pass
        return self._sheets

    def get_columns_actual_in_working_file(self):
        try:
            df = pd.read_excel(self.working_file,self._sheetactual)
            self._cols = df.columns
            self._mylistbox.delete(0, END)
            for item in self._cols:
                self._mylistbox.insert(END, item)
            self._mylistbox.place(x=600,y=500)
            self._mylistbox.bind('<<ListboxSelect>>', self.onselect_col)
            self._slave_column_to_change = []
        except IndexError:
            pass

    def define_key_columns_selection_in_working_file(self):
        self._df_active = pd.read_excel(self.working_file,self._sheetactual)
        self._cols = self._df_active.columns
        selection = self._mylistbox.curselection()
        if selection[0] not in self._slave_columns_selection:
            self._slave_column_to_change.append(self._cols[selection[0]])
        else: pass
        self._slave_column_to_change = pd.Series(self._slave_column_to_change).unique().tolist()
        self._information_label_working_columns_to_change.destroy()
        self._information_label_working_columns_to_change = tk.Label(window, text=self._slave_column_to_change)
        self._information_label_working_columns_to_change.place(x=290, y=320)
        print(self._slave_column_to_change)

    def confirm_column_to_change(self):

        self.confirmed_col_to_change = self._slave_column_to_change[0]
        self._information_label_working_column_to_change.destroy()
        self._information_label_working_column_to_change = tk.Label(window, text=self.confirmed_col_to_change)
        self._information_label_working_column_to_change.place(x=220, y=390)
        print(self.confirmed_col_to_change)
        self._slave_column_to_change = []

    def confirm_column_ethalon(self):

        self.confirmed_col_ethalon = self._slave_column_to_change[0]
        self._information_label_working_column_to_change.destroy()
        self._information_label_working_column_ethalon = tk.Label(window, text=self.confirmed_col_ethalon)
        self._information_label_working_column_ethalon.place(x=220, y=421)
        print(self.confirmed_col_ethalon, self.confirmed_col_to_change)
        self._slave_column_to_change = []

    def erase_confirmed_column_to_change(self):

        self.confirmed_col_to_change = []
        self._information_label_working_column_to_change.destroy()
        print(self.confirmed_col_to_change)

    def df_column_match_to_ethalon_column_by_percent(self):
        accuracy = int(self.accuracy.get())
        print(accuracy)
        df_raw = pd.read_excel(self.working_file,sheet_name=self.raw_selected_sheet_name)
        conf_change_list = df_raw[self.confirmed_col_to_change].values
        df_eth = pd.read_excel(self.working_file,sheet_name=self.ethalon_selected_sheet_name)
        conf_eth_list = df_eth[self.confirmed_col_ethalon].values
        corrected_list,problematic_items = f.list_correction_to_ethalon_naming_list(conf_change_list,conf_eth_list,accuracy)
        df_raw[f"mapped_from_{self.confirmed_col_ethalon}"] =  corrected_list
        if 'key' in df_raw.columns:
            df_raw = df_raw.rename(columns={'key':f"{self.confirmed_col_ethalon}_original",
                                            f"mapped_from_{self.confirmed_col_ethalon}":'key'})
        print(corrected_list)
        f.soft_add_sheet_to_existing_xlsx(self.working_file,df_raw,self.raw_selected_sheet_name)
        os.startfile(self.working_file)

    def create_key_column(self):
        key_list = []
        for col in self._slave_column_to_change:
            self._remove_unnesc_symb_list = []
            for string in self._df_active[col].values:

                string = string.replace("ГРИНДЕКС","")
                changed_string,mapping_item_dictionary = f.intermediate_changed_list(string,self.unnecessary_symbols_list)
                for symb in changed_string:
                    if symb in  self.unnecessary_symbols_replace_dict.keys():
                        changed_string = changed_string.replace(symb, self.unnecessary_symbols_replace_dict.get(symb))
                self._remove_unnesc_symb_list.append(changed_string)
            self._df_active[f"{col}_modified"] = self._remove_unnesc_symb_list
            key_list.append(f"{col}_modified")
        self._df_active['key'] = self._df_active[key_list[0]]
        for col in key_list[1:]:
            self._df_active['key'] = self._df_active['key'].map(str) + self._df_active[col].map(str)
        self._df_active = self._df_active.drop(columns=key_list)

        f.soft_add_sheet_to_existing_xlsx(self.working_file, self._df_active, self._sheetactual)
        os.startfile(self.working_file)

    def point_key_field(self):
        self.key_field = 'key'

    def vlookup_necessary_columns_to_raw(self):
        try:
            sheet_source = self.ethalon_selected_sheet_name
            sheet_raw = self.raw_selected_sheet_name
            df_source = pd.read_excel(self.working_file,sheet_source)
            df_current = pd.read_excel(self.working_file,sheet_raw)
            key_field = self.key_field.get()
            df = f.vlookup_column(df_current,df_source,key_field,self._slave_column_to_change)
            f.soft_add_sheet_to_existing_xlsx(self.working_file,df,self.raw_selected_sheet_name)
            os.startfile(self.working_file)
        except Exception:
            sheet_source = 'frame'
            sheet_raw = self.raw_selected_sheet_name
            df_source = pd.read_excel(self.working_file,sheet_source)
            df_current = pd.read_excel(self.working_file,sheet_raw)
            key_field = self.key_field.get()
            df = f.vlookup_column(df_current,df_source,key_field,self._slave_column_to_change)
            f.soft_add_sheet_to_existing_xlsx(self.working_file,df,self.raw_selected_sheet_name)
            os.startfile(self.working_file)

    def create_ethalon_column(self):
        self.created_ethalon_column = self._slave_column_to_change[0]
        df = pd.read_excel(self.working_file,self._sheetactual)
        created_column = df[self.created_ethalon_column].values
        changed_column = []
        for item in created_column:
            try:
                item_ = str(item).casefold().capitalize()
                changed_column.append(item_)
            except Exception:
                changed_column.append(item)
        unique_column = pd.Series(changed_column).unique()
        dubs = []
        for item in unique_column:
            count = 0
            for raw_item in changed_column:
                if item == raw_item:
                    count +=1
            dubs.append(count-1)
        df_changed = pd.DataFrame()
        df_changed[f'ethalon_processed_{self.created_ethalon_column}'] = unique_column
        df_changed[f'dublicates_count'] = dubs
        f.soft_add_sheet_to_existing_xlsx(self.working_file,df_changed,self.technical_sheet)
        print(f"changed_column: {len(unique_column)}\n created_column: {len(created_column)}")
        os.startfile(self.working_file)
    def loc_df_by_column_value(self):
        loc_value = self.pch_value_to_loc.get()
        ind = int(self.pch_col_index.get())

        print(f"So far the value is {loc_value}")
        df0 = pd.read_excel(self.working_file, self._sheetactual)
        self._df_active = df0.loc[df0[self._df_active.columns[ind]] == str(loc_value)]

        rows, cols = self._df_active.shape
        window_dataframe = tk.Tk()
        for r in range(rows):
            for c in range(cols):
                e = tk.Entry(window_dataframe,width=30)
                e.insert(0, self._df_active.iloc[r, c])
                e.grid(row=r, column=c)
                # ENTER
                e.bind('<Return>', lambda event, y=r, x=c: ff.change(self._df_active,event, y, x))
                # ENTER on keypad
                e.bind('<KP_Enter>', lambda event, y=r, x=c: ff.change(self._df_active,event, y, x))
        print(self._df_active)

    def pch_clear_df(self):
        self._df_active = pd.read_excel(self.working_file, self._sheetactual)

    def pch_show_df(self):
        df =  pd.read_excel(self.working_file, self._sheetactual)
        self._df_active = pd.DataFrame()
        for i in self._slave_column_to_change:
            self._df_active[i] = df[i]
        rows, cols = self._df_active.shape
        window_dataframe = tk.Tk()
        for r in range(rows):
            for c in range(cols):
                e = tk.Entry(window_dataframe,width=30)
                e.insert(0, self._df_active.iloc[r, c])
                e.grid(row=r, column=c)
                # ENTER
                e.bind('<Return>', lambda event, y=r, x=c: ff.change(self._df_active,event, y, x))
                # ENTER on keypad
                e.bind('<KP_Enter>', lambda event, y=r, x=c: ff.change(self._df_active,event, y, x))

    def pch_export_frame_to_excel(self):
        f.soft_add_sheet_to_existing_xlsx(self.working_file, self._df_active,self.pch_export_sheetname)
        os.startfile(self.working_file)

    def pch_replace_symbols(self):
        col_index = int(self.pch_col_index.get())
        raw = self._df_active[self._df_active.columns[col_index]].values
        dictionary = dict(zip(self.pch_col_to_replace_symbols_entry.get(),self.pch_col_to_replace_for_symbols_entry.get()))
        result = []
        for string in raw:
            new_string = ''
            for symb in string:
                if symb in dictionary.keys():
                    new_symb = dictionary.get(symb)
                    new_string += new_symb
                else:
                    new_string += symb
            result.append(new_string)
        print(result)
        self._df_active[self._df_active.columns[col_index]] = result
        rows, cols = self._df_active.shape
        window_dataframe = tk.Tk()
        for r in range(rows):
            for c in range(cols):
                e = tk.Entry(window_dataframe,width=30)
                e.insert(0, self._df_active.iloc[r, c])
                e.grid(row=r, column=c)
                # ENTER
                e.bind('<Return>', lambda event, y=r, x=c: ff.change(self._df_active,event, y, x))
                # ENTER on keypad
                e.bind('<KP_Enter>', lambda event, y=r, x=c: ff.change(self._df_active,event, y, x))
        f.soft_add_sheet_to_existing_xlsx(self.working_file,self._df_active,self.pch_export_sheetname)
        os.startfile(self.working_file)

    def perform_dataframe_checks(self):

        window_checks = tk.Toplevel()
        bg = PhotoImage(file="C:\\ageless\\migration_helper\\images\\login_background.png")
        background = Label(window_checks, image=bg)
        window_checks.title("Checks window")
        window_checks.geometry("830x700+450+1")

        self.pch_value_to_loc = tk.Entry(window_checks)
        self.pch_value_to_loc.place(x=555, y=146)
        pch_value_label = tk.Label(window_checks, text='Value to loc', bg="white")
        pch_value_label.place(x=555, y=166)

        self.pch_col_index = tk.Entry(window_checks)
        self.pch_col_index.insert(0, 0)
        self.pch_col_index.place(x=555, y=196)
        pch_col_index_label = tk.Label(window_checks, text='Column index to loc', bg="white")
        pch_col_index_label.place(x=555, y=211)

        self.pch_col_to_replace_symbols_entry = tk.Entry(window_checks)
        self.pch_col_to_replace_symbols_entry.insert(0, ["_","A"])
        self.pch_col_to_replace_symbols_entry.place(x=555, y=250)

        self.pch_col_to_replace_for_symbols_entry = tk.Entry(window_checks)
        self.pch_col_to_replace_for_symbols_entry.insert(0, ["",""])
        self.pch_col_to_replace_for_symbols_entry.place(x=555, y=300)


        clear_button = tk.Button(window_checks, text="Clear dataframe", bg="#FC0804",fg="#F9F3F3", command=self.pch_clear_df,
                                     font='Times 13')
        clear_button.place(x=10, y=90)

        open_file_button = tk.Button(window_checks, text="Loc by value", bg="#B4D2F3", fg="black", command=self.loc_df_by_column_value,
                                     font='Times 13')
        open_file_button.place(x=10, y=50)
        show_df_button = tk.Button(window_checks, text="Show dataframe", bg="#B4D2F3", fg="black", command=self.pch_show_df,
                                     font='Times 13')
        show_df_button.place(x=10, y=10)
        export_button = tk.Button(window_checks, text="Export to excel", bg="#07FE68",fg="black", command=self.pch_export_frame_to_excel,
                                     font='Times 13')
        export_button.place(x=10, y=130)
        replace_button = tk.Button(window_checks, text="Replace symbols in column", bg="#07FE68",fg="black", command=self.pch_replace_symbols,
                                     font='Times 13')
        replace_button.place(x=10, y=130)

        background.place(x=0, y=0)
        window_checks.mainloop()

window = tk.Tk()
bg = PhotoImage(file="C:\\ageless\\migration_helper\\images\\login_background.png")
background = Label(window, image=bg)

ff = CFunctions_for_app(window,background)
f = CFunctions()

window.title("Migration Helper v 0.1")
window.geometry("830x700+450+1")

ff.accuracy = tk.Entry(window)
ff.accuracy.insert(0, 75)
ff.accuracy.place(x=555, y=146)
ff.accuracy_label = tk.Label(window,text='Accuracy percent',bg="white")
ff.accuracy_label.place(x=690, y=146)

ff.key_field = tk.Entry(window)
ff.key_field.insert(0, "key")
ff.key_field.place(x=555, y=172)
ff.key_field_label = tk.Label(window,text='Key column for vlookup',bg="white")
ff.key_field_label.place(x=690, y=172)

background.place(x=0, y=0)

open_file_button = tk.Button(window,text="Open file",bg="#B4D2F3",fg="black",command=ff.get_path,font='Times 13')
open_file_button.place(x=10, y=10)


define_sheets_button = tk.Button(window,text="Get sheets",bg="#B4D2F3",fg="black",command=ff.get_sheets,font='Times 13')
define_sheets_button.place(x=10, y=44)

define_columns_button = tk.Button(window,text="Choose sheet",bg="#B4D2F3",fg="black",command=ff.get_columns_actual,font='Times 13')
define_columns_button.place(x=10, y=78)

define_selected_columns_button = tk.Button(window,text="Choose columns",bg="#B4D2F3",fg="black",command=ff.define_key_columns_selection,font='Times 13')
define_selected_columns_button.place(x=10, y=112)

show_df_button = tk.Button(window,text="Show dataframe",bg="#B4D2F3",fg="black",command=ff.show_dataframe,font='Times 13')
show_df_button.place(x=10, y=146)

to_xlsx_df_button = tk.Button(window,text="Selected raw dataframe to xlsx",bg="#FED807",fg="black",command=ff.put_selected_income_data_to_temporary_xlsx,font='Times 13')
to_xlsx_df_button.place(x=10, y=214)



switch_to_ethalon_button = tk.Button(window,text="Selected ethalon dataframe to xlsx",bg="#07FE68",fg="black",command=ff.add_selected_ethalon_data_to_temporary_xlsx,font='Times 13')
switch_to_ethalon_button.place(x=10, y=248)

get_sheets_in_working_file_button = tk.Button(window,text="Get sheets from working file",bg="#FCA65E",fg="black",command=ff.get_sheets_in_working_file,font='Times 13')
get_sheets_in_working_file_button.place(x=10, y=282)

get_cols_in_working_file_button = tk.Button(window,text="Get columns from working file",bg="#FCA65E",fg="black",command=ff.get_columns_actual_in_working_file,font='Times 13')
get_cols_in_working_file_button.place(x=10, y=316)

add_columns_actual_in_working_file_button = tk.Button(window,text="Add columns from working file for workout",bg="#FCA65E",fg="black",command=ff.define_key_columns_selection_in_working_file,font='Times 13')
add_columns_actual_in_working_file_button.place(x=10, y=350)

confirm_columns_to_change_button = tk.Button(window,text="Confirm column to change",bg="#FED807",fg="black",command=ff.confirm_column_to_change,font='Times 13')
confirm_columns_to_change_button.place(x=10, y=384)

erase_listbox_button = tk.Button(window,text="Clear main selection",bg="#FC0804",fg="#F9F3F3",command=ff.destroy_listbox,font='Times 13')
erase_listbox_button.place(x=555, y=10)

erase_working_button = tk.Button(window,text="Clear change data selection",bg="#FC0804",fg="#F9F3F3",command=ff.destroy_list_col_to_change,font='Times 13')
erase_working_button.place(x=555, y=44)

erase_confirmed_column_to_change_button = tk.Button(window,text="Clear confirmed column to change",bg="#FC0804",fg="#F9F3F3",command=ff.erase_confirmed_column_to_change,font='Times 13')
erase_confirmed_column_to_change_button.place(x=555, y=78)

confirm_columns_to_change_button = tk.Button(window,text="Confirm column ethalon",bg="#07FE68",fg="black",command=ff.confirm_column_ethalon,font='Times 13')
confirm_columns_to_change_button.place(x=10, y=419)

map_colums_button = tk.Button(window,text="Simple map columns",bg="#B4D2F3",fg="black",command=ff.df_column_match_to_ethalon_column_by_percent,font='Times 13')
map_colums_button.place(x=10, y=490)

key_colum_button = tk.Button(window,text="Create key column",bg="#B4D2F3",fg="black",command=ff.create_key_column,font='Times 13')
key_colum_button.place(x=10, y=455)

vlookup_button = tk.Button(window,text="Vlookup to raw from source",bg="#B4D2F3",fg="black",command=ff.vlookup_necessary_columns_to_raw,font='Times 13')
vlookup_button.place(x=10, y=525)

ethalon_column_creation_button = tk.Button(window,text="Ethalon column creation",bg="#B4D2F3",fg="black",command=ff.create_ethalon_column,font='Times 13')
ethalon_column_creation_button.place(x=10, y=560)

perform_checks_creation_button = tk.Button(window,text="Perform dataframe checks",bg="#B4D2F3",fg="black",command=ff.perform_dataframe_checks,font='Times 13')
perform_checks_creation_button.place(x=10, y=595)

window.mainloop()





