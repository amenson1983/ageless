
import os
import re
import pandera
from markupsafe import Markup
from pandera import Column, Check
import numpy as np
import pandas as pd
from fuzzywuzzy import process
from openpyxl import load_workbook

import altair as alt
import datapane as dp
from jinja2.utils import markupsafe
markupsafe.Markup()
Markup('')


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

    def divide_col_by_col(self,df,col_div,col_to_div,new_col_name):
        new_ = []
        col_d = df[col_div].values
        col_to_d = df[col_to_div].values
        for x in range(0,len(col_d)):
            try:
                z = float(col_d[x]) / float(col_to_d[x]) / 30
                new_.append(z)
            except ValueError:
                new_.append(0.00)
        df[new_col_name] =new_
        return df


class CAnalisys:
    def __init__(self,full_path):
        self.full_path = full_path
        self.df = pd.read_excel(full_path)
        self.sheet_active = ''
        self.columns = self.df.columns
        self.test_med_file = "mid.xlsx"


    def get_df_and_columns(self):
        try:
            self.df = pd.read_excel(self.full_path)
            columns_list = self.df.columns
        except Exception:
            self.df = pd.read_csv(self.full_path)
            columns_list = self.df.columns
        print(f"{20 * '*'} COLUMNS {20 * '*'}")
        print(columns_list)
        return self.df

    def get_break_even_period(self):
        file = pd.ExcelFile(self.full_path)
        self.sheets = file.sheet_names
        print(self.sheets)
        self.df = pd.read_excel(self.full_path)
        self.sheet_active = 'months_for_break_even'
        self.df = f.divide_col_by_col(self.df,'total_price','profit_24_h',self.sheet_active)
        return self.df

    def get_unique_values_and_lenght_for_columns(self,df):
        for col in df.columns:
            print(f"{20 * '*'} COLUMN: {col} {20 * '*'}")
            print(f"1. Qunatity of unique values: {len(self.df[col].unique())} entity/ies")
            print(f"")
            print(f"2. Description:\n{self.df[col].describe()}")
            print(f"")
            print(f"3. Values unique:\n{self.df[col].unique()}")
            print(f"")

    def get_value_by_column_by_percentile(self,df,col,percentile_min,percentile_max):
        percentile_value_min = np.percentile(np.array(df[col]), percentile_min, axis=None, out=None)
        percentile_value_max = np.percentile(np.array(df[col]), percentile_max, axis=None, out=None)
        df = df.loc[df[col] > percentile_value_min]
        df = df.loc[df[col] < percentile_value_max]
        print(df.shape)
        print(df.head)
        return df

    def filter_df_by_column_by_mean_value(self,df,col,dev):
        average_value = df[col].mean() + dev
        step_value = average_value - dev
        print(average_value)
        df = df.loc[df[col] < average_value]
        df = df.loc[df[col] > step_value]
        print(df.shape)
        return df

    def filter_df_by_column_equal_value(self,df,col,value):
        df = df.loc[df[col] == value]
        print(df.shape)
        return df

    def filter_df_by_column_by_min_max_value(self,df,col,min,max):
        df = df.loc[df[col] < max]
        df = df.loc[df[col] > min]
        print(df.shape)
        return df

    def put_to_datapane_web(self,df):
        dp.Report(
            # dp.Plot(plot),
            dp.DataTable(df)
        ).upload(name='Test',
                 open=True)  # edit your report at https://datapane.com/u/amenson1983/reports/dkjLWe3/test/edit/
        # View and share your report at https://datapane.com/u/amenson1983/reports/dkjLWe3/test/

    def check_df_for_nans(self,df):
        [print(df[col].hasnans) for col in df.columns]  # Check columns for NaN-s

    def get_sheets(self,path_):
        file = pd.ExcelFile(path_)
        sheets = file.sheet_names
        print(sheets)
        return sheets

    def transform_mean(self,df,new_col_name_for_mean,groupby_list,to_calculate_col):
        df[new_col_name_for_mean] = df.groupby(groupby_list)[to_calculate_col].transform("mean")
        return df

    def transform_sum(self,df,new_col_name_for_mean,groupby_list,to_calculate_col):
        df[new_col_name_for_mean] = df.groupby(groupby_list)[to_calculate_col].transform("sum")
        return df

f = CFunctions()


def cards_analisys_selection():
    df = a.get_break_even_period()
    # df = a.get_value_by_column_by_percentile(df,'months_for_break_even',25,30)
    a.get_unique_values_and_lenght_for_columns(df)
    # df = a.filter_df_by_column_by_mean_value(df,'months_for_break_even',2)
    df['months_for_break_even'] = [int(x) for x in df['months_for_break_even'].values]
    df = a.filter_df_by_column_by_min_max_value(df, 'months_for_break_even', 5, 15)
    # df = a.filter_df_by_column_by_mean_value(df, 'total_price', 500)
    df = a.filter_df_by_column_by_min_max_value(df, 'total_price', 400, 1000)
    df = a.filter_df_by_column_equal_value(df, 'condition', 'Brand New')

    f.soft_add_sheet_to_existing_xlsx(path_, df, 'calculated_break_even_df')
    #os.startfile(path_)


def analyze_cards_and_show_in_datapane():
    path_ = "C:\\ageless\\migration_helper\\raw_files_folder\\cards_costs.xlsx"
    df = pd.read_excel(path_, 'calculated_break_even_df')
    a.get_sheets(path_)
      # View and share your report at https://datapane.com/u/amenson1983/reports/dkjLWe3/test/
    # a.check_df_for_nans(df)
    new_col_name_for_mean = 'average_month_for_card'
    groupby_list = ["card", "profit_24_h"]
    to_calculate_col = "months_for_break_even"
    df = a.transform_mean(df, new_col_name_for_mean, groupby_list, to_calculate_col)
    acc_per = []
    #a.put_to_datapane_web(df)
    for ind in range(0,len(df['months_for_break_even'].values)):
        if df['months_for_break_even'].values[ind] < df['average_month_for_card'].values[ind]:
            acc_per.append("excellent")
        elif df['months_for_break_even'].values[ind] == df['average_month_for_card'].values[ind]:
            acc_per.append("good")
        else: acc_per.append("need_to_consider")


    df['accent_to_period'] = acc_per

    df.to_excel(a.test_med_file, index=False)
    os.startfile(a.test_med_file)
    #a.put_to_datapane_web(df)


if __name__ == "__main__":
    path_ = "C:\\ageless\\migration_helper\\raw_files_folder\\cards_costs.xlsx"
    a = CAnalisys(path_)
    cards_analisys_selection()
    analyze_cards_and_show_in_datapane()
    '''path_ = "C:\\ageless\\migration_helper\\raw_files_folder\\Legacy Data CMD V2.xlsx"
    df_s = pd.read_excel(path_,"TIBAN")
    df_e = pd.read_excel(path_, "KBNK")
    df_merged = pd.merge(df_s,df_e,on="BANK Key")
    print(df_merged.columns)'''




