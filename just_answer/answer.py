import os

import pandas as pd
import numpy as np
import datetime
from openpyxl import load_workbook
path = "C:\\ageless\\just_answer\\raw_files\\test_just_answer.xlsx"

class CTransform:

    def date_transform_and_week_apply(self,df):
        years, months, days, week_num = [], [], [], []
        dt_list = []
        df['converted_visitors'] = [int(x) for x in df['converted_visitors'].values]
        df['visitors'] = [int(x) for x in df['visitors'].values]
        for rd in df['date']:
            y = int(str(rd)[:4])
            years.append(y)
            m = int(str(rd)[4:6].lstrip("0"))
            months.append(m)
            d = int(str(rd)[6:].lstrip("0"))
            days.append(d)
            dt = datetime.date(y,m,d)
            dt_list.append(dt)
            week_num.append(dt.isocalendar()[1])
        df['date_added'] = dt_list
        df['week_added'] = week_num
        return df

    def filter_by_col_value(self,df,column,value):
        df = df.loc[df[column] == value]
        return df

    def filter_by_col_not_like_value(self,df,column,value):
        df = df.loc[df[column] != value]
        return df

    def replace_nans(self,df,column,value):
        df[column] = np.nan_to_num(df[column],nan=value)
        return df

    def split_error_messages_get_unique_errors(self,df):

        unique_errors = []
        for mess in df['Error_message'].values:
            x = str(mess).split(", ")
            for x_ in x:
                z = str(x_).split("||")
                for z_ in z:
                    if z_ not in unique_errors:
                        unique_errors.append(z_)


        return df,unique_errors

    def calculate_presence_of_unique_errors_in_entries(self,df,unique_errors):
        unique_errors = unique_errors[:-1]
        unique_errors.append("NO ERRORS")
        df['Error_message'] = df['Error_message'].fillna("NO ERRORS")
        for error in unique_errors:
            count = []
            for entry in df['Error_message'].values:
                if error in entry:
                    count.append(1)
                else: count.append(0)
            df[error] = count
        return df

    def melt_df(self,df, hold_columns_list, feature_rename_to, value_rename_to):
        df = df.melt(hold_columns_list).rename(columns={'variable':feature_rename_to,'value':value_rename_to})
        return df

    def soft_add_sheet_to_existing_xlsx(self,full_path,df,sheet_name):
        sheet = df
        book = load_workbook(full_path)
        writer = pd.ExcelWriter(full_path, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        sheet.to_excel(writer, sheet_name, index=False)
        writer.save()


if __name__ == '__main__':
    ct = CTransform()
    df = pd.read_excel(path)
    print(df.shape) # (3348, 7)

    week_num_to_analize = 1

    df = ct.date_transform_and_week_apply(df) # transform date to datetime and week_num
                                              # calculation, convert visitors quantity and conversion to int

    df = ct.filter_by_col_value(df,'week_added',week_num_to_analize) # filter dataframe by week_num requested
    print(df.shape) # (2356, 9)

    df = ct.replace_nans(df,'number_of_errors',0.0) #replace nans with 0.0
    #df = ct.filter_by_col_not_like_value(df, 'number_of_errors', 0.0) #filter dataframe. As we need df with errors
    print(df.shape) # (2156, 9)

    df,unique_errors = ct.split_error_messages_get_unique_errors(df) # get unique errors list
    df = ct.calculate_presence_of_unique_errors_in_entries(df,unique_errors) # calculate presence of each unique
                                                                             # error and add to dataframe


    raw_sheet_name_ = 'filtered_raw'
    #df.to_excel('test_solution.xlsx',raw_sheet_name_,index=False)
    ct.soft_add_sheet_to_existing_xlsx('test_solution.xlsx', df, raw_sheet_name_) # export filtered dataframe

    hold_columns_list = ['visitors', 'deviceCategory', 'medium', 'number_of_errors',
       'Error_message', 'converted_visitors', 'date', 'date_added',
       'week_added']
    feature_rename_to = 'error_unique'
    value_rename_to = 'number_of_errors_added'

    df_melted = ct.melt_df(df, hold_columns_list, feature_rename_to, value_rename_to) # create flat table and add to
                                                                                      # file
    raw_sheet_name = 'melted_raw'
    ct.soft_add_sheet_to_existing_xlsx('test_solution.xlsx', df_melted, raw_sheet_name)

    df_med = pd.DataFrame()
    df_med['medium'] = df_melted['medium'].unique()
    raw_sheet_name = 'medium_dictionary'
    ct.soft_add_sheet_to_existing_xlsx('test_solution.xlsx', df_med, raw_sheet_name)

    df_dev = pd.DataFrame()
    df_dev['deviceCategory'] = df_melted['deviceCategory'].unique()
    raw_sheet_name = 'devices_dictionary'
    ct.soft_add_sheet_to_existing_xlsx('test_solution.xlsx', df_dev, raw_sheet_name)

    os.startfile('test_solution.xlsx')


