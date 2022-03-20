import os
from fuzzywuzzy import process
import pandas as pd
import numpy as np
import jellyfish as jf
import string as st
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
        df[new_column_name] = [dictionary.get(x) for x in df[target_column].values]
        return df

class COperations:

    def list_correction_to_ethalon_naming_list(self,incoming_list,ethalon_naming_list,percent):
        corrected_list, problematic_items = [], {}

        for item in incoming_list:
            value = f.item_match_in_list_by_percent(item, ethalon_naming_list, percent)
            if value == '':
                s = process.extractOne(item, ethalon_naming_list)
                problematic_items.update({item:s[1]})
            corrected_list.append(value)
        if problematic_items != {}:
            print(f"Minimal percent: {min(problematic_items.values())}%")
        if problematic_items != {}:
            print(f"COGI: {problematic_items}")
        return corrected_list,problematic_items

    def complex_mapping_to_ethalon(self,incoming_list_1,ethalon_naming_list_1,resulting_file,unnecessary_symbols_list):
        mapping_item_dictionary_incoming, mapping_item_dictionary_ethalon = {}, {}
        incoming_list_2, ethalon_naming_list_2 = [], []
        df = pd.DataFrame()
        for item in incoming_list_1:
            changed_string, mapping_item_dictionary = f.intermediate_changed_list(item, unnecessary_symbols_list)
            mapping_item_dictionary_incoming.update(mapping_item_dictionary)
            incoming_list_2.append(changed_string)
        for item_ in ethalon_naming_list_1:
            changed_string_, mapping_item_dictionary_ = f.intermediate_changed_list(item_, unnecessary_symbols_list)
            mapping_item_dictionary_ethalon.update(mapping_item_dictionary_)
            ethalon_naming_list_2.append(changed_string_)

        corrected_list, problematic_items = o.list_correction_to_ethalon_naming_list(incoming_list_2,
                                                                                     ethalon_naming_list_2,
                                                                                     percent)
        corrected_list = [mapping_item_dictionary_ethalon.get(x) for x in corrected_list]
        problematic_items = [mapping_item_dictionary_incoming.get(x) for x in problematic_items]
        if problematic_items != []:
            print(problematic_items)
        df['incoming_list_1'] = incoming_list_1
        df['result'] = corrected_list
        mapping = dict(zip(incoming_list_1,corrected_list))
        df_error = pd.DataFrame()
        df_error['problematic_items'] = problematic_items

        df.to_excel(resulting_file, sheet_name='results', engine='openpyxl', index=False)
        f.soft_add_sheet_to_existing_xlsx(resulting_file, df_error, 'problematic_items')
        os.startfile(resulting_file)
        print(f"Corrected list with applied actual percent threshold {percent}%: \n{corrected_list}")
        return mapping

    def dataframe_two_field_progressive_key(self,df,two_columns_list,unnecessary_symbols_list):

        df = f.key_field_two_columns_insertion_to_dataframe(df, two_columns_list,
                                                            'progressive_key')
        changed_string_list = []
        for key in df['progressive_key'].values:
            changed_string, mapping_item_dictionary = f.intermediate_changed_list(key, unnecessary_symbols_list)
            changed_string_list.append(changed_string)
        df['progressive_key'] = changed_string_list
        return df

    def ethalon_field_creation(self,df,target_field,support_field,ethalon_field_name):
        df[support_field] = [int(x) for x in df[support_field].values]
        target_field_list = df[target_field].values
        support_field = df[support_field].values
        dictionary = dict(zip(support_field,target_field_list))
        df_final = pd.DataFrame()
        df_final['support_field'] = [int(x) for x in dictionary.keys()]
        df_final['target_field'] = [str(x).capitalize() for x in dictionary.values()]
        dictionary = dict(zip(df_final['support_field'].values, df_final['target_field'].values))
        df[ethalon_field_name] = [dictionary.get(x) for x in df['Postal Code'].values]
        return df,dictionary

f = CFunctions()
o = COperations()

if __name__ == '__main__':
    percent = 72
    unnecessary_symbols_list = ["№","_","%","-","/","|",",",".",".",",","!"," "]
    resulting_file = 'result.xlsx'

    tab_sap = 'SAP Data'
    df_ethalon = pd.read_excel('Comparison test_20220210.xlsx',sheet_name=tab_sap,engine='openpyxl')
    ethalon_target_columns = ['Country', 'City', 'Street','Postal Code']
    #print(df_ethalon.columns)

    tab_legacy = 'Legacy Data'
    df_legacy = pd.read_excel('Comparison test_20220210.xlsx', sheet_name=tab_legacy, engine='openpyxl')
    legacy_target_columns = ['Country', 'City', 'Street','Postal Code']
    #print(df_legacy.columns)

    incoming_list = df_legacy[legacy_target_columns[1]].values
    ethalon_naming_list = df_ethalon[legacy_target_columns[1]].values

    #mapping_dictionary = o.complex_mapping_to_ethalon(incoming_list,ethalon_naming_list,resulting_file,unnecessary_symbols_list)
    #df = o.dataframe_two_field_progressive_key(df,['item_sales_report','item_kpi_report'],unnecessary_symbols_list)
    ethalon_field_name = 'city_ethalon'
    df_legacy_with_common_city,dictionary_postcode_common_city = o.ethalon_field_creation(df_legacy,legacy_target_columns[1],legacy_target_columns[3],ethalon_field_name) #City ethalon for SAP DATA
    df_ethalon[ethalon_field_name] = [dictionary_postcode_common_city.get(x) for x in df_ethalon[legacy_target_columns[3]].values]
    df_legacy_with_common_city.to_excel(resulting_file,sheet_name=tab_legacy,engine='openpyxl', index=False)
    f.soft_add_sheet_to_existing_xlsx(resulting_file,df_ethalon, tab_sap)
    os.startfile(resulting_file)


