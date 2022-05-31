import os
from fuzzywuzzy import process
import pandas as pd
import numpy as np
import jellyfish as jf
import string as st
from openpyxl import load_workbook

class CInformations:

    def information_input_for_transfer_check(self,unnecessary_symbols_list):
        percent = 82
        percent_names = 70

        resulting_file = 'result_LEGACY_SAP.xlsx'
        analisys_tab = 'Analisys tab'
        tab_sap = 'SAP Data'
        df_ethalon = pd.read_excel('Comparison test_20220210.xlsx', sheet_name=tab_sap, engine='openpyxl')
        ethalon_target_columns = ['Country', 'City', 'Postal Code']

        print(df_ethalon.columns)

        tab_legacy = 'Legacy Data'
        df_legacy = pd.read_excel('Comparison test_20220210.xlsx', sheet_name=tab_legacy, engine='openpyxl')
        legacy_target_columns = ['Country', 'City', 'Postal Code']
        print(df_legacy.columns)

        return df_legacy, df_ethalon, legacy_target_columns, percent, percent_names, unnecessary_symbols_list, resulting_file, tab_sap, tab_legacy,analisys_tab

    def information_input_for_TIBAN_KBNK_check(self,unnecessary_symbols_list):
        percent = 82
        percent_names = 70

        resulting_file = 'result_TIBAN_KBNK.xlsx'

        tab_sap = 'Bank Data UPLOAD'
        df_sap = pd.read_excel('Legacy Data CMD V2.xlsx', sheet_name=tab_sap, engine='openpyxl')

        print(df_sap.columns)

        tab_kbnk = 'KBNK'
        df_kbnk = pd.read_excel('Legacy Data CMD V2.xlsx', sheet_name=tab_kbnk, engine='openpyxl')
        kbnk_key_columns = ['BANK Key', 'BANK Account']
        kbnk_columns_to_migrate = ['Bank partner Type', 'CollectionAuthorization (KNBK)']

        print(df_kbnk.columns)

        tab_tiban = 'TIBAN'
        df_tiban = pd.read_excel('Legacy Data CMD V2.xlsx', sheet_name=tab_tiban, engine='openpyxl')
        tiban_key_columns = ['BANK Key', 'BANK Account']
        print(df_tiban.columns)

        return df_sap, df_kbnk, df_tiban, percent, percent_names, unnecessary_symbols_list, \
               resulting_file, tab_sap, tab_kbnk, tab_tiban,kbnk_key_columns,kbnk_columns_to_migrate,tiban_key_columns

class CFunctions:
    def check_columns_for_nans(self,df):
        [print(df[col].hasnans) for col in df.columns]  # Check columns for NaN-s

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

    def complex_mapping_to_ethalon(self,incoming_list_1,ethalon_naming_list_1,resulting_file,unnecessary_symbols_list,percent):
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

        #f.soft_add_sheet_to_existing_xlsx(resulting_file, df, 'results_mapping')
        #f.soft_add_sheet_to_existing_xlsx(resulting_file, df_error, 'problematic_items')
        #os.startfile(resulting_file)
        print(f"Corrected list with applied actual percent threshold {percent}%")
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

    def ethalon_target_field_creation_with_support_field(self,df,target_field,support_field,ethalon_field_name):
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

    def name_one_name_two_mapping(self,df_legacy_with_common_city,df_ethalon,names_field_list,resulting_file,
                                  unnecessary_symbols_list,percent_names):

        df_legacy_with_common_city = f.key_field_two_columns_insertion_to_dataframe(df_legacy_with_common_city,
                                                                                    names_field_list,'names_key')


        df_ethalon = f.key_field_two_columns_insertion_to_dataframe(df_ethalon, names_field_list,'names_key')

        names_field_list_ = ['key_ethalon', 'names_key']

        df_legacy_with_common_city = f.key_field_two_columns_insertion_to_dataframe(df_legacy_with_common_city,
                                                                                    names_field_list_,'total_names_key')

        df_ethalon = f.key_field_two_columns_insertion_to_dataframe(df_ethalon, names_field_list_,'total_names_key')

        incoming_list = df_ethalon['total_names_key'].values
        ethalon_list = df_legacy_with_common_city['total_names_key'].values


        mapping_dict = o.complex_mapping_to_ethalon(incoming_list,ethalon_list,resulting_file,unnecessary_symbols_list,percent_names)
        df_ethalon['names_ethalon'] = [mapping_dict.get(x) for x in df_ethalon['total_names_key'].values]
        df_ethalon = df_ethalon.drop(columns=['names_key','key_ethalon','total_names_key'])
        df_legacy_with_common_city = df_legacy_with_common_city.drop(columns=['names_key', 'key_ethalon'])

        return df_legacy_with_common_city,df_ethalon

    def kbnk_to_tiban_vlookup(self,df_tiban,df_kbnk,key_field,columns_to_migrate,resulting_file):
        df_tiban,df_kbnk = f.map_data_to_first_df_from_second_by_key(df_tiban, df_kbnk, key_field, columns_to_migrate)
        df_tiban = df_tiban.drop(columns=[key_field])

        df_tiban.to_excel(resulting_file,engine='openpyxl',sheet_name='TIBAN',index=False)
        #os.startfile(resulting_file)
        return df_kbnk, df_tiban,resulting_file

    def tiban_to_bank_data_upload_vlookup(self,df_sap,df_tiban,key_field,columns_to_migrate,
                                          resulting_file,tab_sap):
        df_sap,df_tiban= f.map_data_to_first_df_from_second_by_key(df_sap, df_tiban, key_field, columns_to_migrate)
        sap_iban_list = df_sap[key_field].values
        tiban_iban_list = df_tiban[key_field].values
        sap_iban_count_tiban_iban_match_dict = {}
        for sap_iban in sap_iban_list:
            count = 0
            for tiban_iban in tiban_iban_list:
                if sap_iban == tiban_iban:
                    count += 1
            sap_iban_count_tiban_iban_match_dict.update({sap_iban:count})
        df_sap = f.map_dataframe_column_via_dictionary_and_get_new_df(df_sap,'IBAN','matches_in_TIBAN',sap_iban_count_tiban_iban_match_dict)
        f.soft_add_sheet_to_existing_xlsx(resulting_file,df_sap,tab_sap)
        os.startfile(resulting_file)



def transfer_check(df_legacy,df_ethalon,legacy_target_columns,percent,percent_names,unnecessary_symbols_list,resulting_file,
                   tab_sap,tab_legacy):

    ethalon_field_name = 'city_ethalon'
    df_legacy_with_common_city, dictionary_postcode_common_city = o.ethalon_target_field_creation_with_support_field(
        df_legacy, legacy_target_columns[1], legacy_target_columns[2], ethalon_field_name)  # City ethalon for SAP DATA
    df_ethalon[ethalon_field_name] = [dictionary_postcode_common_city.get(x) for x in
                                      df_ethalon[legacy_target_columns[2]].values]

    ethalon_target_columns_ = ['Country', ethalon_field_name, 'Postal Code', 'Street']
    legacy_target_columns_ = ['Country', ethalon_field_name, 'Postal Code', 'Street']
    df_ethalon = f.key_field_four_columns_insertion_to_dataframe(df_ethalon, ethalon_target_columns_, 'key_full')
    df_legacy_with_common_city = f.key_field_four_columns_insertion_to_dataframe(df_legacy_with_common_city,
                                                                                 ethalon_target_columns_, 'key_full')

    df_legacy_streets_list = df_legacy_with_common_city['key_full'].values
    df_sap_streets_list = df_ethalon['key_full'].values

    mapping_dictionary = o.complex_mapping_to_ethalon(df_sap_streets_list, df_legacy_streets_list, resulting_file,
                                                      unnecessary_symbols_list, percent)

    ethalon_field_name_full_key = 'key_ethalon'
    df_ethalon[ethalon_field_name_full_key] = [mapping_dictionary.get(x) for x in df_ethalon['key_full'].values]
    df_legacy_with_common_city[ethalon_field_name_full_key] = df_legacy_streets_list

    ethalon_field_name = 'streets_ethalon'
    dict_key_full_to_street = dict(zip(df_legacy_with_common_city[ethalon_field_name_full_key].values,
                                       df_legacy_with_common_city[legacy_target_columns_[3]].values))

    df_ethalon[ethalon_field_name] = [dict_key_full_to_street.get(x) for x in
                                      df_ethalon[ethalon_field_name_full_key].values]

    df_legacy_with_common_city[ethalon_field_name] = df_legacy_with_common_city[legacy_target_columns_[3]]

    df_legacy_with_common_city = df_legacy_with_common_city.drop(columns=['key_full'])
    df_ethalon = df_ethalon.drop(columns=['key_full'])
    names_field_list = ['Name 1', 'Name 2']
    df_legacy_with_common_city,df_ethalon = o.name_one_name_two_mapping(df_legacy_with_common_city,df_ethalon,
                                                                        names_field_list,resulting_file,unnecessary_symbols_list,percent_names)

    df_legacy_with_common_city.to_excel(resulting_file, sheet_name=tab_legacy, engine='openpyxl', index=False)
    f.soft_add_sheet_to_existing_xlsx(resulting_file, df_ethalon, tab_sap)
    return df_legacy_with_common_city,df_ethalon, resulting_file


def LegacySapWorkout(unnecessary_symbols_list):
    df_legacy, df_ethalon, legacy_target_columns, percent, percent_names, \
    unnecessary_symbols_list, resulting_file, tab_sap, tab_legacy, analisys_tab = i.information_input_for_transfer_check(unnecessary_symbols_list)
    df_legacy_with_common_city, df_ethalon, resulting_file = transfer_check(df_legacy, df_ethalon,
                                                                            legacy_target_columns, percent,
                                                                            percent_names, unnecessary_symbols_list,
                                                                            resulting_file,
                                                                            tab_sap, tab_legacy)
    keyfield = 'key_for_analisys'
    df_legacy_with_common_city = f.key_field_three_columns_insertion_to_dataframe(df_legacy_with_common_city,
                                                                                  ['Postal Code', 'city_ethalon',
                                                                                   'streets_ethalon'],
                                                                                  keyfield)
    df_ethalon = f.key_field_three_columns_insertion_to_dataframe(df_ethalon,
                                                                  ['Postal Code', 'city_ethalon', 'streets_ethalon'],
                                                                  keyfield)
    df_ethalon = df_ethalon.rename(columns={"Name 1": "Name 1 SAP", "Name 2": "Name 2 SAP"})
    df_ethalon, df_legacy_with_common_city = f.map_data_to_first_df_from_second_by_key(df_ethalon,
                                                                                       df_legacy_with_common_city,
                                                                                       keyfield,
                                                                                       ['Name 1',
                                                                                                           'Name 2'])
    df_ethalon = df_ethalon.drop(columns=['key_for_analisys'])
    df_ethalon = df_ethalon.rename(columns={"Name 1": "Name 1 Legacy", "Name 2": "Name 2 Legacy"})
    f.soft_add_sheet_to_existing_xlsx(resulting_file, df_ethalon, tab_sap)
    df_analisys = df_ethalon
    df_analisys['names_ethalon'] = df_analisys['names_ethalon'].fillna('Have doubts')
    df_analisys = f.loc_df_by_column_equals_to(df_analisys, 'names_ethalon', 'Have doubts')
    f.soft_add_sheet_to_existing_xlsx(resulting_file, df_analisys, analisys_tab)
    os.startfile(resulting_file)


def TibanKnbkUploadCheck(unnecessary_symbols_list):
    i = CInformations()
    df_sap, df_kbnk, df_tiban, percent, percent_names, \
    unnecessary_symbols_list, resulting_file, tab_sap, \
    tab_kbnk, tab_tiban, kbnk_key_columns, kbnk_columns_to_migrate, tiban_key_columns = i.information_input_for_TIBAN_KBNK_check(unnecessary_symbols_list)
    key_col_list = ['BANK Key', 'BANK Account']
    df_kbnk = o.dataframe_two_field_progressive_key(df_kbnk, key_col_list, unnecessary_symbols_list)
    df_tiban = o.dataframe_two_field_progressive_key(df_tiban, key_col_list, unnecessary_symbols_list)
    df_kbnk, df_tiban, resulting_file = o.kbnk_to_tiban_vlookup(df_tiban, df_kbnk, 'progressive_key',
                                                                kbnk_columns_to_migrate, resulting_file)
    key_field = "IBAN"
    clean_list = []
    for string in df_sap[key_field].values:
        changed_string,mapping_item_dictionary = f.intermediate_changed_list(string,unnecessary_symbols_list)
        clean_list.append(changed_string.upper())
    df_sap[key_field] = clean_list
    o.tiban_to_bank_data_upload_vlookup(df_sap, df_tiban, key_field, kbnk_columns_to_migrate, resulting_file,tab_sap)


f = CFunctions()
o = COperations()
i = CInformations()

if __name__ == '__main__':

    unnecessary_symbols_list = ["№", "_", "%", "/", "|", ",", ".", ".", ",", "!", " "]

    #LegacySapWorkout(unnecessary_symbols_list) #Legacy Data - SAP Data migration entries check

    #TibanKnbkUploadCheck(unnecessary_symbols_list) # TIBAN - KNBK - UPLOAD check
    df_current = pd.DataFrame()
    df_current['number'] = [1,2,3,4]

    df_source = pd.DataFrame()
    df_source['number'] = [1,2,3,4,1]
    df_source['value'] = ['Азитромицин','Апилак ,  мазь, 10мг/г, 50г',
                          'Атракурий Калцекс ,р-р д/ин, 10мг/мл по 5 мл  ',
                          'АукСИЛен®, 50 мг/2 мл по 2 мл',
                          'Азитромицин']
    df_source['sales_packs'] = [4,3,2,0,10]
    df_source['sales_euro'] = df_source['sales_packs'] * 9.17
    key_field, columns_to_migrate = 'number',['value']
    columns_to_sum = ['sales_packs','sales_euro']

    ethalon_map_values = ['Азитромицин Гриндекс табл. 500мг №3',
                          'Апилак ,  мазь, 10мг/г, 50г в тубе',
                          'Атракурий Калцекс ,р-р д/ин, 10мг/мл по 5 мл в ампулах №5',
                          'Ауксилен®,  50 мг/2 мл по 2 мл в ампулах №5']

    values_mapped_list,problematic_items = o.list_correction_to_ethalon_naming_list(df_source['value'].values,ethalon_map_values,77)
    df_source['value'] = values_mapped_list
    temp_file = 'temp.xlsx'

    df_current,temp_file_ = f.vlookup_column(df_current,df_source,key_field,columns_to_migrate,temp_file)
    temp_file = 'temp1.xlsx'
    df_current,temp_file = f.sumif_column(df_current,df_source,key_field,columns_to_sum,temp_file)



    f.soft_add_sheet_to_existing_xlsx()

