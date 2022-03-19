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

        df_error = pd.DataFrame()
        df_error['problematic_items'] = problematic_items

        df.to_excel(resulting_file, sheet_name='results', engine='openpyxl', index=False)
        f.soft_add_sheet_to_existing_xlsx(resulting_file, df_error, 'problematic_items')
        os.startfile(resulting_file)
        print(f"Corrected list with applied actual percent threshold {percent}%: \n{corrected_list}")


f = CFunctions()
o = COperations()





if __name__ == '__main__':

    #incoming_list = ['prague_e,','kiev','kopenhagen','stock_gohlm','paris ','berlin','Berdichev']
    #ethalon_naming_list = ['Berlin','Kyiv','Prague','Kopenhagen','Paris','Stockgohlm','Berdychiv']

    percent = 75
    unnecessary_symbols_list = ["№","_","%","-","/","|",",",".",".",",","!"," "]
    resulting_file = 'result.xlsx'
    df = pd.read_excel('test.xlsx',engine='openpyxl')
    incoming_list = df['item_sales_report'].values
    ethalon_naming_list = df['item_kpi_report'].values

    o.complex_mapping_to_ethalon(incoming_list,ethalon_naming_list,resulting_file,unnecessary_symbols_list)




