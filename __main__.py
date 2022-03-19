import os
from fuzzywuzzy import process
import pandas as pd
import numpy as np
import jellyfish as jf
from openpyxl import load_workbook

class CFunctions:
    def item_match_in_list_by_percent(self,item,list_values,percent):
        s = process.extractOne(item, list_values)
        matched_value = ''
        if s[1] >= percent:
            matched_value = s[0]
        return matched_value

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


f = CFunctions()
o = COperations()
if __name__ == '__main__':

    incoming_list = ['praguee','kiev','kopenhagen','stock_gohlm','paris','berlin','Berdichev']
    ethalon_naming_list = ['Berlin','Kyiv','Prague','Kopenhagen','Paris','Stockgohlm','Berdychiv']
    percent = 75

    corrected_list,problematic_items = o.list_correction_to_ethalon_naming_list(incoming_list,ethalon_naming_list,percent)
    print(f"Corrected list with applied actual percent threshold {percent}%: \n{corrected_list}")



