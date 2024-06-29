# Import of the relevant modules

import ifcopenshell
import ifcopenshell.util
import ifcopenshell.util.element

import pandas as pd

# Helpers import (custom classes / objects)

import Helpers.config_data as cod
import Helpers.export_data as ed

# global variables - lists

conf_lst = []
export_lst = []


### STEP 1

# load files > ifc + excel 

print('Start: Starte Grunddateien einlesen...')

source_path = input('Bitte Name der IFC Datei eingeben inkl. Dateiendung.')
ifc_file = ifcopenshell.open(source_path)

xls_file = ('./Elementplan_Config.xlsx')

s_data_category = pd.read_excel(xls_file, sheet_name='Objektkatalog')
s_data_attributes = pd.read_excel(xls_file, sheet_name='Attributsliste')

# read data from config file and structured save

for index, row in s_data_category.iterrows():

    attr_df = s_data_category.iloc[index, 2] # könnte auch mit der row[] funktion gemacht werden

    if attr_df: # ist attr_df true?

        attr_in_group = s_data_attributes[s_data_attributes['Gruppe'] == attr_df]

        print('Klasse: ' + row['IfcClass'])
        print(attr_in_group)
        print('Next')

        prop_lst = []

        for _i, _r in attr_in_group.iterrows():

            p_h = cod.prop_holder(_r['Pset'], _r['Property'])
            prop_lst.append(p_h)


        _target_obj = cod.filterConfig(row['IfcClass'], conf_lst)

        if _target_obj != None:
            _target_obj.prop_list.extend(prop_lst)
            print('Adding')
        else:
            d_h = cod.data_holder(row['IfcClass'], prop_lst)
            conf_lst.append(d_h)
            print('creating new')


### STEP 2

# ifc data read and structured save

for _c in conf_lst:
    elements = ifc_file.by_type(_c.category)

    tmp_exp_data_from_current_config = []

    for _e in elements:

        tmp_exp_data = []

        try:
            tmp_exp_data.append(_e.get_info().get('GlobalId')) 
            tmp_exp_data.append(_e.get_info().get('type')) #type ist ifc Klasse, z.B. IfcSpace
        except:
            continue

        act_prop = ifcopenshell.util.element.get_psets(_e, psets_only=False)

        for _p in _c.prop_list:
            try:
                act_val = act_prop.get(_p.pset).get(_p.prop)
                tmp_exp_data.append(act_val)
            except:
                act_val = ''
                tmp_exp_data.append(act_val)

            tmp_exp_data_from_current_config.append(tmp_exp_data)

    exp_holder = ed.Exp_Holder(_c.category, tmp_exp_data_from_current_config)
    export_lst.append(exp_holder)


 ### STEP 3  

exp_path = './export_datei_ifc.xlsx'
pd.DataFrame([]).to_excel(exp_path)

 # export of extracted data

if len(export_lst) > 0:

    for exp_set in export_lst:

        col_lst = []
        col_lst.append('GUID')
        col_lst.append('Kategorie')

        _target_obj = cod.filterConfig(exp_set.branch, conf_lst)

        if _target_obj != None:

            for _p_l in _target_obj.prop_list:
                #print(_p_l)
                _ps = _p_l.pset
                _pr = _p_l.prop

                _title = '{}:{}'.format(_ps, _pr)

                col_lst.append(_title)

            df = pd.DataFrame(exp_set.data, columns = col_lst)


            with pd.ExcelWriter(exp_path, engine='openpyxl', if_sheet_exists='replace', mode='a') as writer:
                df.to_excel(writer, sheet_name= exp_set.branch, index=False)

    print('Letzter Schritt: Excel Export fertig gestellt')
