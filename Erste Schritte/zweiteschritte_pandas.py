import pandas as pd

### Datei lesen

#excel_file_path = 'Raumliste_Demo.xlsx' # Wenn die Datei im gleichen Ordner ist
excel_file_path = 'Raumliste_Demo.xlsx'

### Datei in Pandas-Modul einlesen

base_df = pd.read_excel(excel_file_path, sheet_name='Sheet1')
#print(base_df)



### Iteration über die einzelnen Reihen und aus 2 Einträgen eine neue generieren

sig_code_lst = []

for index,row in base_df.iterrows():

    act_roomcode = row['Raumcode']
    act_storycode = row['Geschosscode']
    sig_code = f'{act_storycode}__{act_roomcode}'
    sig_code_lst.append(sig_code)

base_df['Signaletik'] = sig_code_lst

export_df = base_df.loc[ : ,['Building Story','Nettoraumvolumen [m³]','Signaletik']]


filt_df_iloc = base_df.iloc[0:10, base_df.columns.get_loc('SIA416')] # Get loc holt die Indexnr von der gwünschten Spalte


export_df.to_excel('Raumliste_Signaletik.xlsx',sheet_name='daten',index=False)