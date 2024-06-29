import pandas as pd

### Datei lesen

#excel_file_path = 'Raumliste_Demo.xlsx' # Wenn die Datei im gleichen Ordner ist
excel_file_path = r'D:\ime\OneDrive - B+S AG\Desktop\Kurs heute\4.3 - Programming\Beispiel Python\Raumliste_Demo.xlsx'


### Datei in Pandas-Modul einlesen

base_df = pd.read_excel(excel_file_path, sheet_name='Sheet1')
#print(base_df)



### Iteration über die einzelnen Reihen, Option in Klammer Spaltenname eingeben

for index,row in base_df.iterrows():
    act_code = row['Raumcode']
    act_roomnumber = row['Raumnummer']
    space_id = f'{act_code}---{act_roomnumber}'
    #print(space_id)


### filtern nach allen zeilen mit HNF in Spalte SIA416 

hnf_df = base_df.loc[base_df['SIA416'] == 'HNF']
#print(hnf_df)


### Summe der Raumfläche von allen HNF

room_area_hnf = hnf_df['Raumflaeche [m²]'].sum()
#print(f'{room_area_hnf} m2')


### Alles ausser die genannten Spalten entfernen

filt_df = base_df.loc[ : ,['SIA416', 'Bodenmaterial']]
#print(filt_df)


