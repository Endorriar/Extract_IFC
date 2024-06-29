import pandas as pd

base_file_path = r"C:\Users\MichalRontsinsky\OneDrive - beyondBIM\PERSONAL\PROJEKTE\HSLU\Digital Twin Design and Engineering\Raumliste_Demo.xlsx"


base_df = pd.read_excel(base_file_path, sheet_name="Sheet1")


#BEISPIEL > Einlesen CSV
#base_df = pd.read_csv(base_file_path)

#HNF Filtern > Beispiel DataFrame filtern durch Abgleich Spalte und Wert

hnf_df = base_df.loc[base_df["SIA416"] == "HNF"]
print(hnf_df)


raum_flaeche_hnf = hnf_df["Raumflaeche [m²]"].sum()
print(raum_flaeche_hnf)

#Loc nur Spalten SIA416, Bodenmaterial

filt_df = base_df.loc[ : ,["SIA416", "Bodenmaterial"]]
print(filt_df)

#iLoc mit index, z.B. erst 10 Zeilen von SIA416 Spalte

filt_df_iloc = base_df.iloc[0:10, base_df.columns.get_loc('SIA416')]
print(filt_df_iloc)

filt_df.to_excel("Räume.xlsx",sheet_name="Räume", index=False)





