import pandas as pd
from openpyxl import Workbook
from openpyxl import load_workbook
#from MP_final_Funktionen import save_excel
pd.options.mode.chained_assignment = None

def mehrfachbestimmung(excelpath, new_poly=None):
    if new_poly is None:
        polymere = ['PE', 'PP', 'PS', 'PMMA', 'PET']
    else:
        polymere = ['PE', 'PP', 'PS', 'PMMA', 'PET', new_poly]
    spannweiten = []

    for polys in polymere:
        df_spannweite = 'df_spannweite_' + polys
        df_spannweite = pd.read_excel(excelpath, sheet_name=polys)
        df_spannweite_raw = df_spannweite[~df_spannweite['Experiment-ID'].str.contains('Methode', na=False)]
        df_spannweite_raw['Proben-ID'] = df_spannweite_raw.apply(lambda x: x['Proben-ID'][6:-2], axis=1)

        # Bestimmung Anzahl an Mehrfachbestimmungen
        anzahl_mehrfachbest = df_spannweite_raw.groupby('Proben-ID').size().reset_index(name='counts')

        # Bestimmung Spannweite aus Maximalwert- Minimalwert, Mittelwert und Standardabweichung
        spannweite = df_spannweite_raw.groupby('Proben-ID')[['Verhältnis [µg/mg]']].agg(
            {'mean', lambda x: x.max() - x.min(), 'std'}).reset_index()
        spannweite['Anzahl Mehrfachbest.'] = anzahl_mehrfachbest['counts']

        # Entfernung aller Messungen, die nur einmal gemessen wurden
        spannweite = spannweite[~anzahl_mehrfachbest['counts'].astype('str').str.contains('1')]
        spannweiten.append(spannweite)

    return spannweiten#, poly_dfs

#spannweiten = mehrfachbestimmung(r'Final_Excels\MP_calc_ohneMatrix.xlsx')

def save_excel_mehrfachbestimmung(excelpath, new_poly=None):
    workbook = load_workbook(excelpath)
    writer = pd.ExcelWriter(excelpath, engine='openpyxl')
    writer.book = workbook
    writer.sheets = {ws.title: ws for ws in workbook.worksheets}

    spannweiten[0].to_excel(writer, sheet_name='PE')
    spannweiten[1].to_excel(writer, sheet_name='PP')
    spannweiten[2].to_excel(writer, sheet_name='PS')
    spannweiten[3].to_excel(writer, sheet_name='PMMA')
    spannweiten[4].to_excel(writer, sheet_name='PET')
    if new_poly is not None:
        spannweiten[5].to_excel(writer, sheet_name='PA')

    writer.save()
    writer.close()

    return writer

#save_excel_mehrfachbestimmung(r'Excel_Auswertung\Mehrfachbestimmung_raw.xlsx')