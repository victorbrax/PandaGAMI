import pandas as pd

# > Reads
SourceSheet = pd.ExcelFile(rf'sheetfiles\BU KPI.xlsb', engine='pyxlsb') 
SourceDF = pd.read_excel(SourceSheet, 'Latam KPI') # vira
EmailSheet = pd.read_excel(rf'sheetfiles\Libro.xlsx')