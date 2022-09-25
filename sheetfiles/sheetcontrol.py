import pandas as pd

# > Reads
SourceSheet = pd.ExcelFile(rf'sheetfiles\BU KPI.xlsb', engine='pyxlsb') 
KingDF = pd.read_excel(SourceSheet, 'Latam KPI') # vira
EmailSheet = pd.read_excel(rf'sheetfiles\Libro.xlsx')

# > # Columns constraints
KingDF = KingDF[['MV Owner',
                 'Pay Status',
                 'Markview Action',
                 'URN',
                 'Markview Link',
                 'Invoice Number',
                 'Invoice Date',
                 'Invoice Due Date',
                 'Invoice Currency',
                 'Invoice Amount',
                 'Vendor Name',
                 'PO Number',
                 'Description Field']]

