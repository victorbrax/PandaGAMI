#%% > Imports
import os

import pandas as pd

# Esse terá interface
print('Prepare-se')

from utils.mailOutput.mailOut import shoot_mail
from utils.manage.organizer import format_header, junk_drop, user_cross
from utils.notification.nalerts import start_working, success_notification

# > Reads
SourceSheet = pd.ExcelFile(rf'sheetfiles\BU KPI.xlsb', engine='pyxlsb')
SourceDF = pd.read_excel(SourceSheet, 'Latam KPI')
EmailSheet = pd.read_excel(rf'sheetfiles\Libro.xlsx')

# > Junk drop
SourceDF = junk_drop(SourceDF)

# > New Name
KingDF1 = SourceDF

# > # Columns constraints
KingDF1 = KingDF1[['MV Owner',
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

# > Usernames capture
usernamesSource = pd.unique(EmailSheet['Full Name'])
usernamesOnHold = pd.unique(KingDF1['MV Owner'])

# > Username data crossover
usernames = user_cross(usernamesOnHold=usernamesOnHold, usernamesSource=usernamesSource)

# > Alert notification
start_working()

# > Send engine
for i in range(len(usernames)):
    # > Recipient Attachment set
    IndividualAttachment = KingDF1.loc[KingDF1['MV Owner'] == usernames[i]]

    # > Recipient mail by name
    recipient = EmailSheet.loc[EmailSheet['Full Name'] == usernames[i]]
    recipient = recipient[['Sencinet Email']]
    recipient = recipient.to_string(index=False, header=False)

    # > Attachment username column drop
    IndividualAttachment = IndividualAttachment.drop(columns=['MV Owner'])
    # - ↑ Para ratificar correspondência, comente ↑

    # > Sheet Attachment filename & path + save
    namefile = f'Items pendientes Markview -'+usernames[i]+'.xlsx'
    path = f'sheetoutput/{namefile}'
    IndividualAttachment.to_excel(path, index=False)

    # > Header format:
    format_header(path)

    # > Shoot the email
    shoot_mail(True, usernames[i], recipient, namefile)
    
    # > Remove attachment file to avoid garbage
    os.remove(path)
    # - ↑ Para ratificar correspondência, comente ↑


qtusers = len(usernames)
success_notification(qtusers)

# * Generate .exe:
# pyinstaller -F --onefile main.pyw
