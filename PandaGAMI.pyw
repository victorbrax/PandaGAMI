# > Imports
import os
import pandas as pd

from sheetfiles.sheetcontrol import EmailSheet, KingDF
from utils.mailOutput.mailOut import attachment, localize_recipient, shoot_mail
from utils.manage.organizer import format_header, junk_drop, user_cross
from utils.notification.nalerts import start_working, success_notification

# > Junk drop
KingDF = junk_drop(KingDF)

# > Usernames capture
usernamesSource = pd.unique(EmailSheet['Full Name'])
usernamesOnHold = pd.unique(KingDF['MV Owner'])

# > Username data crossover
usernames = user_cross(usernamesOnHold=usernamesOnHold, usernamesSource=usernamesSource)

# > Alert notification
start_working()

# > Send engine
for i in range(len(usernames)):

    # > Recipient Attachment set
    IndividualAttachment = KingDF.loc[KingDF['MV Owner'] == usernames[i]]

    # > Recipient mail by name 
    recipient = localize_recipient(EmailSheet, usernames, i)

    # > Attachment username column drop
    IndividualAttachment = IndividualAttachment.drop(columns=['MV Owner'])

    # > Sheet Attachment filename & path + save
    namefile, path = attachment(usernames, IndividualAttachment, i)

    # > Header format:
    format_header(path)

    # > Shoot the email
    shoot_mail(True, usernames[i], recipient, namefile)
    
    # > Remove attachment file to avoid garbage
    os.remove(path)

qtusers = len(usernames)
success_notification(qtusers)
