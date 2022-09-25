import os
import win32com.client as win32
from time import sleep


def shoot_mail(trigger: bool, nameuser: str, mailuser: str, namefile: str):
    
    # > Chamada do Outlook
    olApp = win32.Dispatch('Outlook.Application')
    mailItem = olApp.CreateItem(0) #= [0 parameter] -> New Message+

    # > Corpo de mensagem
    CorpoEmail = open(fr'utils\mailOutput\body.html').read().format(colaborador=nameuser)
    mailItem.HTMLBody = CorpoEmail

    # > Escopo de Formatação (opcional)
    # mailItem.BodyFormat = 1 #= [1 parameter] permite plain text
    # mailItem.BodyFormat = 2 #= [2 parameter] permite utilizar o HTML

    # > Dados do E-mail
    mailItem.Subject = f'Su acción es requerida - Markview - {nameuser}' #= Title
    mailItem.To = mailuser #= Recipient
    mailItem.Attachments.Add(os.path.join(os.getcwd(), fr'sheetoutput/{namefile}'))

    if trigger:
        return mailItem.Send()  
        
    #* Triggers
    # mailItem.Display()
    # mailItem.Save()
    sleep(0.3)