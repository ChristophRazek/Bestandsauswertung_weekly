import shutil
import win32com.client as win32
from datetime import date
import os
from tkinter import messagebox


def send_email():


    olApp = win32.Dispatch('Outlook.Application')
    olNS = olApp.GetNameSpace('MAPI')
    os.startfile("outlook")

    #receivers = {'michael.pacher@emea-cosmetics.com': 'Michael','stefan.kandler@emea-cosmetics.com': 'Stefan','martin.bergler@emea-cosmetics.com': 'Martin',
                     #'josef.geisberger@emea-cosmetics.com': 'Josef', 'christoph.razek@emea-cosmetics.com': 'Chris', 'rainer.kienle@emea-cosmetics.com': 'Rainer'}
    receivers = ['stefan.kandler@emea-cosmetics.com','christoph.razek@emea-cosmetics.com','martin.bergler@emea-cosmetics.com', 'rainer.kienle@emea-cosmetics.com',
                 'rudolf.swerak@emea-cosmetics.com','josef.geisberger@emea-cosmetics.com']

    today =date.today()


    text = f"""Hallo \n\nAnbei findest du die wöchentliche Bestandsauswertung.\nBei Fragen einfach melden\nMit freundlichen Grüßen\n\nChris"""
    mail_item = olApp.CreateItem(0)
    mail_item.Subject = f'Lagerbestände am: {today}'
    mail_item.BodyFormat = 1
    mail_item.Body = text

    mail_item.Attachments.Add(r'S:\EMEA\Lagerbestand\LagerbestandV2.xlsx')
    mail_item.Attachments.Add(r'S:\EMEA\Lagerbestand\Lagerwert_KEK.xlsx')
    mail_item.Attachments.Add(r'S:\EMEA\Lagerbestand\Lagerwert_VK.xlsx')
    #mail_item.Attachments.Add(r'S:\EMEA\Lagerbestand\Pivot_KEK.png')
    #mail_item.Attachments.Add(r'S:\EMEA\Lagerbestand\Pivot_VK.png')

    mail_item.Sender = 'christoph.razek@emea-cosmetics.com'
    mail_item.To = ";".join(receivers)


    mail_item.Display()
    mail_item.Save()
    #mail_item.Send()
    #shutil.copy(r'S:\EMEA\Lagerbestand\LagerbestandV2.xlsx',fr'S:\EMEA\Lagerbestand\Wochenauswertung\Lagerbestand_{today}.xlsx')






