import shutil
import win32com.client as win32
from datetime import date
import os
from tkinter import messagebox
import logging

# Logging konfigurieren
logging.basicConfig(
    filename=r'L:\Datenaustausch\Log\bestaende_script.log',  # Pfad zur Logdatei
    level=logging.INFO,                       # Log-Level: INFO, DEBUG, ERROR, ...
    format='%(asctime)s - %(levelname)s - %(message)s'  # Format der Logzeilen
    )

def send_email():

    try:

        logging.info(rf"Starte Emailversand.")

        olApp = win32.Dispatch('Outlook.Application')
        olNS = olApp.GetNameSpace('MAPI')
        #os.startfile("outlook")


        receivers = ['stefan.kandler@emea-cosmetics.com','christoph.razek@emea-cosmetics.com','martin.bergler@emea-cosmetics.com', 'rainer.kienle@emea-cosmetics.com',
                     'anton.usanov@emea-cosmetics.com','daniel.kovacs@emea-cosmetics.com','josef.geisberger@emea-cosmetics.com']


        today =date.today()


        text = f"""Hallo \n\nAnbei findest du die wöchentliche Bestandsauswertung.\nBei Fragen einfach melden\nMit freundlichen Grüßen\n\nChris"""
        mail_item = olApp.CreateItem(0)
        mail_item.Subject = f'Lagerbestände am: {today}'
        mail_item.BodyFormat = 1
        mail_item.Body = text

        #mail_item.Attachments.Add(r'S:\EMEA\Lagerbestand\LagerbestandV2.xlsx')
        #mail_item.Attachments.Add(r'S:\EMEA\Lagerbestand\Lagerwert_KEK.xlsx')
        #mail_item.Attachments.Add(r'S:\EMEA\Lagerbestand\Lagerwert_VK.xlsx')

        mail_item.Attachments.Add(rf"L:\Datenaustausch\Bestaende\LagerbestandV2.xlsx")
        mail_item.Attachments.Add(rf"L:\Datenaustausch\Bestaende\Lagerwert_KEK.xlsx")
        mail_item.Attachments.Add(rf"L:\Datenaustausch\Bestaende\Lagerwert_VK.xlsx")


        #mail_item.Sender = 'christoph.razek@emea-cosmetics.com'
        mail_item.To = ";".join(receivers)


        mail_item.Display()
        mail_item.Save()
        mail_item.Send()

        logging.info(rf"Emails versendet.")

    except Exception as e:
        logging.error(f"Fehler aufgetreten: {e}")





