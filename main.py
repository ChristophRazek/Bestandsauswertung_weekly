import pandas as pd
import Act_Inventory
import Old_Inventory
import Email
import logging
from datetime import date


today = date.today()

# Logging konfigurieren
logging.basicConfig(
    filename=r'L:\Datenaustausch\Log\bestaende_script.log',  # Pfad zur Logdatei
    level=logging.INFO,                       # Log-Level: INFO, DEBUG, ERROR, ...
    format='%(asctime)s - %(levelname)s - %(message)s'  # Format der Logzeilen
    )

try:

    logging.info("Script gestartet.")
    logging.info(f"Aktuelles Datum: {date.today()}")


    Act_Inventory.act_inventory()
    #df_gesamt = Old_Inventory.old_invent()

    def create_pivot(df, name):
        df_pivot = pd.pivot_table(df, values=f'{name}', index=['DATUM'], columns=['LAGER'], aggfunc='sum',
                                      margins=True).fillna(0)

        df_pivot = df_pivot.round(decimals=0)
        df_pivot.drop(df_pivot.index[-1], inplace=True)

        #df_pivot.to_excel(rf"S:\EMEA\Lagerbestand\{name}.xlsx", index=True)
        df_pivot.to_excel(rf"L:\Datenaustausch\Bestaende\{name}.xlsx", index=True)

        logging.info(rf"Erstelle {name}.")
        return df_pivot



    kek_pivot = create_pivot(Old_Inventory.old_invent(), "Lagerwert_KEK")
    vk_pivot = create_pivot(Old_Inventory.old_invent(), "Lagerwert_VK")



    #Old_Inventory.old_invent().to_excel(rf"S:\EMEA\Lagerbestand\LagerbestandV2.xlsx", index=False)
    Old_Inventory.old_invent().to_excel(rf"L:\Datenaustausch\Bestaende\LagerbestandV2.xlsx", index=False)
    logging.info(rf"Lagerbestandsliste erstellt.")


    Email.send_email()

except Exception as e:
    logging.error(f"Fehler aufgetreten: {e}")


logging.info("Script beendet.")
