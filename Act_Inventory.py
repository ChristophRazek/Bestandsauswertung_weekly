import pandas as pd
import pyodbc
import SQL as s
from datetime import date, datetime
import warnings

warnings.simplefilter('ignore')

def act_inventory():

    lagernr = []

    today = date.today()
    now = datetime.now()

    # Query Connection
    connx_string = r'DRIVER={SQL Server}; server=172.19.128.2\emeadb; database=emea_enventa_live; UID=usr_razek; PWD=wB382^%H3INJ'
    conx = pyodbc.connect(connx_string)

    #Ermittlung der betreffenden Läger
    df_lagernr = pd.read_sql(s.sql_lagernr, conx)

    for index, row in df_lagernr.iterrows():
        lagernr.append((row['LAGERNR']))

    for l in lagernr:
        sql_history = rf'''with cte_artikel as (select artikelnr,bezeichnung, code2 as auslister, vk1/VKPRO as 'VK/Stk', kek/EKPRO as 'KEK/Stk' from artikel),
        cte_lager as (SELECT Distinct[LAGERNR], BEZEICHNUNG FROM [emea_enventa_live].[dbo].[LAGERORT] 
        where lagernr not in (101,102,103,110,195,198,199))
    
        select l.ARTIKELNR, Max(cte_artikel.BEZEICHNUNG) as 'BEZEICHNUNG', SUM(menge) as 'BESTAND', Max(l.LAGERNR) as 'LAGERNR', Max(cte_lager.BEZEICHNUNG) as 'LAGER', 
        Max(auslister) as 'AUSLISTER', Max([VK/Stk]) as 'VK/Stk', Max([KEK/Stk]) as 'KEK/Stk'  from LAGERJOURNAL as l
        left join cte_artikel on l.ARTIKELNR = cte_artikel.ARTIKELNR
        left join cte_lager on l.LAGERNR = cte_lager.LAGERNR
        where branchkey = 0110 
        and DATUM <= '{today}'
        and l.LAGERNR = {l}
        group by l.ARTIKELNR, BranchKey
        having sum(menge) > 0'''

        # DataFrame erstellung aus SQL
        df = pd.read_sql(sql_history, conx)

        # Bei keinem Bestand -> DF = "leer", überspringen
        if df.shape[0] == 0:
            continue
        else:

            df.to_excel(rf"S:\EMEA\Lagerbestand\Hist_Bestaende_wochenuebersicht\'{today}'_{l}.xlsx", index=False)
            #print(df_gesamt.to_markdown())


