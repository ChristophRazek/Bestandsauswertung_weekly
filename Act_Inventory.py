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
            """#Erstellung des Auslister-Filters
            df['FILTER'] = df['AUSLISTER'].str[:5]
            df['FILTER'] = df['FILTER'].replace(['aktiv', 'XX/XX', 'XX'], ['30/03', '10/03', '10'])
            df['FILTER'] = df['FILTER'].apply(lambda x: '20' + x + '/30')
            df['FILTER'] = pd.to_datetime(df['FILTER'], format='%Y/%m/%d')
    
            #Filter auf aktuelle Artikel
            df = df[df['FILTER'] >now]
    
            #VK der FW bei PM gemäß Stückliste
            df_gesamt = pd.merge(df,df_stkl, how='left', on=['ARTIKELNR'])
            df_gesamt['VK'] = df_gesamt.apply(lambda x: x['VK/Stk'] if x['VK/Stk'] > 0 else x['VK/Stk_Stkl'], axis=1)
    
            # Lagerwert
            df_gesamt['Lagerwert_KEK'] = df_gesamt.apply(lambda x: x['BESTAND'] * x['KEK/Stk'], axis=1)
            df_gesamt['Lagerwert_VK'] = df_gesamt.apply(lambda x: x['BESTAND'] * x['VK'], axis=1)
            df_gesamt['DATUM'] = today
    
            df_gesamt.drop(['FILTER', 'LAGERNR', 'VK/Stk', 'KEK/Stk','VK/Stk_Stkl', 'VK', 'DATUM'],axis=1, inplace=True)"""

            df.to_excel(rf"S:\EMEA\Lagerbestand\Hist_Bestaende_wochenuebersicht\'{today}'_{l}.xlsx", index=False)
            #print(df_gesamt.to_markdown())


