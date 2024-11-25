import pandas as pd
import glob
import os
import pyodbc
import SQL as s
import warnings

warnings.simplefilter('ignore')

def old_invent():

    # Query Connection
    connx_string = r'DRIVER={SQL Server}; server=172.19.128.2\emeadb; database=emea_enventa_live45; UID=usr_razek; PWD=wB382^%H3INJ'
    conx = pyodbc.connect(connx_string)
    df_stkl = pd.read_sql(s.stkl, conx)
    dfs = {}

    #Reading all old Inventories in Folder
    path = r'S:\EMEA\Lagerbestand\Hist_Bestaende_wochenuebersicht'
    #path = r'S:\EMEA\Lagerbestand\woche'
    excel_files = glob.glob(os.path.join(path,"*.xlsx"))


    for f in excel_files:

        """Creating a Dictionary Key = Date+Lager, Value = DataFrame"""
        df = pd.read_excel(f)
        datum = f.split("'")[1]
        key = f.split("\\")[4].split(".")[0]
        #print(datum)
        #print(key)
        #print(df)

        #Adding the 'Historic' Date to the DataFrame
        df['DATUM'] = datum
        dfs[key] = df


    #Concat the Dataframes of the Dictionary
    dfs_concat = pd.concat(dfs.values(), axis=0)

    dfs_concat['FILTER'] = dfs_concat['AUSLISTER'].str[:5]
    dfs_concat['FILTER'] = dfs_concat['FILTER'].replace(['aktiv', 'XX/XX', 'XX'], ['30/03', '10/03', '10'])
    dfs_concat['FILTER'] = dfs_concat['FILTER'].apply(lambda x: '20' + x + '/30')
    dfs_concat['FILTER'] = pd.to_datetime(dfs_concat['FILTER'], format='%Y/%m/%d')

    # Filter auf aktuelle Artikel
    dfs_concat = dfs_concat[dfs_concat['FILTER'] > dfs_concat['DATUM']]

    # VK der FW bei PM gemäß Stückliste
    df_gesamt = pd.merge(dfs_concat, df_stkl, how='left', on=['ARTIKELNR'])
    df_gesamt['VK'] = df_gesamt.apply(lambda x: x['VK/Stk'] if x['VK/Stk'] > 0 else x['VK/Stk_Stkl'], axis=1)

    # Lagerwert
    df_gesamt['Lagerwert_KEK'] = df_gesamt.apply(lambda x: x['BESTAND'] * x['KEK/Stk'], axis=1)
    df_gesamt['Lagerwert_VK'] = df_gesamt.apply(lambda x: x['BESTAND'] * x['VK'], axis=1)

    #Bereinigen
    df_gesamt.drop(['FILTER', 'LAGERNR', 'VK/Stk', 'KEK/Stk', 'VK/Stk_Stkl', 'VK'], axis=1, inplace=True)

    return df_gesamt

