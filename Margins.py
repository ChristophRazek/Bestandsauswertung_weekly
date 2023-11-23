import SQL as s
import pandas as pd
import pyodbc
import warnings
import matplotlib.pyplot as plt

warnings.simplefilter('ignore')

connx_string = r'DRIVER={SQL Server}; server=172.19.128.2\emeadb; database=emea_enventa_live; UID=usr_razek; PWD=wB382^%H3INJ'
conx = pyodbc.connect(connx_string)
df = pd.read_sql(s.margins, conx)
df = df.fillna(0)
df['MARGE'] = df['MARGE'].apply(lambda x: x*100)

#df['KW'] = df.apply(lambda x: str(x['FAKTURA_JAHR']) +"/"+ str(x['FAKTURA_KW']), axis=1)
#df.drop(['FAKTURA_JAHR', 'FAKTURA_KW'],axis=1, inplace=True)

df= df.round(decimals=0)

print(df.to_markdown())
df.to_excel(rf"S:\EMEA\Lagerbestand\Umsätze_Liste.xlsx", index=False)



df_pivot = pd.pivot_table(df, values=['UMSATZ', 'MARGE'],index=['FAKTURA_JAHR','FAKTURA_KW'] , columns=['SUCHNAME'], aggfunc='sum').fillna(0)



print(df_pivot.to_markdown())
df_pivot.to_excel (rf"S:\EMEA\Lagerbestand\Umsätze_Pivot.xlsx")

"""plt.figure(figsize=(30, 10))
plt.plot(df['MARGE'])

plt.title(f'MARGE', fontsize=14)
plt.xticks(rotation=45)
#plt.xlabel('Datum')
#plt.ylabel('Lagerwert')
plt.grid()
plt.legend()
#plt.savefig(rf"S:\EMEA\Lagerbestand\Pivot_{name}.png")
plt.show()"""