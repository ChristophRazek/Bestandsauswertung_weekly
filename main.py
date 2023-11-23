import pandas as pd
import matplotlib.pyplot as plt
import Act_Inventory
import Old_Inventory
import Email

Act_Inventory.act_inventory()
#df_gesamt = Old_Inventory.old_invent()

def create_pivot(df, name):
    df_pivot = pd.pivot_table(df, values=f'{name}', index=['DATUM'], columns=['LAGER'], aggfunc='sum',
                                  margins=True).fillna(0)

    df_pivot = df_pivot.round(decimals=0)
    df_pivot.drop(df_pivot.index[-1], inplace=True)

    df_pivot.to_excel(rf"S:\EMEA\Lagerbestand\{name}.xlsx", index=True)

    return df_pivot

def create_plot(df, name):
    df_pivot_plot = df.apply(lambda x: x / 1000)
    plt.figure(figsize=(30, 10))
    plt.plot(df_pivot_plot[list(df_pivot_plot.columns)], label=list(df_pivot_plot.columns), linewidth=3)

    plt.title(f'Lagerbestände {name} in (Tsd)', fontsize=14)
    plt.xticks(rotation=45)
    plt.xlabel('Datum')
    plt.ylabel('Lagerwert')
    plt.grid()
    plt.legend()
    plt.savefig(rf"S:\EMEA\Lagerbestand\Pivot_{name}.png")
    #plt.show()


kek_pivot = create_pivot(Old_Inventory.old_invent(), "Lagerwert_KEK")
vk_pivot = create_pivot(Old_Inventory.old_invent(), "Lagerwert_VK")

create_plot(kek_pivot, "KEK")
create_plot(vk_pivot, "VK")

Old_Inventory.old_invent().to_excel(rf"S:\EMEA\Lagerbestand\LagerbestandV2.xlsx", index=False)


r"""#Create KEK Pivot
df_pivot_kek = pd.pivot_table(df_gesamt, values='Lagerwert_KEK', index=['DATUM'], columns=['LAGER'], aggfunc='sum', margins=True).fillna(0)
df_pivot_kek = df_pivot_kek.round(decimals=0)
df_pivot_kek.drop(df_pivot_kek.index[-1], inplace=True)
df_pivot_kek_plot = df_pivot_kek.apply(lambda x: x / 1000)
df_pivot_kek.to_excel(rf"S:\EMEA\Lagerbestand\Pivot_KEK.xlsx", index=True)

#Create VK Pivot
df_pivot_vk = pd.pivot_table(df_gesamt, values='Lagerwert_VK', index=['DATUM'], columns=['LAGER'], aggfunc='sum', margins=True).fillna(0)
df_pivot_vk = df_pivot_vk.round(decimals=0)
df_pivot_vk.drop(df_pivot_vk.index[-1], inplace=True)
df_pivot_vk.to_excel(rf"S:\EMEA\Lagerbestand\Pivot_VK.xlsx", index=True)"""

#print(df_pivot_kek.to_markdown())

#Code2 Fehler auffangen


r"""plt.figure(figsize=(30,10))
plt.plot(df_pivot_kek_plot[list(df_pivot_kek.columns)], label = list(df_pivot_kek.columns), linewidth = 3)
#plt.yticks([500000, 10000000, 1500000, 2000000, 2500000, 3000000, 3500000, 4000000,4500000,5000000, 5500000, 6000000,6500000])

plt.title('Lagerbestände KEK in (Tsd)', fontsize=14)
plt.xticks(rotation = 45)
plt.xlabel('Datum')
plt.ylabel('Lagerwert')
plt.grid()
plt.legend()
plt.savefig(rf"S:\EMEA\Lagerbestand\Pivot_KEK.png")
plt.show()"""

Email.send_email()