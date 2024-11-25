

stkl = """SELECT s.[ARTIKELNR], max(a.VK1/a.VKPRO) as 'VK/Stk_Stkl'

  FROM [EMEA_enventa_live45].[dbo].[STUECKLISTE] as s
  left join [EMEA_enventa_live45].[dbo].[ARTIKEL] as a
  on s.BAUGRUPPE = a.ARTIKELNR
  where s.branchkey = 110 and s.bezeichnung like 'PM%'
  group by s.ARTIKELNR"""

sql_lagernr = '''SELECT Distinct[LAGERNR]
      FROM [emea_enventa_live45].[dbo].[LAGERORT] 
      where lagernr not in (101,102,103,110,195,198,199)
      order by LAGERNR'''


margins = """with cte_auftrag as (
  SELECT 
        ap.belegnr, ap.artikelnr, ap.bezeichnung,
		k.SUCHNAME
		,(ap.[VK]/ap.[VKPRO]) as 'VK/Stk'
		,(a.KEK/a.EKPRO) as 'KEK/Stk'
		,ap.MENGE_GELIEFERT
		,(ap.[VK]*ap.MENGE_GELIEFERT/ap.[VKPRO]) as 'UMSATZ'
       ,case when ap.BEZEICHNUNG like 'TS%' then 0
	    else ((ap.[VK]/ap.[VKPRO]) - (a.KEK/a.EKPRO))*ap.[MENGE_GELIEFERT]
	   End as 'DB'
      ,datepart(YEAR,ap.[FAKTURADATUM]) as 'FAKTURA_JAHR'
	  ,datepart(ISO_WEEK,ap.[FAKTURADATUM]) as 'FAKTURA_KW'
    
  FROM [emea_enventa_live45].[dbo].[AUFTRAGSPOS] as ap
  left join [emea_enventa_live45].[dbo].[ARTIKEL] as a on ap.ARTIKELNR = a.ARTIKELNR
  left join [emea_enventa_live45].[dbo].[KUNDEN] as k on ap.KUNDENNR = k.KUNDENNR

  WHERE ap.BranchKey = 110 and ap.STATUS = 4 and ap.BELEGART in ('1','2','3','4') and ap.KUNDENNR like '10%' and FAKTURADATUM > '2022-12-31' and ap.BEZEICHNUNG not like 'TS%'
  
 )

 select SUCHNAME, FAKTURA_JAHR, FAKTURA_KW, sum(UMSATZ) as 'UMSATZ', sum(DB) as 'DB', sum(DB)/sum(UMSATZ) as MARGE from cte_auftrag
 group by SUCHNAME, FAKTURA_JAHR, FAKTURA_KW
 order by FAKTURA_JAHR, FAKTURA_KW"""