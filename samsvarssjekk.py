import pandas as pd

# Mangler: sjekk for ikraftdato

# Variabler som er hardkodet inn og som kanskje må endres for andre excel filer:
# Kolonnenavnene i excel, siste rad i bokdel

# Med en excel fil med flere ark så kan hver ark lastes inn individuelt
with pd.ExcelFile('samsvarssjekkTest.xlsx') as xls:
    # Lager dataframes for delene som skal sammenliknes
    bokdel = pd.read_excel(xls, 'merknader ut fra bokdel', usecols=['Arealplan-ID', 'Plannavn', 'Plantype', 'Planstatus', 'IKraft'])
    kartdel1 = pd.read_excel(xls, 'fra kartdel Rp1', usecols=['planidentifikasjon', 'planstatus']).dropna()
    kartdel2 = pd.read_excel(xls, 'fra kartdel Rp2', usecols=['planidentifikasjon', 'planstatus']).dropna()
    kartdel3 = pd.read_excel(xls, 'fra kartdel Rp3', usecols=['planidentifikasjon', 'planstatus']).dropna()
    
# Kombiner kartdelene til én del
kartdel = pd.concat([kartdel1, kartdel2, kartdel3], ignore_index=True)
# Denne linjen fjerner bokstavene fra planIDene som inneholder bokstaver
kartdel['planidentifikasjon'] = kartdel['planidentifikasjon'].apply(lambda x: int(''.join(filter(str.isdigit, str(x)))) if isinstance(x, str) else int(x))
# Den siste raden i bokdel er ikke et datapunkt. Den fjernes
bokdel = bokdel.iloc[:-1]

# Regler for hvordan statuser konverteres til statuser på tallformat
mapping = {
    "Planlegging igangsatt": 1,
    "Planforslag": 2,
    "Endelig vedtatt arealplan": 3
}

# Konverterer statuser i bokdelen til tall som vist ovenfor
# Alle andre statuser får koden 4 siden disse trenger man ikke å skille mellom.
bokdel['statuskode'] = bokdel['Planstatus'].apply(lambda x: mapping.get(x, 4))
# Nye kolonne merknad
bokdel['merknad'] = ''

# Ny kolonne kode. Hver rad i bokdel får en kode hvor første siffer er den faktiske satusen og følgende siffer er status/statuser den har i kartdelen.
# Disse kodene vil inneholde informasjonen som trengs for en samsvarssjekk.
# Feks kode 33 betyr at planIDen har status vedtatt i begge delene så den stemmer. Kode 21 vil bety forslag i bokdel og igangsatt i kartdel. 
# Dette vil skrives til merknad kolonnen.
# Koden vil også plukke opp om en ID er oppført flere ganger i kartdelen. Feks betyr 212 at den finnes både som igangsatt og som forslag i kartdelen.

bokdel['kode'] = 0

# Her sammenliknes bokdel og kartdel og kodene genereres.
for i1, id1 in enumerate(bokdel['Arealplan-ID']):
    kode = ''
    status1 = bokdel['statuskode'][i1]
    if status1==4:
        kode+='4'
        for i2, id2 in enumerate(kartdel['planidentifikasjon']):
            if id1==id2:
                status2 = kartdel['planstatus'][i2]
                kode+=str(int(status2))
        if len(kode)==1:
            kode+=str(0)

    else:
        kode+=str(int(status1))
        for i2, id2 in enumerate(kartdel['planidentifikasjon']):
            if id1==id2:
                status2 = kartdel['planstatus'][i2]
                kode+=str(int(status2))
        if len(kode)==1:
            kode+=str(0)

    bokdel['kode'][i1] = kode

# Her tolkes kodene og skrives som eventuell merknad
for i, kode in enumerate(bokdel['kode']):
    merknad = ''
    if kode=='10' or kode=='20' or kode=='30':
        merknad = 'Ikke i kartdel'
    elif kode == '33' or kode == '22' or kode == '11':
        merknad = 'OK'
    elif kode=='21' or kode=='31' or kode=='41':
        merknad = 'I kartdel som igangsatt'
    elif kode=='12' or kode=='32' or kode=='42':
        merknad = 'I kartdel som forslag'
    elif kode=='13' or kode=='23' or kode=='43':
        merknad = 'I kartdel som vedtatt'
    elif len(kode)>2: 
        merknad = 'Oppført i kartdel flere ganger'

    
    bokdel['merknad'][i] = merknad

# Sjekk for om det er IDer i kartdel som ikke er i bokdel:
for i1, id1 in enumerate(kartdel['planidentifikasjon']):
    if id1 not in bokdel['Arealplan-ID'].tolist():
        bokdel.loc[len(bokdel)] = None
        bokdel.loc[len(bokdel)-1, 'Arealplan-ID'] = id1
        bokdel.loc[len(bokdel)-1, 'Planstatus'] = kartdel['planstatus'][i1]
        

# Select the columns 'col1', 'col2', and 'col3'
selected_columns = bokdel[['Arealplan-ID', 'Plannavn', 'Plantype', 'Planstatus', 'IKraft', 'merknad', 'kode']]

# Write the DataFrame to an Excel file
selected_columns.to_excel('outputsamsvarssjekkTest.xlsx', index=False)