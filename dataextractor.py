 
import sqlite3
import pandas as pd
from sqlalchemy import create_engine

file = 'Dragos intrari.xlsx'
output = 'output.xlsx'

engine_my = create_engine('sqlite://', echo=False)
df = pd.read_excel(file, sheet_name='Foaie1')
df1 = df.iloc[:,0:4]
january_1 = []
january_6 = []
january_7 = []
january_25 = []
january_30 = []
febrary_14 = []
febrary_24 = []
march_9 = []
march_25 = []
april_12 = []
easter = []
april_23 = []
may_21 = []
whitsuntide = []
june_24 = []
june_29 = []
july_20 = []
august_6 = []
august_15 = []
august_26 = []
september_8 = []
octomber_14 = []
octomber_26 = []
november_8 = []
november_30 = []
december_6 = []
december_25 = []
december_27 = []



for i in df['In ce zile de sărbătoare oferiți flori?']:
    if type(i) == str and '01.Ianuarie - Sf. Vasile' in i:
        january_1.append('y')
    else:
        january_1.append('')
    if type(i) == str and '06.Ianuarie - Boboteaza' in i:
        january_6.append('y')
    else:
        january_6.append('')
    if type(i) == str and '07.Ianuarie - Sf. Ioan' in i:
        january_7.append('y')
    else:
        january_7.append('')
    if type(i) == str and '25.Ianuarie - Sf. Grigore' in i:
        january_25.append('y')
    else:
        january_25.append('')
    if type(i) == str and '30.Ianuarie - 3 Ierarhi Vasile, Grigore și Ioan' in i:
        january_30.append('y')
    else:
        january_30.append('')
    if type(i) == str and "Valentine's Day" in i:
        febrary_14.append('y')
    else:
        febrary_14.append('')
    if type(i) == str and "Dragobete" in i:
        febrary_24.append('y')
    else:
        febrary_24.append('')
    if type(i) == str and '09.Martie - 40 Mucenici' in i:
        march_9.append('y')
    else:
        march_9.append('')
    if type(i) == str and '25.Martie - Bunavestire' in i:
        march_25.append('y')
    else:
        march_25.append('')
    if type(i) == str and '12.Aprilie - Floriile' in i:
        april_12.append('y')
    else:
        april_12.append('')
    if type(i) == str and 'Paște' in i:
        easter.append('y')
    else:
        easter.append('')
    if type(i) == str and '23.Aprilie - Sf. Gheorghe' in i:
        april_23.append('y')
    else:
        april_23.append('')
    if type(i) == str and 'Sf. Constantin și Elena' in i:
        may_21.append('y')
    else:
        may_21.append('')
    if type(i) == str and 'Rusalii' in i:
        whitsuntide.append('y')
    else:
        whitsuntide.append('')
    if type(i) == str and '24.Iunie - Drăgaica' in i:
        june_24.append('y')
    else:
        june_24.append('')
    if type(i) == str and '29.Iunie - Sf. Apostoli Petru și Pavel' in i:
        june_29.append('y')
    else:
        june_29.append('')
    if type(i) == str and '20.Iulie - Sf. Ilie' in i:
        july_20.append('y')
    else:
        july_20.append('')
    if type(i) == str and '06.August - Schimbarea la față' in i:
        august_6.append('y')
    else:
        august_6.append('')
    if type(i) == str and '15.August - Sf. Maria - Adormirea Maicii Domnului' in i:
        august_15.append('y')
    else:
        august_15.append('')
    if type(i) == str and '26.August - Sf. Mucenici Adrian și Natalia' in i:
        august_26.append('y')
    else:
        august_26.append('')
    if type(i) == str and '08.Septembrie - Sf. Maria - Nașterea Maicii Domnului' in i:
        september_8.append('y')
    else:
        september_8.append('')
    if type(i) == str and '14.Octombrie - Sf. Parascheva' in i:
        octomber_14.append('y')
    else:
        octomber_14.append('')
    if type(i) == str and '26.Octombrie - Sf.Mucenic Dimitrie' in i:
        octomber_26.append('y')
    else:
        octomber_26.append('')
    if type(i) == str and '8.Noiembrie - Sf. Arhangheli Mihai și Gavril' in i:
        november_8.append('y')
    else:
        november_8.append('')
    if type(i) == str and '30.Noiembrie - Sf. Andrei' in i:
        november_30.append('y')
    else:
        november_30.append('')
    if type(i) == str and '06.Decembrie - Sf. Nicolae' in i:
        december_6.append('y')
    else:
        december_6.append('')
    if type(i) == str and 'Crăciunul' in i:
        december_25.append('y')
    else:
        december_25.append('')
    if type(i) == str and 'Sf. Stefan' in i:
        december_27.append('y')
    else:
        december_27.append('')

df1['1 ianuarie'] = january_1
df1['6 ianuarie'] = january_6
df1['7 ianuarie'] = january_7
df1['25 ianuarie'] = january_25
df1['30 ianuarie'] = january_30
df1['14 februarie'] = febrary_14
df1['24 februarie'] = febrary_24
df1['9 martie'] = march_9
df1['25 martie'] = march_25
df1['12 aprilie'] = april_12
df1['Paste'] = easter
df1['23 aprilie'] = april_23
df1['21 mai'] = may_21
df1['Rusalii'] = whitsuntide
df1['24 iunie'] = june_24
df1['29 iunie'] = june_29
df1['20 iulie'] = july_20
df1['6 august'] = august_6
df1['15 august'] = august_15
df1['26 august'] = august_26
df1['8 septembrie'] = september_8
df1['14 octombrie'] = octomber_14
df1['26 octombrie'] = octomber_26
df1['8 noiembrie'] = november_8
df1['30 noiembrie'] = november_30
df1['6 decembrie'] = december_6
df1['25 decembrie'] = december_25
df1['27 decembrie'] = december_27
df1['Cadou pentru zi de nastere sotie (zi, luna)'] = df['Cadou pentru zi de nastere sotie (zi, luna)'].values
df1['Cadou pentru zi de nastere mama (zi, luna)'] = df['Cadou pentru zi de nastere mama (zi, luna)'].values
df1['Cadou pentru zi de nastere sora (zi, luna)'] = df['Cadou pentru zi de nastere sora (zi, luna)'].values
df1['Cadou pentru zi de nastere cumnata (zi, luna)'] = df['Cadou pentru zi de nastere cumnata (zi, luna)'].values
df1['Cadou pentru zi de nastere soacra (zi, luna)'] = df['Cadou pentru zi de nastere soacra (zi, luna)'].values
df1['Cadou pentru aniversare casatorie (zi, luna)'] = df['Cadou pentru aniversare casatorie (zi, luna)'].values
df1['Cadou pentru aniversare casatorie (zi, luna)'] = df['Cadou pentru aniversare casatorie (zi, luna)'].values

day_1_march = []
day_8_march = []
day_school_start = []


for i in df['Zile Speciale']:
    if type(i) == str and '1 Martie' in i:
        day_1_march.append('y')
    else:
        day_1_march.append('')
    if type(i) == str and '8 Martie' in i:
        day_8_march.append('y')
    else:
        day_8_march.append('')
    if type(i) == str and 'Inceperea Școlii' in i:
        day_school_start.append('y')
    else:
        day_school_start.append('')

df1['Zi speciala : 1 Martie'] = day_1_march
df1['Zi speciala : 8 Martie'] = day_8_march
df1['Zi speciala : Inceperea Scolii'] = day_school_start
# df1
# df1
df1.to_sql('datainput', engine_my, if_exists='replace', index=False)

results = engine_my.execute('Select * from datainput')

final=pd.DataFrame(results, columns=df1.columns)

writer = pd.ExcelWriter(output,
                        engine='xlsxwriter',
                        datetime_format='dd/mm/yy hh:mm:ss',
                        date_format='dd/mm/yy')

final.to_excel(writer, sheet_name='Sheet1', index=False)

workbook  = writer.book
worksheet = writer.sheets['Sheet1']
# worksheet.set_column('A:A', 20, 'dd/mm/yy hh:mm:ss')


writer.save()


# final
# result  = engine.query.with_entities(engine.col1)
# df1