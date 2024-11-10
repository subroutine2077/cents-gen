import pandas as pd

df = pd.read_excel('cents.xlsx')

columns_of_interest = [
    'Number 7000',
    'Vote #2 Coloration: Verdigris',
    'Vote #2\nColoration: Red',
    'Vote #2\nColoration: Gold',
    'Vote #2\nColoration: Desert',
    'Vote #2\nColoration: Obsidian',
    'Inscription ID',
    'Image URL'
]

filtered = df[columns_of_interest]

filtered.rename(columns={
    'Number 7000': 'Number',
    'Vote #2 Coloration: Verdigris': 'Verdigris',
    'Vote #2\nColoration: Red': 'Red',
    'Vote #2\nColoration: Gold': 'Gold',
    'Vote #2\nColoration: Desert': 'Desert',
    'Vote #2\nColoration: Obsidian': 'Obsidian'
}, inplace=True)

sheets_names = [
    'Verdigris',
    'Red',
    'Gold',
    'Desert',
    'Obsidian'
]

sheets = {}

for sheet_name in sheets_names:
    df_tmp = filtered[[
        'Number',
        sheet_name,
        'Inscription ID',
        'Image URL'
    ]]

    df_tmp = df_tmp[df_tmp[sheet_name] == 'Notable']

    # IMAGE URL display not working
    '''
    9988    =IMAGE(0       https://rutherfordchang.com/cen...
    9990    =IMAGE(0       https://rutherfordchang.com/cen...
    9991    =IMAGE(0       https://rutherfordchang.com/cen...
    9993    =IMAGE(0       https://rutherfordchang.com/cen...
    9997    =IMAGE(0       https://rutherfordchang.com/cen...
    '''
    df_tmp['Image URL'] = "=IMAGE({0})".format(df_tmp['Image URL'])
    df_tmp = df_tmp.rename(columns={'Image URL': 'IMAGE'})
    sheets[sheet_name] = df_tmp['IMAGE']

    print(sheets[sheet_name].tail())

for sheet_name, sheet in sheets.items():
    sheet.to_csv(f"{sheet_name}.csv", index=False)