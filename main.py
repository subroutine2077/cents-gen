import pandas as pd
import openpyxl

# cents.xlsx is pre-downloaded
df = pd.read_excel('cents.xlsx')

# Filtering for columns
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

# Rename columns for readability
filtered.rename(columns={
    'Number 7000': 'Number',
    'Vote #2 Coloration: Verdigris': 'Verdigris',
    'Vote #2\nColoration: Red': 'Red',
    'Vote #2\nColoration: Gold': 'Gold',
    'Vote #2\nColoration: Desert': 'Desert',
    'Vote #2\nColoration: Obsidian': 'Obsidian'
}, inplace=True)

# Prepare for sheets generation
sheets_names = [
    'Verdigris',
    'Red',
    'Gold',
    'Desert',
    'Obsidian'
]

sheets = {}

# Loop generation of all notables
for sheet_name in sheets_names:
    df_tmp = filtered[[
        'Number',
        sheet_name,
        'Inscription ID',
        'Image URL'
    ]]

    # Filter for notables
    notables = df_tmp[df_tmp[sheet_name] == 'Notable']
    # Rename column name
    notables.rename(columns={'Image URL': 'IMAGE'}, inplace=True)

    # Leave only IMAGE column in place
    image_list = notables['IMAGE'].tolist()
    
    # Open new workbook with openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Loop through list and fill column
    for row_num, item in enumerate(image_list, start=1):
        google_sheet_img_formula = f"=IMAGE(item)"
        ws.cell(row=row_num, column=1, value=item)
    
    # Save to file for testing
    # Need to view files, unable to view in VSCode server
    wb.save(f"notables/{sheet_name}.xlsx")

# Save all notables to files
'''
for sheet_name, sheet in sheets.items():
    sheet.to_csv(f"notables/{sheet_name}.csv", index=False)'''