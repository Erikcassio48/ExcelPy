#! python3
# readCensusExcel.py - Tabulates population and number of census tracts for 
# each county.

import openpyxl, pprint
print('Opening workbook...')
wb = openpyxl.load_workbook('c:\Venv\dados.xlsx')
sheet = wb.get_sheet_by_name('Plan1')
dadosregionais = {}

row_count = sheet.max_row
column_count = sheet.max_column
# Fill in countyData with each county's population and tracts.
print('Reading rows...')
for row in range(2, row_count + 1):
    # Each row in the spreadsheet has data for one census tract.
    estados  = sheet['B' + str(row)].value
    produtos = sheet['C' + str(row)].value
    vendas    = sheet['D' + str(row)].value

    # Make sure the key for this state exists.
    dadosregionais.setdefault(estados, {})
    # Make sure the key for this county in this state exists.
    dadosregionais[estados].setdefault(produtos, {'tracts': 0, 'vendas': 0})

    # Each row represents one census tract, so increment by one.
    dadosregionais[estados][produtos]['tracts'] += 1
    # Increase the county pop by the pop in this census tract.
    dadosregionais[estados][produtos]['vendas'] += int(vendas)

# Open a new text file and write the contents of countyData to it.
print('Writing results...')
resultFile = open('census2010.py', 'w')
resultFile.write('allData = ' + pprint.pformat(dadosregionais))
resultFile.close()
print('Done.')
