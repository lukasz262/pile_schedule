'''create excel file for import to Revit'''

import pandas as pd

pile_input = pd.read_excel('./data/pile_schedule_input.xlsx', index_col='Mark')
headerList = list(pile_input.columns)

revitExport = pd.read_excel('./data/export_from_revit.xlsx', index_col=0, header=None, names=headerList)

for i in list(revitExport.index):
    if i in list(pile_input.index):
        revitExport.loc[i, headerList[0]] = pile_input.loc[i, headerList[0]]
        revitExport.loc[i, headerList[1]] = pile_input.loc[i, headerList[1]]
        revitExport.loc[i, headerList[2]] = pile_input.loc[i, headerList[2]]
        revitExport.loc[i, headerList[3]] = pile_input.loc[i, headerList[3]]
    else:
        print('Brak danych wejściowych dla {}'.format(i))

revitExport.to_excel('./data/schedule_for_revit.xlsx', header=False)

input('\nNaciśnij enter żeby wyjść')