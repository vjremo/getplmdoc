import pandas as pd
xl = pd.ExcelFile('Techpack_RTM.xlsx')
print('SHEETS:', xl.sheet_names)
for s in xl.sheet_names:
    df = pd.read_excel('Techpack_RTM.xlsx', sheet_name=s, header=None)
    print(f'\n=== {s} ({df.shape[0]}x{df.shape[1]}) ===')
    for i, row in df.head(25).iterrows():
        print(i, list(row))
