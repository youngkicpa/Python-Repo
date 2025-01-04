# %%

import openpyxl as xl

fileName = "C:\\Data\\Samples\\인텔릭스.xlsx"

wb = xl.load_workbook(filename = fileName)
for x in wb.sheetnames:
    print(x)

ws = wb['계정별원장Table']

data = []
for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
    data.append(row)

for x in data:
    print(x)

# %%
import openpyxl as xl
import pandas as pd
from itertools import islice

fileName = "C:\\Data\\Samples\\인텔릭스.xlsx"
wb = xl.load_workbook(filename = fileName)
data = wb["계정별원장Table"].values #generator

cols = next(data)[:]
data = list(data)
data = (islice(r, 0, None) for r in data)
df = pd.DataFrame(data, index=None ,columns=cols)

trialBalance = df[['계정코드', '계정과목', '차변', '대변']].groupby(['계정코드', '계정과목']).sum()

print(df['차변'].sum() - df['대변'].sum())

# %%
import openpyxl as xl
import pandas as pd
from itertools import islice

fileName = "C:\\Data\\Samples\\인텔릭스.xlsx"
wb = xl.load_workbook(filename = fileName)
ws = wb["계정별원장Table"]

rows = tuple(ws.iter_rows(values_only=True))

cols = rows[0]
data = rows[1:]

df = pd.DataFrame(data, index=None ,columns=cols)

trialBalance = df[['계정코드', '계정과목', '차변', '대변']].groupby(['계정코드', '계정과목']).sum()

print(df['차변'].sum() - df['대변'].sum())
# %%
