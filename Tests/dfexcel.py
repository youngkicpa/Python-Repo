# %%
import pandas as pd
import numpy as np

filePath = "C:\\DataTest\\계정별원장_정리_인텔릭스_FY23.xlsx"
data = pd.read_excel(filePath, sheet_name="계정별원장")

# %%
trialBalance = data[["계정코드", "계정명", "차변", "대변"]].groupby(["계정코드", "계정명"]).sum()
trialBalance["잔액"] = trialBalance['차변'] - trialBalance['대변']
count = data[["계정코드", "계정명"]].groupby(["계정코드", "계정명"]).size().reset_index(name="빈도")
trialBalance = trialBalance.merge(count, on=['계정코드', "계정명"], how='left')


# %%
