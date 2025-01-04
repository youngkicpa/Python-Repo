# %%
import pandas as pd
import numpy as np

# KS system의 경우 계정별원장에 코드가 누락되어 있고, 감가상각충당금이 없고, 유형자산이 순액이며,
# 전일이월의 경우 차변과 대변의 잔액이 없고, 부채와 자본도 -로 되어 있지 않다.

filePath = "C:\\DataTest\\분개장_KSsystem_2409.xlsx"
data = pd.read_excel(filePath, sheet_name="Sheet")

# %%
trialBalance = data[["계정과목", "차변", "대변"]].groupby(["계정과목"]).sum()
trialBalance["잔액"] = trialBalance['차변'] - trialBalance['대변']
count = data[["계정과목"]].groupby(["계정과목"]).size().reset_index(name="빈도")
trialBalance = trialBalance.merge(count, on=["계정과목"], how='left')

# %%
trialBalance.to_excel("Trial Balance_KS.xlsx", index=True)
# %%
