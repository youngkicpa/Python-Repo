# %%
import pandas as pd

ledger = pd.read_csv(".\\계정별원장.csv")
# %%

# 합계잔액시산표를 만드는 코드
trialBalance = ledger[['계정코드', '계정명', '차변', '대변']].groupby(['계정코드', '계정명']).sum()
trialBalance['잔액'] = trialBalance['차변'] - trialBalance['대변']
# 각 계정코드와 계정명의 빈도를 계산
frequency = ledger.groupby(['계정코드', '계정명']).size().reset_index(name='빈도')

# trialBalance에 빈도 열 추가
trialBalance = trialBalance.merge(frequency, on=['계정코드', '계정명'], how='left')

# %%
# 계정별원장에서 거래처가 엘아이지넥스원인 행을 추출하는 코드
lignexone = ledger[ledger['거래처'].str.contains("엘아이지넥스원")]

# 추출된 항목에 대하여 계정과목별로 합계를 만드는 코드
summarylignexone = lignexone[['계정코드', '계정명', '차변', '대변']].groupby(['계정코드', '계정명']).sum()

# %%

# 적요란이 전기이월인 항목을 추출하는 코드
prioryear = ledger[ledger['적요란'].str.contains('전기이월')]
prioryear.head()
# %%

# 엑셀에 저장하는 코드
trialBalance.to_excel("sample.xlsx", index=True)
# %%
