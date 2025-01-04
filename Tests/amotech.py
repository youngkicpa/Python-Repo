# %%
import pandas as pd
import numpy as np

filePath = "C:\\Data\\Programming\\python\\Tests\\아모텍_24년9월.xlsx"
data = pd.read_excel(filePath, sheet_name="Sheet1")

filtered_data = data[data['계정과목'].str[2:4] == '매출']

monthly_sales = filtered_data[["관리항목1", "회계일", "대변금액"]].groupby([filtered_data['관리항목1'], filtered_data['회계일'].str.slice(5,2)])['대변금액'].sum().unstack(fill_value=0)

monthly_sales.reset_index()

# %%
monthly_sales.to_excel("월별품목별매출액.xlsx", index=True)
# %%
