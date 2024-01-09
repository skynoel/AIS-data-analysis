import numpy as np
import openpyxl
import pandas as pd
import os
import matplotlib.pyplot as plt

os.system("cls")

pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.width', 180) 

df=pd.read_excel("Decodered LOG (2023-06-03) 001.xlsx")

#整理資料，計算每分鐘封包數
df2=df.pivot_table(index=["RMC_TAKEN_AT_HOUR","RMC_TAKEN_AT_MIN"],values=["PACKET_TYPE"],aggfunc={"RMC_TAKEN_AT_HOUR":np.size})
df2 = df2.rename(columns={'RMC_TAKEN_AT_HOUR': 'times'})
#計算小時中每分鐘平均封包數
df3=df2.groupby('RMC_TAKEN_AT_HOUR').mean()
#匯出整理後資料
df3.to_excel("Decodered LOG (2023-06-03) 001.xlsx"+"data.xlsx")

#引入整理後資料，並生成統計圖
data=openpyxl.load_workbook("Decodered LOG (2023-06-03) 001.xlsx"+"data.xlsx")
sheet=data.worksheets[0]
sheet.delete_rows(idx=1)
x=[]
y=[]
for i in sheet["A"]:
    A=i.value
    x.append(A)
for i in sheet["B"]:
    B=i.value
    y.append(B)
plt.plot(x, y, marker='.')
plt.xlabel('小時')
plt.ylabel('每分鐘封包數')
plt.title('每日個小時頻道每分鐘封包數平均值')
plt.show()