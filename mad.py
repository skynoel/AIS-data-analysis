import numpy as np
from numpy import char
import openpyxl
import pandas as pd
import os
import matplotlib.pyplot as plt
from tqdm import tqdm

os.system("cls")
folder_path=r"C:\Users\no122\Desktop\AIS-data-analysis\Decodered LOG (2023-06-03)"
items = os.listdir(folder_path)

count = [[0]*65 for t in range(4)]
count[0][0]="RMC_DATE_YEAR"
count[0][1]="RMC_DATE_MON"
count[0][2]="RMC_DATE_DAY"
count[0][3]="RMC_TAKEN_AT_HOUR"
count[0][4]="CHANNEL"
for y in range(5):
    count[1][y]=" "
for y in range(60):
    count[1][y+5]=y
for y in range(60):
    count[0][y+5]=" "
row=2
rowchange=13
r2=30
for item in tqdm(items):
    # 構建完整的路徑
    item_path = os.path.join(folder_path, item)
    if os.path.isfile(item_path):
        df=pd.read_excel(item_path)
        for i in tqdm(range(df.shape[0])):
            data=df.iloc[i]
            if rowchange<data.iloc[3]:
                row=row+2
                count.append([0]*65)
                count.append([0]*65)
                #print(row)
                rowchange=data.iloc[3]
            elif r2<data.iloc[2]:
                row=row+2
                count.append([0]*65)
                count.append([0]*65)
                #print(row)
                r2=data.iloc[2]
            if data.iloc[8] == "A":
                for k in range(0,4):
                    count[row][k]=int(data.iloc[k])
                count[row][4]="A"
                count[row][data.iloc[4]+5]=int(count[row][data.iloc[4]+5])+1
                rowchange=data.iloc[3]
                r2=data.iloc[2]
            elif data.iloc[8] == "B":
                for k in range(0,4):
                    count[row+1][k]=int(data.iloc[k])
                count[row+1][4]="B"
                count[row+1][data.iloc[4]+5]=int(count[row+1][data.iloc[4]+5])+1
                rowchange=data.iloc[3]
                r2=data.iloc[2]
result = pd.DataFrame(count)
name="mad20230603.xlsx"
result.to_excel(name)

