import tkinter as tk
from tkinter import filedialog

root = tk.Tk()

root.geometry("1000x650+500+180")
root.title('AIS資料檔案')
root.configure(bg="#D0D0D0")

root.iconbitmap('messageImage_1704871636487.ico')

text = tk.StringVar()
text.set('')

def show():
    file_path = filedialog.askopenfilename()
    text.set(file_path)
def off():
    root.destroy()

btn = tk.Button(root,text='選擇檔案',font=('Arial',15,'bold'),command=show)
btn.place(x=705,y=100,height=50,width=100)

E1=tk.Entry(root,width=40,textvariable=text)
E1.place(x=200,y=100,height=50,width=500)
btn2 = tk.Button(root,text="取消",command=off,font=('Arial',15,'bold'))
btn2.place(x=850,y=550,height=50,width=100,)
btn3 = tk.Button(root,text="確認",command=off,font=('Arial',15,'bold'))
btn3.place(x=700,y=550,height=50,width=100,)
mylabel = tk.Label(root, text="檔案：", font=('Arial',15,'bold'),bg="#D0D0D0")
mylabel.place(x=100,y=100,height=50,width=70)
root.mainloop()


import AIS.mad  
file=text   # 讀取欲進行處裡的檔案
file2=AIS.mad.avperm(file)  #將經過時間碼標記的原始資料轉變為個小時每分鐘平均封包數的資料
AIS.mad.graph(file2)    #利用該資料繪製成MAD

import AIS.patd  
file3=AIS.patd.chose(file2)#在MAD資料中以特定算法標記出封包數較多的時間
AIS.patd.graph(file3)#利用該資料繪製成PATD表格

import AIS.chr_patd   
file4=AIS.chr_patd.search(file,file3)#根據PATD資料從原始資料中尋找出船舶MMSI碼
file5=AIS.chr_patd.dynamic(file,file4)#利用MMSI碼提取原始資料中的船舶動態資訊並篩選出船種為漁船且船速低於3節的


import AIS.asd
file6=AIS.asd.gis(file5)#將資料送入GIS系統中進行分析，並統計出不同區域之信號密度
AIS.asd.graph(file6)#根據資料繪製成ASD圖表

import AIS.pfg
file7=AIS.pfg.filter(file6)#將資料中地理位置屬於在港狀態的封包數剔除
AIS.pfg.graph(file7)#將最終資料產生PFG圖用作魚場區域分析
