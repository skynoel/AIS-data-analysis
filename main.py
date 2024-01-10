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



import AIS.chr_patd   



import AIS.asd
#


import AIS.pfg

