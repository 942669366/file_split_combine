#实现excel表格按行数分拆的功能
from tkinter.messagebox import showinfo

import openpyxl
import datetime
import time
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
window = tk.Tk()
window.title('歌词信息提取')
window.geometry('500x250')
window.config(background='#d6d6d6')

lie1 = 20
lie2 = 190
# tk.Label(window, text="请输入单个文件分割数量：", bg="#d6d6d6").place(x=lie1, y=30)
# chaifen_cout = tk.StringVar()  # 文件输入路径变量
# tk.Label(window, text="请选择要执行的文件：", bg="#d6d6d6").place(x=lie1, y=70)
# var_name = tk.StringVar()  # 文件输入路径变量
tk.Label(window, text="请选择需要合并的文件夹：",bg = "#d6d6d6").place(x=lie1, y=35)
var_name2 = tk.StringVar()  # 文件夹输入路径变量

# entry_name = tk.Entry(window, textvariable=chaifen_cout, width=20)
# entry_name.place(x=lie2, y=30)
# entry_name = tk.Entry(window, textvariable=var_name, width=30)
# entry_name.place(x=lie2, y=70)
entry_name2 = tk.Entry(window, textvariable=var_name2, width=30)
entry_name2.place(x=lie2, y=35)
# 输入文件夹路径
def selectPath_dir():
    path_t = filedialog.askdirectory()
    var_name2.set(path_t)

def appends(): #path：所有需要合并的excel文件所在的文件夹
    filename_excel = [] # 建立一个空list,用于储存所有需要合并的excel名称
    frames = [] # 建立一个空list,用于储存dataframe
    path = var_name2.get()
    for root, dirs, files in os.walk(path):
        for file in files:
            file_with_path = os.path.join(root, file)
            filename_excel.append(file_with_path)
            df = pd.read_excel(file_with_path, engine='openpyxl')
            frames.append(df)
    df = pd.concat(frames, axis=0)
    df.to_excel(path+"/合并的excel.xlsx")
    showinfo('提示', '合并完成!')
    window.quit()
tk.Button(window, text="存储路径选择", command=selectPath_dir,bg = '#d1d1d1').place(x=410, y=30)
tk.Button(window, text="运行", command=appends,width = 15,height = 2,bg = '#d1d1d1').place(x=200, y=100)
window.mainloop()  # 显示窗口