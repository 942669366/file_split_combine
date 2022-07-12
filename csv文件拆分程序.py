import tkinter as tk
import csv
import datetime
import os
from tkinter import filedialog
from tkinter.messagebox import *

window = tk.Tk()  # 实例化Tk
window.title("csv文件拆分")  # 设置标题
window.geometry('500x300')  # 设置窗口的大小
window.config(background = "#d6d6d6")


lie1 = 20
lie2 = 180

tk.Label(window, text="请输入单份文件拆分数量：",bg = "#d6d6d6").place(x=lie1, y=40)
chaifen_cout = tk.StringVar()  # 文件输入路径变量



# tk.Label(window, text="数据总数量" + "sss").place(x=385, y=1)
# tishi_cout = tk.StringVar()  # 文件输入路径变量


tk.Label(window, text="请选择需要拆分的文件：",bg = "#d6d6d6").place(x=lie1, y=90)
var_name = tk.StringVar()  # 文件输入路径变量

tk.Label(window, text="请选择拆分文件后存储位置：",bg = "#d6d6d6").place(x=lie1, y=140)
var_name2 = tk.StringVar()  # 文件夹输入路径变量

entry_name = tk.Entry(window, textvariable=chaifen_cout, width=15)
entry_name.place(x=lie2, y=40)

entry_name = tk.Entry(window, textvariable=var_name, width=30)
entry_name.place(x=lie2, y=90)

entry_name2 = tk.Entry(window, textvariable=var_name2, width=30)
entry_name2.place(x=lie2, y=140)
# conut = chaifen_cout.get()  # 得到文本输入框的值
#     files = var_name.get()
#     files_dir = var_name2.get()
#     start = datetime.datetime.now()



# 输入文件路径
def selectPath_file():
    path_ = filedialog.askopenfilename(filetypes=[("数据表", [".xlsx"])])
    var_name.set(path_)

# 输入文件夹路径
def selectPath_dir():
    path_t = filedialog.askdirectory()
    var_name2.set(path_t)


def yunxing():

    conut = chaifen_cout.get()  # 得到文本输入框的值
    files = var_name.get()
    files_dir = var_name2.get()
    start = datetime.datetime.now()
    print(files)
    print(files_dir)
    print(conut)
    couts = 0
    with open(files, 'rt', newline='', encoding='utf-8', errors='ignore') as file:
        csvreader = csv.reader(file)
        i = j = 1
        for row in csvreader:
            couts += 1
            print(row)
            # 没1000个就j加1， 然后就有一个新的文件名
            if i % int(conut) == 0:
                j += 1
                print(f"csv {j} 生成成功")
            csv_path = os.path.join(files_dir +'/'+ str(j) + '.csv')
            # 不存在此文件的时候，就创建
            if not os.path.exists(csv_path):
                with open(csv_path, 'a', newline='', encoding='utf-8', errors='ignore') as file:
                    csvwriter = csv.writer(file)
                    # csvwriter.writerow(['image_url','编号'])#首行字段
                    csvwriter.writerow(row)
                i += 1
            # 存在的时候就往里面添加
            else:
                with open(csv_path, 'a', newline='', encoding='utf-8', errors='ignore') as file:
                    csvwriter = csv.writer(file)
                    csvwriter.writerow(row)
                i += 1
    end = datetime.datetime.now()
    showinfo('提示','总数据：'+str(couts)+'条'+'\n'+'程序运行时间: '+str(((end - start) / 60).seconds) + '分钟')
    window.quit()


tk.Button(window, text="文件选择", command=selectPath_file,bg = '#d1d1d1').place(x=410, y=85)
tk.Button(window, text="存储路径选择", command=selectPath_dir,bg = '#d1d1d1').place(x=410, y=135)

tk.Button(window, text="运行", command=yunxing,width = 15,height = 2,bg = '#d1d1d1').place(x=200, y=190)


window.mainloop()  # 显示窗口