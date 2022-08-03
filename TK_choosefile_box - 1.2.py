import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np
import datetime
from tkinter import messagebox

pd.set_option('display.width', 100)
pd.set_option('precision', 0)
pd.set_option('expand_frame_repr', False)
def selectFile():
    global filepath
    filepath = filedialog.askopenfilename()
    select_path.set(filepath)


def dif():
    file1 = pd.read_excel(filepath,
                          header=1,  # 指定行作为头
                          sheet_name=0,  # 可表示为  sheet_name=‘sheet’具体名字；也可以sheet_name = [0,1]
                          names=['markA', 'sizeA', 'quantityA', 'siteA', 'markB', 'sizeB', 'quantityB', 'siteB'],
                          usecols="A:H",
                          # usecols="A:D,G:H"#选择列
                          )

    try:
        quantityA = list(map(lambda x: int(x), list(file1['quantityA'])))
        quantityB = list(map(lambda x: int(x), list(file1['quantityB'])))

    except:
        print('报错咯 quantity必须是数字(int)栏')
        messagebox.showwarning(title='warning', message='报错咯 quantity必须是数字(int)栏')

    else:
        # print(quantityA)
        # print(quantityB)
        # print(map(lambda x:int(x),list(file1['quantityA'])))
        #     print(file1.shape)
        #     print(file1.loc[0][3])  #行查看
        num_row, num_colu = file1.shape  # 获取当前最大行列
        nan = np.nan
        list_markA = [x for x in (list(file1['markA']))]
        list_markB = [str(x) for x in (list(file1['markB']))]

        # 把NAN转化为字符串 然后remove 注意remove用法
        if 'nan' in list_markA:
            list_markA.remove('nan')
        else:
            pass
        if 'nan' in list_markB:
            list_markB.remove('nan')
        else:
            pass
        # list_markB = list(file1['markB'])
        # 转化为set 然后判断重复项
        set_markA = set(list_markA)
        set_markB = set(list_markB)
        if len(list_markA) == len(set_markA):
            print('A列 无重复项')
        else:
            # print('A列 存在重复项')
            messagebox.showwarning(title='warning',message = 'A列 存在重复项')
            exit()
        if len(list_markB) == len(set_markB):
            print('B列 无重复项')
        else:
            # print('B列 存在重复项')
            messagebox.showwarning(title='warning',message = 'B列 存在重复项')

            exit()

        print('只在A的器件编码{}'.format(str(set_markA - set_markB)))
        print('B中新增的器件编码{}'.format(set_markB - set_markA))
        for mark in (set_markA & set_markB):  # 取交集
            for i in range(num_row):
                # iloc[[行]，[列]]
                if file1.iloc[[i], [0]].values[0][0] == mark:
                    # print('i ={}'.format(i))
                    break
            for j in range(num_row):

                if file1.iloc[[j], [4]].values[0][0] == mark:
                    # print('j = {}'.format(j))
                    print('-------分割线-------------')
                    break
            set1 = set(
                (file1.iloc[[i], [3]].values[0][0].replace(' ', '').replace('\n', '').replace('  ', '')).split(','))
            set2 = set(
                (file1.iloc[[j], [7]].values[0][0].replace(' ', '').replace('\n', '').replace('  ', '')).split(','))
            set_all = set1 & set2
            # print(set_all)
            set1_in = set1 - set2
            set2_in = list(set2 - set1)
            print('{}'.format(mark))
            if int(file1.iloc[[i], [2]].values[0][0]) == len(set1):
                res = 'PASS'
            else:
                res = '数量校验不通过'
                print('数量校验不通过')
            # print('A的个数{}\nB的个数{}'.format(len(set1), len(set2)))
            # print('只在A的元素{}\n只在B的元素{}'.format(sorted(set1_in, key=lambda x: x), sorted(set2_in, key=lambda x: x)))
            print('只在A的元素 \'{}\n只在B的元素 \'{}'.format(#mark,
                                                        ','.join(sorted(set1_in, key=lambda x: x)),
                                                        ','.join(sorted(set2_in, key=lambda x: x))))
            with open('2.txt', 'a', encoding="utf-8") as t1:
                t1.write(
                    '--------{}----------\n{} \n{} \nA的个数{}\nB的个数{}\n只在A的元素 \'{}\n只在B的元素 \'{}\n'.format(
                                                                                                    datetime.datetime.now().strftime("%H:%M:%S"),
                                                                                                   mark,
                                                                                                    res,
                                                                                                   len(set1),
                                                                                                   len(set2),
                                                                                                   ','.join(sorted(set1_in, key=lambda x: x)),
                                                                                                   ','.join(sorted(set2_in, key=lambda x: x))))

        siteA = list(file1['sizeB'])
        # for i in range(len(siteA)):
        #     if siteA[i].isnull:
        #         print('1')

        a = file1.iloc[[0], [7]].values[0][0]
        # print(a)


win = tk.Tk()
win.resizable(True, True)  # 窗口大小可调（长 /宽）
# 获取当前分辨率
screenwidth = win.winfo_screenwidth()
screenheight = win.winfo_screenheight()
# print(type(screenheight))
win.geometry('450x200+{}+{}'.format(int(screenwidth / 3), int(screenheight / 3)))
win.title('file_choose-ver1.2')
win.attributes("-alpha",1)#设置透明度
# canvas = tk.Canvas(win,bg = 'pink')
# canvas.pack()
select_path = tk.StringVar()
but1 = tk.Button(win, text='文件选择', command=lambda: selectFile())
but1.place(x=300, y=50, width=100, height=20)

but2 = tk.Button(win, text='比对', command=lambda: dif())
but2.place(x=300, y=110, width=100, height=20)

but2 = tk.Button(win, text='2')

entry1 = tk.Entry(win, textvariable=select_path)
entry1.place(x=50, y=50, width=240, height=20)  # 大小调节放到这

entry2 = tk.Entry(win, textvariable=None,state = 'disabled')
entry2.place(x=50, y=80, width=240, height=20)  # 大小调节放到这
entry3 = tk.Entry(win, textvariable=None,state = 'disabled')
entry3.place(x=50, y=110, width=240, height=20)  # 大小调节放到这

win.mainloop()
