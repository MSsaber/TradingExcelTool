#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   trader_excel_view.py
@Time    :   2023/02/07 12:23:00
@Author  :   xiaobai
@Version :   1.0
@Contact :   1752615737@qq.com
@Desc    :   报表核对界面
'''

from enum import Enum
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import tkinter.messagebox as msgbox

import trading_excel

def select_file(open: bool = False, action = None):
    filepath = filedialog.askopenfilename()
    if open is False:
        return filepath
    elif open and action is not None:
        return action(filepath)

def select_folder():
    f = filedialog.askdirectory()
    return f

def check_filefomat(file):
    file_s = file.split('.')
    if file_s[len(file_s) - 1] != 'xls' and \
        file_s[len(file_s) - 1] != 'xlsx':
        return False
    return True

class TraderExcelMerge:
    def __init__(self):
        self.fileptah = None
        self.root = tk.Tk(sync=True)
        self.root.title("表格合并")
        self.root.geometry("300x150")
        self.root.resizable(0,0)
        #self.txtfile = tk.StringVar()
        self.text = ttk.Entry(self.root, width=25)
        self.text.place(x=10, y = 30)
        
        self.selectBn = ttk.Button(self.root, text="表格导入", width=10, command=self.select_click,)
        self.selectBn.place(x=10 + 30 + 170, y=30)

        self.textexp = ttk.Entry(self.root,width=25)
        self.textexp.place(x=10, y = 30 +10 + 30)

        self.sexpBn = ttk.Button(self.root, text="表格合并", width=10, command=self.export_file)
        self.sexpBn.place(x=10 + 30 + 170, y=30 + 10 + 30)
        self.root.mainloop()

    def select_click(self):
        self.fileptah = select_file()
        self.text.delete(0,tk.END)
        self.textexp.delete(0,tk.END)
        if self.fileptah is not None:
            filename = self.fileptah.split('/')
            self.resfilename = filename[len(filename)-1]
            if check_filefomat(self.resfilename) is False:
                msgbox.showinfo('警告', '文件类型错误！')
                return
            fn = self.resfilename.split('.')
            self.resfilename = fn[0] + '_汇总表.' + fn[len(fn) - 1]
            self.text.insert(0,filename[len(filename)-1])
            self.textexp.insert(0, self.resfilename)
        self.root.focus()

    def export_file(self):
        path = select_folder()
        self.resfilename = self.textexp.get()
        self.expfile = path + '/' + self.resfilename
        trading_excel.merge_excel(self.fileptah ,self.expfile)
        print(self.expfile)

class ExcelToolWnd:
    class ExcelMode(Enum):
        trader_merge = '交易表合并'
        invalid = '暂无'

    def __init__(self):
        self._init_view()

    def button_click(self):
        smode = self.cbMode.get()
        if smode == ExcelToolWnd.ExcelMode.trader_merge.value:
            TraderExcelMerge()

    def _init_view(self):
        self.rootWnd = tk.Tk()
        self.rootWnd.title("表格自动化工具")
        self.rootWnd.geometry("300x150")
        self.rootWnd.resizable(0, 0)

        self.cbMode = ttk.Combobox(self.rootWnd, justify=tk.RIGHT, values=('交易表合并', '暂无'))

        self.activeBn = ttk.Button(self.rootWnd, text='确定', command=self.button_click)

        self.cbMode.pack(side=tk.LEFT, padx=20)
        self.activeBn.pack(side=tk.LEFT, padx=10)

        self.rootWnd.mainloop()

if __name__ == "__main__":
    exe = ExcelToolWnd()