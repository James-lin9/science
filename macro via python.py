import numpy as np #需要的python库
import sys
import win32com
from win32com.client import Dispatch, constants
def useVBA(macro_name,sht,k_ix):#对特定的excel的特定sheet执行特定的宏
    xlApp = win32com.client.DispatchEx("Excel.Application")
    xlApp.Visible = True
    xlApp.DisplayAlerts = 0
    sht.Application.Run(macro_name)
    sht.name = "day"+str(k_ix)
    sht=[]
def Macro(VBA,file_start,file_stop,sheet_start,sheet_stop):#range [  );对多个文件多个sheet执行特定的宏
    for i in range(file_start,file_stop):
        a="C:\\Users\\96471\\Desktop\\"+str(i)+".xlsm"
            file_path = a
        xlApp = win32com.client.DispatchEx("Excel.Application")  #打开excel操作环境
        xlApp.Visible = True    #进程可见，False是它暗自进行
        xlApp.DisplayAlerts = 0
        Book = xlApp.Workbooks.Open(file_path,False)
        for k in range(sheet_start,sheet_stop) :
            b="sheet"+str(k)
            sheet_name = b
            sht = Book.Worksheets(sheet_name)
            macro_name=VBA
             k_ix=k
            useVBA(macro_name,sht,k_ix)
        Book.Close(True)
            xlApp.quit()
            k=[]
Macro(VBA,file_start,file_stop,sheet_start,sheet_stop)
