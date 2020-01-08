import xlwt #写入excel文件的库
import pandas as pd
import numpy
import nptdms
from nptdms import TdmsFile
import win32com
from win32com.client import Dispatch, constants

def import_data(file_name,sheet_name,saved_name,column_num,x,y,tdms_name):#导入数据到EXCEL



    hr_book= xlwt.Workbook(encoding='ascii')
    hr_sheet=hr_book.add_sheet(sheet_name,cell_overwrite_ok=True) #创建表格

    with open(file_name,'r+') as title: #'r+'表示对文件是进行"读取和写入的模式"
        hrtitle = title.read()
        hrtitle_list= hrtitle.split() #读取txt文件内容默认是str类型，此处将其分割成一个个元素形成列表
        i=x
        j=y
        k=0
        n = column_num
        for hl in hrtitle_list:     #此处写入excel文件
            hr_sheet.write(i,j,hl)  #i，j控制表格坐标，左定格为（0，0）
            k=k+1
            if k % n != 0:
                  j = j +1
            else :
                  j=0
                  i=i+1
    tdms_file = nptdms.TdmsFile(tdms_name)
    channel_object_Cue = tdms_file.object("CueGroup", "8")
    channel_object_TH  = tdms_file.object("SignaleGroup", "1")
    channel_object_GABA = tdms_file.object("SignaleGroup", "2")
    data_Cue = channel_object_Cue.data
    data_TH = channel_object_TH.data
    data_GABA = channel_object_GABA.data
    k=2
    for i in data_Cue:
        if i == 0 :
            k=k+1
        else:
            ix=k
            break

    for i in range(0, len(data_TH)):
        hr_sheet.write(i+1,7, data_TH[i])

    for i in range(0,len(data_GABA)):
        hr_sheet.write(i+1,8,data_GABA[i])
    style = xlwt.easyxf('pattern: pattern solid, fore_colour ice_blue')
    hr_sheet.write(ix-1, 7,data_TH[ix],style)
    hr_sheet.write(ix-1, 8,data_GABA[ix],style)
    hr_book.save(saved_name)
file_name = r''
sheet_name = "sheet1"
saved_name = r''
column_num = 3
tdms_name = r'D:\lm\待处理\钙信号原始\csds1\record\2019-12-21\2.tdms'
import_data(file_name,sheet_name,saved_name,column_num,1,0,tdms_name)
