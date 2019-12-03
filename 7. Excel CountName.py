import xlrd
from time import sleep
import win32com.client as win32
import os
from os.path import getsize
import sys

name = []
dirlist = os.listdir(r'D:\Users\Desktop\STAR')  #获取star目录下所有的excel文件进行处理
for i in range(len(dirlist)):
      if dirlist[i].endswith ('xls'):
            name.append(dirlist[i])
for i in range(len(dirlist)):
      if dirlist[i].endswith ('xlsx'):
            name.append(dirlist[i])
names = []
map2 = {}
map3 = {}
#print(ops)
for j in range (len(name)):
      #print (name[i])
      data = xlrd.open_workbook(r'D:\Users\Desktop\STAR\\'+str(name[j]))
      datasize = getsize(r'D:\Users\Desktop\STAR\\'+ str(name[j])) #获取文件内存大小
      
      sheet1_name= data.sheet_names()[0]
      sheet = data.sheet_by_name(sheet1_name)
      print(name[j])
      for i in range (6):
            if i != "":
                  #print(name[j])
                  a = sheet.cell(i,1) # B列的数据
                  if (str(a).split(":")[0] !='empty'):
                        #print(str(a).split(":")[1].upper())
                        opname = str(a).split(":")[1].strip("'").upper()
                        if opname >= 'A' and opname <='Z' and len(opname)<=15:
                              names.append(str(a).split(":")[1].strip("'").upper())
                              map2[datasize] = str(a).split(":")[1].strip("'").upper()
#print(map2)
s = set(names)
mapp = {} 
for i in s:
      cot = 0
      for k,v in map2.items():
            if i.strip().upper()==v.strip().upper():
                  cot += int(k)
      mapp[i]=names.count(i)
      map3[i]=cot
#print(map3)
#写入excel文件中
app = 'Excel'
xl = win32.gencache.EnsureDispatch('%s.Application'% app)
ss = xl.Workbooks.Add() #添加一个工作簿 
sh = ss.ActiveSheet     #获取活跃的sheet
xl.Visible = True #可以在桌面可见win32 的操作
sleep(1)
for i in range(1,len(list(mapp.keys()))+1):
      sh.Cells(i,1).Value= list(mapp.keys())[i-1].upper()
      sh.Cells(i,2).Value= mapp[list(mapp.keys())[i-1]]
      sh.Cells(i,3).Value = map3[list(mapp.keys())[i-1]]
sleep(1)
#ss.Close(False)
      

      
       
