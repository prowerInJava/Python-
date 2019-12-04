import xlrd
import xlwt
from time import sleep
import win32com.client as win32
import os
from os.path import getsize
import sys

def getNewestSheet(path):
      data = xlrd.open_workbook(path)
      sheet_names = data.sheet_names()
      #print(sheet_names)
      index = 0
      for i in range (len(sheet_names)):
            if '需求单' in sheet_names[i] or "需求單" in sheet_names[i]:
                  index = i
      sheet_name = sheet_names[index]
      sheet = data.sheet_by_name(sheet_name)
      return sheet

def getFiles(dirs):
      namelist = []
      dirlist = os.listdir(dirs)
      for i in range(len(dirlist)):
            if dirlist[i].endswith ('xls'):
                  namelist.append(dirlist[i])
      for i in range(len(dirlist)):
            if dirlist[i].endswith ('xlsx'):
                  namelist.append(dirlist[i])
      return namelist

def count(dirs):
      name = getFiles(dirs) #所有excel文件名list
      names = []
      mapp = {} 
      map2 = {}
      map3 = {}
      for j in range(len(name)):
            path = os.path.join(dirs,str(name[j]))
            datasize = getsize(path)
            try:
                  sheet = getNewestSheet(path)
            except Exception:
                  continue
            for i in range(6):
                  if j != "":
                        try:
                              a = sheet.cell(i,1) #B列的数据
                              if (str(a).split(":")[0] != 'empty'):
                                    opname = str(a).split(":")[1].strip("'").upper()
                                    if opname >= 'A' and opname <='Z' and len(opname)<=15:
                                          names.append(str(a).split(":")[1].strip("'").upper())
                                          map2[datasize] = str(a).split(":")[1].strip("'").upper()
                        except IndexError:
                              continue
      s = set(names)
      for i in s:
            cot = 0
            for k,v in map2.items():
                  if i.strip().upper()==v.strip().upper():
                        cot += int(k)
            mapp[i]=names.count(i)
            map3[i]=cot
      return mapp,map3

def writeExcel(dirs,path): #excel文件目录文件夹, 生成的目标结果集excel文件路径
      mapp,map3 = count(dirs)
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
      sh.SaveAs(os.path.join(path))
      ss.Close(True)

if __name__ == '__main__':
      writeExcel(r'D:\Users\Jacky_Yin\Desktop\BISON',r'D:\Users\Jacky_Yin\Desktop\Bison.xlsx')
                        
                                
      
