# -*- coding: utf-8 -*-
"""构造"改船名"测试场景:
   folder = TEST NEW NAME (新名, 用户改过的目录名)
   源 Excel 前段标题写旧名 TEST OLD NAME, 后段写新名 TEST NEW NAME
   PIC 汇总「重命名」列 = TEST OLD NAME
期望: 前段块头显示 TEST OLD NAME, 后段显示 TEST NEW NAME, 两块 VID 相同。
"""
import os, shutil, openpyxl
from datetime import datetime

BASE = r"C:\Users\culadmin\Claw-Report\.cache\renametest"
shutil.rmtree(BASE, ignore_errors=True)
src = os.path.join(BASE, "src"); os.makedirs(os.path.join(src, "TEST NEW NAME"))
HDR = ["PORT","man in","wait","Proforma","ltm eta","ltm etd","VOY. NO",
       "date","ETA","ETB","ETD","run","Port Stay(hr)","fsp distance","speed","ETA Delay/Ahead"]

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Official"
def seg(row, lane, c4, c9):
    ws.cell(row,1,lane); ws.cell(row,4,c4); ws.cell(row,9,c9)
def hdr(row):
    for i,h in enumerate(HDR,1): ws.cell(row,i,h)
def data(row, port, voy, eta, etb, etd):
    ws.cell(row,1,port); ws.cell(row,7,voy); ws.cell(row,9,eta); ws.cell(row,10,etb); ws.cell(row,11,etd)

seg(1, "AEM", "TEST OLD NAME", "TOLD"); hdr(2)
data(3, "CNSHA", "2601E", datetime(2026,10,1,8), datetime(2026,10,1,12), datetime(2026,10,2,6))
data(4, "CNNGB", "2601E", datetime(2026,10,6,8), datetime(2026,10,6,12), datetime(2026,10,7,6))
seg(6, "REX", "TEST NEW NAME", "TNMN"); hdr(7)
data(8, "CNSHA", "2602W", datetime(2026,11,20,8), datetime(2026,11,20,12), datetime(2026,11,21,6))
data(9, "CNXMN", "2602W", datetime(2026,11,25,8), datetime(2026,11,25,12), datetime(2026,11,26,6))
data(10, "MYPKG", "2602W", datetime(2026,12,1,8), datetime(2026,12,1,14), datetime(2026,12,2,6))
wb.save(os.path.join(src, "TEST NEW NAME", "TEST-2026Y.xlsx"))

pic = openpyxl.Workbook(); p = pic.active
p.append(["航线","船名(显示)","重命名","文件夹名","船代码","PIC","状态"])
p.append(["AEM","TEST NEW NAME","TEST OLD NAME","TEST NEW NAME","TNMN","Leah Liu","已有"])
pic.save(os.path.join(BASE, "pic.xlsx"))
print("built:", BASE)
