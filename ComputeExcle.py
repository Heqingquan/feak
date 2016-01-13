#!/usr/bin/env python

# -*- coding: utf-8 -*-
import xlrd
import xlwt


def read_excle(name):
    workbook = xlrd.open_workbook(name)
    sheet1 = workbook.sheet_by_index(0)
    data = {}
    for i in range(1,sheet1.nrows):
         # print(sheet1.cell(i,1).value)
        data[sheet1.cell(i,0).value] = (int(sheet1.cell(i,1).value),int(sheet1.cell(i,2).value))
    return data

def wirte_excle(name,workbook):
    w = xlwt.Workbook(encoding='utf-8')
    ws = w.add_sheet("明细")
    ws.write(0,0,"发票号")
    ws.write(0,1,"名称")
    ws.write(0,2,"单价")
    ws.write(0,3,"数量")
    ws.write(0,4,"总计")
    hight = 1
    lth = len(workbook)
    for i in range(lth):
        print(len(workbook[i][2]))
        ws.write_merge(hight,hight+len(workbook[i][2])-1,0,0,workbook[i][0])
        ws.write_merge(hight,hight+len(workbook[i][2])-1,4,4,workbook[i][1])
        for it in workbook[i][2]:
            ws.write(hight,1,it[0])
            ws.write(hight,2,it[1])
            ws.write(hight,3,it[2])
            hight += 1
    w.save(name)
