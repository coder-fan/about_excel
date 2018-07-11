#!/usr/bin/python
#coding:utf-8

import xlrd
import xlwt
from xlutils.copy import copy
import sys
import os

def compare_sheet(src_excel_name, sheet_num, src, dst):
    old_excel = xlrd.open_workbook("./src/" + src_excel_name)
    new_excel = copy(old_excel)
    ws = new_excel.get_sheet(sheet_num)
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 5
    style = xlwt.XFStyle()
    style.pattern = pattern

    mark = False
    for i in range(dst.nrows):
        for j in range(dst.ncols):
            dst_value =  dst.cell_value(i,j)
            if i >= src.nrows or j >= src.ncols:
                print ("[%d] row  [%d] col is not match, src_value is [null], dst_value is [%s]" % (i+1, j+1, dst_value))
                continue

            src_value =  src.cell_value(i,j)
            if src_value != dst_value:
                if type(src_value) == float and type(dst_value) == float:
                    a = ("%.4f" % src_value)
                    b = ("%.4f" % dst_value)
                    if a != b:
                        print ("[%d] row  [%d] col is not match, src_value is [%s], dst_value is [%s]" % (i+1, j+1, a, b))
                        ws.write(i, j, src_value, style)
                        mark = True
                else:
                    print ("[%d] row  [%d] col is not match, src_value is [%s], dst_value is [%s]" % (i+1, j+1, src_value, dst_value))
                    ws.write(i, j, src_value, style)
                    mark = True

    if mark == True:
        pos = src_excel_name.rfind(".", 0, len(src_excel_name))
        new_excel.save("./out/" + src_excel_name[0:pos] + "_marked.xls")
        print("find different cell, update marked excel to out dir")
    else:
        print ("complete match")

def compare_excel(src, dst):
    if src[0:2] != dst[0:2]:
        print ("file num is not match, src num %s dst num %s" % (src[0:2], dst[0:2]))
        return
    
    src_book = xlrd.open_workbook("./src/" + src)
    src_sheet_names = src_book.sheet_names()

    dst_book = xlrd.open_workbook("./dst/" + dst)
    dst_sheet_names = dst_book.sheet_names()
    
    if len(dst_sheet_names) != len(src_sheet_names):
        print("sheet num is not equal, src sheet num is [%d], dst sheet num is [%d]" % (len(src_sheet_names), len(dst_sheet_names)))

    for i in range(len(dst_sheet_names)):
        print("compare sheet [%s] with [%s]" % (src_sheet_names[i], dst_sheet_names[i]))
        src_sheet = src_book.sheet_by_name(src_sheet_names[i])
        dst_sheet = dst_book.sheet_by_name(dst_sheet_names[i])
        compare_sheet(src, i, src_sheet, dst_sheet)

if __name__ == "__main__":

    src_file_list = os.listdir("./src/")
    dst_file_list = os.listdir("./dst/")

    for i in range(len(dst_file_list)):
        print("compare [%s] with [%s]!" % (src_file_list[i], dst_file_list[i]))
        compare_excel(src_file_list[i], dst_file_list[i])
        print ("\n")

