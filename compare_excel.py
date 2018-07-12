#!/usr/bin/python
#coding:utf-8

import xlrd
import xlwt
from xlutils.copy import copy
import sys
import os

def compare_sheet(ws, src, dst):
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 5
    style = xlwt.XFStyle()
    style.pattern = pattern

    row_col_t_num_not_match = False
    if src.nrows != dst.nrows or src.ncols != dst.ncols:
        print ("sheet row/col total num is not match, src.nrows [%d] dst.nrows [%d] src.ncols [%d] dst.ncols[%d]" % (src.nrows, dst.nrows, src.ncols, dst.ncols))
        row_col_t_num_not_match = True

    max_i = src.nrows if src.nrows < dst.nrows else dst.nrows
    max_j = src.ncols if src.ncols < dst.ncols else dst.ncols

    mark = False
    for i in range(max_i):
        for j in range(max_j):
            dst_value =  dst.cell_value(i,j)
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
    return mark,row_col_t_num_not_match
  
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

    old_excel = xlrd.open_workbook("./src/" + src)
    new_excel = copy(old_excel)

    mark = False
    row_col_t_num_not_match = False
    for i in range(len(dst_sheet_names)):
        print("compare sheet [%s] with [%s]" % (src_sheet_names[i], dst_sheet_names[i]))
        src_sheet = src_book.sheet_by_name(src_sheet_names[i])
        dst_sheet = dst_book.sheet_by_name(dst_sheet_names[i])
        ws = new_excel.get_sheet(i)
        a,b = compare_sheet(ws, src_sheet, dst_sheet)

        if a == True:
            mark = True
        if b == True:
            row_col_t_num_not_match = True

    if mark == True:
        pos = src.rfind(".", 0, len(src))
        if row_col_t_num_not_match == True:
            new_excel.save("./out/" + src[0:pos] + "_marked_overlap.xls")
        else:
            new_excel.save("./out/" + src[0:pos] + "_marked.xls")
        print("find different cell, update marked excel to out dir")
    else:
        if row_col_t_num_not_match == True:
            print ("overlap cell is complete match")
        else:
            print ("all cell is complete match")

if __name__ == "__main__":

    src_file_list = sorted(os.listdir("./src/"))
    dst_file_list = sorted(os.listdir("./dst/"))

    for i in range(len(dst_file_list)):
        print("compare [%s] with [%s]!" % (src_file_list[i], dst_file_list[i]))
        compare_excel(src_file_list[i], dst_file_list[i])
        print ("\n")
